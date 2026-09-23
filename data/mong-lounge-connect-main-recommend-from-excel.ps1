$ErrorActionPreference = "Stop"

$root = "E:\mong-lounge-site"
$master = Join-Path $root "data\몽라운지_업체자동등록_마스터_v2.xlsx"
$index = Join-Path $root "index.html"
$shopsRoot = Join-Path $root "shops"
$jsDir = Join-Path $root "assets\js"
$jsFile = Join-Path $jsDir "main-recommend-auto.js"
$stamp = Get-Date -Format "yyyyMMdd-HHmmss"
$backupRoot = "E:\mong-lounge-main-recommend-backup-$stamp"

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

if (!(Test-Path -LiteralPath $master)) { throw "마스터 엑셀 없음: $master" }
if (!(Test-Path -LiteralPath $index)) { throw "메인 index.html 없음: $index" }

# 엑셀 열림 여부 확인
try {
    $lock = [System.IO.File]::Open(
        $master,
        [System.IO.FileMode]::Open,
        [System.IO.FileAccess]::Read,
        [System.IO.FileShare]::ReadWrite
    )
    $lock.Close()
} catch {
    throw "마스터 엑셀을 읽을 수 없습니다. 저장 상태를 확인하세요."
}

New-Item -ItemType Directory -Force -Path $backupRoot | Out-Null
New-Item -ItemType Directory -Force -Path $jsDir | Out-Null
Copy-Item -LiteralPath $index -Destination (Join-Path $backupRoot "index.html") -Force
if (Test-Path -LiteralPath $jsFile) {
    Copy-Item -LiteralPath $jsFile -Destination (Join-Path $backupRoot "main-recommend-auto.js") -Force
}

function Get-ZipText {
    param([System.IO.Compression.ZipArchive]$Zip,[string]$Name)
    $e = $Zip.GetEntry($Name)
    if (!$e) { throw "XLSX 내부 파일 누락: $Name" }
    $s = $e.Open()
    try {
        $r = [System.IO.StreamReader]::new($s,[System.Text.Encoding]::UTF8,$true)
        try { return $r.ReadToEnd() } finally { $r.Dispose() }
    } finally { $s.Dispose() }
}

function Get-SharedStrings {
    param([System.IO.Compression.ZipArchive]$Zip)
    $list = New-Object System.Collections.Generic.List[string]
    if (!$Zip.GetEntry("xl/sharedStrings.xml")) { return $list }
    [xml]$xml = Get-ZipText $Zip "xl/sharedStrings.xml"
    $ns = [System.Xml.XmlNamespaceManager]::new($xml.NameTable)
    $ns.AddNamespace("m","http://schemas.openxmlformats.org/spreadsheetml/2006/main")
    foreach ($si in $xml.SelectNodes("//m:si",$ns)) {
        $parts = @($si.SelectNodes(".//m:t",$ns) | ForEach-Object { $_.InnerText })
        $list.Add(($parts -join ""))
    }
    return $list
}

function Get-CellText {
    param([System.Xml.XmlElement]$Cell,$Shared)
    if (!$Cell) { return "" }
    $type = $Cell.GetAttribute("t")
    if ($type -eq "inlineStr") {
        $parts = @($Cell.SelectNodes(".//*[local-name()='t']") | ForEach-Object { $_.InnerText })
        return ($parts -join "")
    }
    $v = $Cell.SelectSingleNode("./*[local-name()='v']")
    if (!$v) { return "" }
    $raw = $v.InnerText
    if ($type -eq "s") {
        $n = 0
        if ([int]::TryParse($raw,[ref]$n) -and $n -ge 0 -and $n -lt $Shared.Count) {
            return [string]$Shared[$n]
        }
    }
    return [string]$raw
}

function Get-SheetPath {
    param([xml]$Wb,[xml]$Rels,[string]$Name)

    $wm = [System.Xml.XmlNamespaceManager]::new($Wb.NameTable)
    $wm.AddNamespace("m","http://schemas.openxmlformats.org/spreadsheetml/2006/main")
    $wm.AddNamespace("r","http://schemas.openxmlformats.org/officeDocument/2006/relationships")

    $sheet = $Wb.SelectSingleNode("//m:sheet[@name='$Name']",$wm)
    if (!$sheet) { throw "시트를 찾을 수 없습니다: $Name" }

    $rid = $sheet.GetAttribute("id","http://schemas.openxmlformats.org/officeDocument/2006/relationships")

    $rm = [System.Xml.XmlNamespaceManager]::new($Rels.NameTable)
    $rm.AddNamespace("r","http://schemas.openxmlformats.org/package/2006/relationships")
    $rel = $Rels.SelectSingleNode("//r:Relationship[@Id='$rid']",$rm)
    if (!$rel) { throw "시트 관계 없음: $Name" }

    $target = $rel.GetAttribute("Target").Replace("\","/")
    if ($target.StartsWith("/")) { return $target.TrimStart("/") }
    return ("xl/" + $target.TrimStart("/"))
}

function Read-RecommendSheet {
    param(
        [System.IO.Compression.ZipArchive]$Zip,
        [xml]$Wb,
        [xml]$Rels,
        $Shared,
        [string]$Name,
        [string]$Kind
    )

    $path = Get-SheetPath $Wb $Rels $Name
    [xml]$sx = Get-ZipText $Zip $path
    $rows = New-Object System.Collections.Generic.List[object]

    foreach ($row in $sx.SelectNodes("//*[local-name()='row' and number(@r)>=2]")) {
        $r = [int]$row.GetAttribute("r")
        $vals = @{}
        foreach ($col in @("A","B","C","D","E","F","G","H")) {
            $cell = $sx.SelectSingleNode("//*[local-name()='c' and @r='$col$r']")
            $vals[$col] = (Get-CellText $cell $Shared).Trim()
        }

        if ($vals["A"].ToUpperInvariant() -eq "N") { continue }
        if ([string]::IsNullOrWhiteSpace($vals["B"])) { continue }

        $sort = 999
        [void][int]::TryParse($vals["H"],[ref]$sort)

        $rows.Add([pscustomobject]@{
            kind = $Kind
            shop = $vals["B"]
            region = $vals["C"]
            price = $vals["D"]
            image = $vals["E"]
            link = $vals["F"]
            desc = $vals["G"]
            sort = $sort
        })
    }

    return @($rows | Sort-Object sort,shop)
}

function Resolve-ShopLink {
    param([string]$Shop,[string]$Region,[string]$CurrentLink)

    if (![string]::IsNullOrWhiteSpace($CurrentLink)) {
        $x = $CurrentLink.Trim()
        if ($x -match '^(https?://|/)') { return $x }
        return "/" + $x.TrimStart("/")
    }

    if (!(Test-Path -LiteralPath $shopsRoot)) { return "#" }

    $candidates = @(
        Get-ChildItem -LiteralPath $shopsRoot -Recurse -File -Filter "index.html" |
        Where-Object {
            (Split-Path (Split-Path $_.FullName -Parent) -Leaf) -eq $Shop
        }
    )

    if ($candidates.Count -eq 0) { return "#" }

    $pick = $null
    if (![string]::IsNullOrWhiteSpace($Region)) {
        $escaped = [regex]::Escape($Region)
        foreach ($c in $candidates) {
            try {
                $t = Get-Content -LiteralPath $c.FullName -Raw -Encoding UTF8
                if ($t -match $escaped) { $pick = $c; break }
            } catch {}
        }
    }
    if (!$pick) { $pick = $candidates[0] }

    $rel = $pick.FullName.Substring($root.Length).Replace("\","/")
    return "/" + $rel.TrimStart("/")
}

$zip = [System.IO.Compression.ZipFile]::OpenRead($master)
try {
    [xml]$wb = Get-ZipText $zip "xl/workbook.xml"
    [xml]$rels = Get-ZipText $zip "xl/_rels/workbook.xml.rels"
    $shared = Get-SharedStrings $zip

    $vip = @(Read-RecommendSheet $zip $wb $rels $shared "메인VIP추천" "vip")
    $premium = @(Read-RecommendSheet $zip $wb $rels $shared "메인프리미엄추천" "premium")
}
finally {
    $zip.Dispose()
}

foreach ($x in @($vip + $premium)) {
    $x.link = Resolve-ShopLink $x.shop $x.region $x.link

    if ([string]::IsNullOrWhiteSpace($x.image)) {
        $x.image = if ($x.kind -eq "vip") {
            "/assets/images/main-recommend/vip-01.svg"
        } else {
            "/assets/images/main-recommend/premium-01.svg"
        }
    } elseif (!$x.image.StartsWith("/") -and !$x.image.StartsWith("http")) {
        $x.image = "/" + $x.image.TrimStart("/")
    }

    if ([string]::IsNullOrWhiteSpace($x.desc)) {
        $x.desc = "업체 정보와 이용 조건을 확인한 뒤 편하게 문의해보세요."
    }
}

if ($vip.Count -eq 0 -and $premium.Count -eq 0) {
    Write-Host ""
    Write-Host "=== 메인 추천 연결 중단 ===" -ForegroundColor Yellow
    Write-Host "메인VIP추천 / 메인프리미엄추천 시트에 업체명이 입력된 행이 없습니다."
    Write-Host "업체명을 먼저 입력한 뒤 다시 실행하세요."
    exit 0
}

$dataObj = [ordered]@{
    vip = @($vip)
    premium = @($premium)
}
$json = $dataObj | ConvertTo-Json -Depth 6 -Compress

$js = @"
window.ML_MAIN_RECOMMEND = $json;

(function(){
  const DATA = window.ML_MAIN_RECOMMEND || {vip:[], premium:[]};

  function esc(v){
    return String(v == null ? "" : v)
      .replace(/&/g,"&amp;")
      .replace(/</g,"&lt;")
      .replace(/>/g,"&gt;")
      .replace(/"/g,"&quot;");
  }

  function money(v){
    const s = String(v || "").trim();
    if(!s) return "가격 안내";
    if(/^\d+$/.test(s)) return Number(s).toLocaleString("ko-KR") + "원";
    return s;
  }

  function hideOldRecommend(){
    const fakeNames = ["강남 골드케어","송파 밸런스룸","광교 리셋테라피","부평 온기케어"];
    document.querySelectorAll("section").forEach(sec=>{
      if(sec.id === "ml-main-recommend-auto") return;
      const t = (sec.innerText || "").replace(/\s+/g," ");
      const oldVip = /VIP\s*추천/i.test(t);
      const oldPremium = /프리미엄\s*추천/.test(t);
      const fake = fakeNames.some(n=>t.includes(n));
      if(oldVip || oldPremium || fake){
        sec.style.setProperty("display","none","important");
        sec.setAttribute("data-ml-old-recommend","hidden");
      }
    });
  }

  function card(item, label){
    const link = item.link && item.link !== "#" ? item.link : "javascript:void(0)";
    const disabled = item.link && item.link !== "#" ? "" : ' aria-disabled="true"';
    return `
      <a class="ml-rec-card" href="${esc(link)}"${disabled}>
        <div class="ml-rec-image">
          <img src="${esc(item.image)}" alt="${esc(item.shop)} 추천 이미지" loading="lazy">
          <span>${esc(label)}</span>
        </div>
        <div class="ml-rec-body">
          <div class="ml-rec-region">${esc(item.region || "추천 업체")}</div>
          <h3>${esc(item.shop)}</h3>
          <p>${esc(item.desc || "")}</p>
          <div class="ml-rec-bottom">
            <b>${esc(money(item.price))}</b>
            <em>상세보기</em>
          </div>
        </div>
      </a>`;
  }

  function block(title, sub, items, label){
    if(!items || !items.length) return "";
    return `
      <section class="ml-rec-group">
        <div class="ml-rec-heading">
          <div>
            <small>${esc(label)}</small>
            <h2>${esc(title)}</h2>
            <p>${esc(sub)}</p>
          </div>
        </div>
        <div class="ml-rec-grid">
          ${items.map(x=>card(x,label)).join("")}
        </div>
      </section>`;
  }

  function ensureStyle(){
    if(document.getElementById("ml-main-recommend-style")) return;
    const style = document.createElement("style");
    style.id = "ml-main-recommend-style";
    style.textContent = `
      #ml-main-recommend-auto{margin:34px 0 42px}
      #ml-main-recommend-auto .ml-rec-group{margin:0 0 34px}
      #ml-main-recommend-auto .ml-rec-heading{display:flex;align-items:end;justify-content:space-between;margin:0 0 16px}
      #ml-main-recommend-auto .ml-rec-heading small{display:inline-block;color:#aa7a00;font-size:12px;font-weight:900;letter-spacing:.08em;margin-bottom:5px}
      #ml-main-recommend-auto .ml-rec-heading h2{margin:0;color:#163b32;font-size:27px;line-height:1.2}
      #ml-main-recommend-auto .ml-rec-heading p{margin:7px 0 0;color:#75827e;font-size:14px}
      #ml-main-recommend-auto .ml-rec-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:18px}
      #ml-main-recommend-auto .ml-rec-card{display:grid;grid-template-columns:44% 56%;min-height:190px;overflow:hidden;border:1px solid #e5e9e7;border-radius:18px;background:#fff;color:inherit;text-decoration:none;box-shadow:0 8px 26px rgba(15,53,44,.07);transition:.18s ease}
      #ml-main-recommend-auto .ml-rec-card:hover{transform:translateY(-2px);box-shadow:0 12px 30px rgba(15,53,44,.12)}
      #ml-main-recommend-auto .ml-rec-image{position:relative;min-height:190px;background:#0b4a3e;overflow:hidden}
      #ml-main-recommend-auto .ml-rec-image img{width:100%;height:100%;object-fit:cover;display:block}
      #ml-main-recommend-auto .ml-rec-image span{position:absolute;left:12px;top:12px;background:#ffc928;color:#083f35;border-radius:999px;padding:6px 10px;font-size:11px;font-weight:900}
      #ml-main-recommend-auto .ml-rec-body{padding:20px 20px 16px;display:flex;flex-direction:column}
      #ml-main-recommend-auto .ml-rec-region{font-size:12px;font-weight:800;color:#a57800;margin-bottom:5px}
      #ml-main-recommend-auto h3{margin:0 0 9px;color:#163b32;font-size:20px}
      #ml-main-recommend-auto .ml-rec-body p{margin:0;color:#6f7c78;font-size:13px;line-height:1.55;flex:1}
      #ml-main-recommend-auto .ml-rec-bottom{display:flex;align-items:center;justify-content:space-between;margin-top:14px;gap:10px}
      #ml-main-recommend-auto .ml-rec-bottom b{color:#163b32;font-size:15px}
      #ml-main-recommend-auto .ml-rec-bottom em{font-style:normal;background:#ffc928;color:#083f35;border-radius:8px;padding:7px 11px;font-size:11px;font-weight:900}
      @media(max-width:760px){
        #ml-main-recommend-auto{margin:24px 0 32px}
        #ml-main-recommend-auto .ml-rec-grid{grid-template-columns:1fr}
        #ml-main-recommend-auto .ml-rec-card{grid-template-columns:39% 61%;min-height:165px}
        #ml-main-recommend-auto .ml-rec-image{min-height:165px}
        #ml-main-recommend-auto .ml-rec-body{padding:15px}
        #ml-main-recommend-auto h3{font-size:18px}
        #ml-main-recommend-auto .ml-rec-heading h2{font-size:23px}
      }`;
    document.head.appendChild(style);
  }

  function render(){
    hideOldRecommend();
    ensureStyle();

    let root = document.getElementById("ml-main-recommend-auto");
    if(!root){
      root = document.createElement("div");
      root.id = "ml-main-recommend-auto";

      const main = document.querySelector("main") || document.body;
      const partner = main.querySelector(".partner");
      const guide = Array.from(main.querySelectorAll("section")).find(s=>{
        const t = (s.innerText || "");
        return /이용안내|지역정보/.test(t);
      });

      const anchor = guide || partner;
      if(anchor && anchor.parentNode === main) main.insertBefore(root, anchor);
      else main.appendChild(root);
    }

    root.innerHTML =
      block("VIP 추천","메인에서 별도로 관리하는 VIP 추천 업체입니다.",DATA.vip,"VIP 추천") +
      block("프리미엄 추천","메인에서 별도로 관리하는 프리미엄 추천 업체입니다.",DATA.premium,"PREMIUM");
  }

  if(document.readyState === "loading"){
    document.addEventListener("DOMContentLoaded",()=>{
      render();
      requestAnimationFrame(render);
      setTimeout(render,120);
    });
  }else{
    render();
    requestAnimationFrame(render);
    setTimeout(render,120);
  }
})();
"@

[System.IO.File]::WriteAllText($jsFile,$js,[System.Text.UTF8Encoding]::new($false))

$html = Get-Content -LiteralPath $index -Raw -Encoding UTF8

# 이전 자동 연결 태그 제거
$html = [regex]::Replace(
    $html,
    '(?is)\s*<!--\s*ML-MAIN-RECOMMEND-AUTO-START\s*-->.*?<!--\s*ML-MAIN-RECOMMEND-AUTO-END\s*-->\s*',
    "`r`n"
)

$tag = @"
<!-- ML-MAIN-RECOMMEND-AUTO-START -->
<script src="assets/js/main-recommend-auto.js?v=$stamp"></script>
<!-- ML-MAIN-RECOMMEND-AUTO-END -->
"@

if ($html -match '(?is)</body>') {
    $html = [regex]::Replace($html,'(?is)</body>',"$tag`r`n</body>",1)
} else {
    $html += "`r`n$tag`r`n"
}

Set-Content -LiteralPath $index -Value $html -Encoding UTF8

# 검증
$checkIndex = Get-Content -LiteralPath $index -Raw -Encoding UTF8
$checkJs = Get-Content -LiteralPath $jsFile -Raw -Encoding UTF8

Write-Host ""
Write-Host "=== 메인 추천업체 엑셀 자동 연결 완료 ==="
Write-Host "VIP 업체:" $vip.Count
Write-Host "프리미엄 업체:" $premium.Count
Write-Host "메인 JS 생성:" (Test-Path -LiteralPath $jsFile)
Write-Host "index 연결 확인:" ($checkIndex -match 'ML-MAIN-RECOMMEND-AUTO-START')
Write-Host "기존 임시 추천 숨김 코드:" ($checkJs -match '강남 골드케어')
Write-Host ""
Write-Host "엑셀 시트: 메인VIP추천 / 메인프리미엄추천"
Write-Host "상세페이지 링크: 빈 칸이면 업체명+지역 기준 자동 탐색"
Write-Host "이미지/소개문구/가격/지역: 엑셀 값 사용"
Write-Host "백업 폴더:" $backupRoot
Write-Host ""
Write-Host "브라우저에서 Ctrl+F5로 확인하세요."
