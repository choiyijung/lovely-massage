$ErrorActionPreference = "Stop"

$root      = "E:\mong-lounge-site"
$dataDir   = Join-Path $root "data"
$master    = Join-Path $dataDir "몽라운지_업체자동등록_마스터_v2.xlsx"
$index     = Join-Path $root "index.html"
$shopsRoot = Join-Path $root "shops"
$imageDir  = Join-Path $root "assets\images\main-recommend"
$finalJs   = Join-Path $root "assets\js\main-recommend-final.js"

$stamp = Get-Date -Format "yyyyMMdd-HHmmss"
$backupRoot = "E:\mong-lounge-main-recommend-source-fix-backup-$stamp"
$tempXlsx = Join-Path $env:TEMP "mong_lounge_main_recommend_$stamp.xlsx"

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

if (!(Test-Path -LiteralPath $master)) { throw "마스터 엑셀 없음: $master" }
if (!(Test-Path -LiteralPath $index))  { throw "메인 index.html 없음: $index" }

# 엑셀을 열어 둔 상태면 안전하게 중단
try {
    $lock = [System.IO.File]::Open(
        $master,
        [System.IO.FileMode]::Open,
        [System.IO.FileAccess]::ReadWrite,
        [System.IO.FileShare]::None
    )
    $lock.Close()
}
catch {
    Write-Host ""
    Write-Host "=== 작업 중단 ===" -ForegroundColor Yellow
    Write-Host "몽라운지_업체자동등록_마스터_v2.xlsx 파일을 닫고 다시 실행하세요."
    exit 1
}

New-Item -ItemType Directory -Force -Path $backupRoot,$imageDir,(Split-Path $finalJs -Parent) | Out-Null
Copy-Item -LiteralPath $master -Destination (Join-Path $backupRoot "몽라운지_업체자동등록_마스터_v2.xlsx") -Force
Copy-Item -LiteralPath $index  -Destination (Join-Path $backupRoot "index.html") -Force
if (Test-Path -LiteralPath $finalJs) {
    Copy-Item -LiteralPath $finalJs -Destination (Join-Path $backupRoot "main-recommend-final.js") -Force
}
Copy-Item -LiteralPath $master -Destination $tempXlsx -Force

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

function Set-ZipText {
    param([System.IO.Compression.ZipArchive]$Zip,[string]$Name,[string]$Text)
    $old = $Zip.GetEntry($Name)
    if ($old) { $old.Delete() }
    $e = $Zip.CreateEntry($Name,[System.IO.Compression.CompressionLevel]::Optimal)
    $s = $e.Open()
    try {
        $w = [System.IO.StreamWriter]::new($s,[System.Text.UTF8Encoding]::new($false))
        try { $w.Write($Text); $w.Flush() } finally { $w.Dispose() }
    } finally { $s.Dispose() }
}

function Get-SharedStrings {
    param([System.IO.Compression.ZipArchive]$Zip)
    $out = New-Object System.Collections.Generic.List[string]
    if (!$Zip.GetEntry("xl/sharedStrings.xml")) { return $out }
    [xml]$x = Get-ZipText $Zip "xl/sharedStrings.xml"
    foreach ($si in $x.SelectNodes("//*[local-name()='si']")) {
        $parts = @($si.SelectNodes(".//*[local-name()='t']") | ForEach-Object { $_.InnerText })
        $out.Add(($parts -join ""))
    }
    return $out
}

function Get-CellText {
    param([System.Xml.XmlElement]$Cell,$Shared)
    if (!$Cell) { return "" }

    $t = $Cell.GetAttribute("t")
    if ($t -eq "inlineStr") {
        return ((@($Cell.SelectNodes(".//*[local-name()='t']") | ForEach-Object { $_.InnerText })) -join "")
    }

    $v = $Cell.SelectSingleNode("./*[local-name()='v']")
    if (!$v) { return "" }

    if ($t -eq "s") {
        $i = 0
        if ([int]::TryParse($v.InnerText,[ref]$i) -and $i -ge 0 -and $i -lt $Shared.Count) {
            return [string]$Shared[$i]
        }
    }
    return [string]$v.InnerText
}

function Get-SheetPath {
    param([xml]$Wb,[xml]$Rels,[string]$SheetName)

    $sheet = $Wb.SelectSingleNode("//*[local-name()='sheet' and @name='$SheetName']")
    if (!$sheet) { throw "시트 없음: $SheetName" }

    $rid = $sheet.GetAttribute(
        "id",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    )
    $rel = $Rels.SelectSingleNode("//*[local-name()='Relationship' and @Id='$rid']")
    if (!$rel) { throw "시트 관계 없음: $SheetName" }

    $target = $rel.GetAttribute("Target").Replace("\","/")
    if ($target.StartsWith("/")) { return $target.TrimStart("/") }
    return "xl/" + $target.TrimStart("/")
}

function Read-Rows {
    param([xml]$Sheet,$Shared,[string[]]$Cols)
    $list = @()
    foreach ($row in $Sheet.SelectNodes("//*[local-name()='row' and number(@r)>=2]")) {
        $r = [int]$row.GetAttribute("r")
        $o = [ordered]@{ Row = $r }
        foreach ($col in $Cols) {
            $cell = $Sheet.SelectSingleNode("//*[local-name()='c' and @r='$col$r']")
            $o[$col] = (Get-CellText $cell $Shared).Trim()
        }
        $list += [pscustomobject]$o
    }
    return $list
}

function Set-InlineCell {
    param([xml]$Sheet,[int]$Row,[string]$Col,[string]$Text)

    $ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    $sheetData = $Sheet.SelectSingleNode("/*[local-name()='worksheet']/*[local-name()='sheetData']")

    $rowNode = $Sheet.SelectSingleNode("//*[local-name()='row' and @r='$Row']")
    if (!$rowNode) {
        $rowNode = $Sheet.CreateElement("row",$ns)
        [void]$rowNode.SetAttribute("r",[string]$Row)
        [void]$sheetData.AppendChild($rowNode)
    }

    $ref = "$Col$Row"
    $cell = $Sheet.SelectSingleNode("//*[local-name()='c' and @r='$ref']")
    if (!$cell) {
        $cell = $Sheet.CreateElement("c",$ns)
        [void]$cell.SetAttribute("r",$ref)
        [void]$rowNode.AppendChild($cell)
    }

    while ($cell.HasChildNodes) { [void]$cell.RemoveChild($cell.FirstChild) }
    [void]$cell.SetAttribute("t","inlineStr")

    $is = $Sheet.CreateElement("is",$ns)
    $t  = $Sheet.CreateElement("t",$ns)
    $t.InnerText = $Text
    [void]$is.AppendChild($t)
    [void]$cell.AppendChild($is)
}

function Min-Price {
    param($CourseRows,[string]$Shop)
    $nums = @()
    foreach ($r in ($CourseRows | Where-Object { $_.B -eq $Shop -and $_.A.ToUpperInvariant() -ne "N" })) {
        $nText = ($r.E -replace '[^\d]','')
        if ($nText) {
            $n = 0
            if ([int]::TryParse($nText,[ref]$n) -and $n -gt 0) { $nums += $n }
        }
    }
    if ($nums.Count -eq 0) { return "가격 문의" }
    $min = ($nums | Measure-Object -Minimum).Minimum
    return ("{0:N0}원~" -f $min)
}

function Region-Text {
    param($AreaRows,[string]$Shop)
    $vals = @(
        $AreaRows |
        Where-Object { $_.B -eq $Shop -and $_.A.ToUpperInvariant() -ne "N" -and $_.C } |
        Select-Object -ExpandProperty C -Unique |
        Select-Object -First 3
    )
    if ($vals.Count -eq 0) { return "추천 지역" }
    return ($vals -join " · ")
}

function Resolve-ShopLink {
    param([string]$Shop,[string]$Region)

    if (!(Test-Path -LiteralPath $shopsRoot)) { return "#" }

    $cands = @(
        Get-ChildItem -LiteralPath $shopsRoot -Recurse -File -Filter "index.html" |
        Where-Object {
            (Split-Path (Split-Path $_.FullName -Parent) -Leaf) -eq $Shop
        }
    )
    if ($cands.Count -eq 0) { return "#" }

    $pick = $null
    $regionParts = @($Region -split '[·,/ ]+' | Where-Object { $_.Length -ge 2 })

    foreach ($c in $cands) {
        try {
            $txt = Get-Content -LiteralPath $c.FullName -Raw -Encoding UTF8
            if ($regionParts | Where-Object { $txt -match [regex]::Escape($_) }) {
                $pick = $c
                break
            }
        } catch {}
    }
    if (!$pick) { $pick = $cands[0] }

    return "/" + $pick.FullName.Substring($root.Length).TrimStart("\").Replace("\","/")
}

function New-RecommendSvg {
    param(
        [string]$Path,
        [string]$Badge,
        [string]$Shop,
        [string]$Region,
        [string]$C1,
        [string]$C2
    )

    $badgeE  = [System.Net.WebUtility]::HtmlEncode($Badge)
    $shopE   = [System.Net.WebUtility]::HtmlEncode($Shop)
    $regionE = [System.Net.WebUtility]::HtmlEncode($Region)

    $svg = @"
<svg xmlns="http://www.w3.org/2000/svg" width="900" height="520" viewBox="0 0 900 520">
  <defs>
    <linearGradient id="g" x1="0" y1="0" x2="1" y2="1">
      <stop offset="0%" stop-color="$C1"/>
      <stop offset="100%" stop-color="$C2"/>
    </linearGradient>
  </defs>
  <rect width="900" height="520" fill="url(#g)"/>
  <circle cx="760" cy="90" r="180" fill="#fff" opacity=".07"/>
  <circle cx="110" cy="450" r="220" fill="#fff" opacity=".05"/>
  <rect x="56" y="56" width="788" height="408" rx="34" fill="#061e19" opacity=".72"/>
  <rect x="90" y="92" width="150" height="44" rx="22" fill="#ffc928"/>
  <text x="165" y="121" text-anchor="middle" font-family="Arial,sans-serif" font-size="20" font-weight="800" fill="#083f35">$badgeE</text>
  <text x="90" y="250" font-family="Arial,sans-serif" font-size="48" font-weight="800" fill="#fff">$shopE</text>
  <text x="92" y="310" font-family="Arial,sans-serif" font-size="24" fill="#ccefe6">$regionE</text>
  <line x1="92" y1="355" x2="805" y2="355" stroke="#ffc928" stroke-width="2"/>
  <text x="92" y="405" font-family="Arial,sans-serif" font-size="18" font-weight="700" fill="#fff" opacity=".85">MONG LOUNGE RECOMMEND</text>
</svg>
"@
    [System.IO.File]::WriteAllText($Path,$svg,[System.Text.UTF8Encoding]::new($false))
}

$zip = [System.IO.Compression.ZipFile]::Open(
    $tempXlsx,
    [System.IO.Compression.ZipArchiveMode]::Update
)

try {
    [xml]$wb   = Get-ZipText $zip "xl/workbook.xml"
    [xml]$rels = Get-ZipText $zip "xl/_rels/workbook.xml.rels"
    $shared    = Get-SharedStrings $zip

    $basicPath = Get-SheetPath $wb $rels "업체기본정보"
    $coursePath= Get-SheetPath $wb $rels "코스정보"
    $areaPath  = Get-SheetPath $wb $rels "적용지역"
    $vipPath   = Get-SheetPath $wb $rels "메인VIP추천"
    $prePath   = Get-SheetPath $wb $rels "메인프리미엄추천"

    [xml]$basicXml  = Get-ZipText $zip $basicPath
    [xml]$courseXml = Get-ZipText $zip $coursePath
    [xml]$areaXml   = Get-ZipText $zip $areaPath
    [xml]$vipXml    = Get-ZipText $zip $vipPath
    [xml]$preXml    = Get-ZipText $zip $prePath

    $basicRows  = @(Read-Rows $basicXml  $shared @("A","B","C","D"))
    $courseRows = @(Read-Rows $courseXml $shared @("A","B","C","D","E"))
    $areaRows   = @(Read-Rows $areaXml   $shared @("A","B","C"))
    $vipRows    = @(Read-Rows $vipXml    $shared @("A","B","C","D","E","F","G","H"))
    $preRows    = @(Read-Rows $preXml    $shared @("A","B","C","D","E","F","G","H"))

    $activeShops = @(
        $basicRows |
        Where-Object { $_.B -and $_.A.ToUpperInvariant() -ne "N" } |
        Select-Object -ExpandProperty B -Unique
    )

    if ($activeShops.Count -lt 8) {
        throw "활성 업체가 8개 미만입니다. 현재 활성 업체: $($activeShops.Count)"
    }

    $used = New-Object System.Collections.Generic.HashSet[string]

    function Fill-SheetRows {
        param(
            [xml]$SheetXml,
            $ExistingRows,
            [string]$Kind,
            [int]$Offset,
            [string]$Prefix
        )

        $result = @()

        for ($slot=1; $slot -le 4; $slot++) {
            $rowNum = $slot + 1
            $existing = $ExistingRows | Where-Object { $_.Row -eq $rowNum } | Select-Object -First 1
            $shop = if ($existing) { $existing.B } else { "" }

            # 이미 입력된 추천업체는 그대로 사용, 없으면 실제 등록업체에서 자동 보충
            if ([string]::IsNullOrWhiteSpace($shop) -or !($activeShops -contains $shop)) {
                $shop = $activeShops | Where-Object { !$used.Contains($_) } | Select-Object -First 1
            }
            if (!$shop) { throw "$Kind 추천업체를 채울 수 없습니다." }
            [void]$used.Add($shop)

            $region = if ($existing -and $existing.C) { $existing.C } else { Region-Text $areaRows $shop }
            $price  = if ($existing -and $existing.D) { $existing.D } else { Min-Price $courseRows $shop }
            $link   = if ($existing -and $existing.F) { $existing.F } else { Resolve-ShopLink $shop $region }

            $desc = if ($existing -and $existing.G) {
                $existing.G
            } elseif ($Kind -eq "VIP") {
                "$shop 업체의 지역과 코스 정보를 편리하게 확인할 수 있는 VIP 추천입니다. 상세페이지에서 이용 조건과 안내를 확인해보세요."
            } else {
                "$shop 업체의 이용 정보와 코스를 한눈에 비교할 수 있는 프리미엄 추천입니다. 원하는 일정에 맞춰 상세 안내를 확인해보세요."
            }

            $imgName = "$Prefix-{0:00}.svg" -f $slot
            $imgWeb  = "/assets/images/main-recommend/$imgName"
            $imgDisk = Join-Path $imageDir $imgName

            $colors = @(
                @("#0b594b","#07352e"),
                @("#6d4a1c","#2e2012"),
                @("#49315d","#21162c"),
                @("#27566a","#112d37"),
                @("#0c5145","#072c26"),
                @("#5e3b20","#2d1d12"),
                @("#43315d","#21182d"),
                @("#3f5961","#1f3036")
            )
            $pair = $colors[$Offset + ($slot-1)]
            New-RecommendSvg $imgDisk $Kind $shop $region $pair[0] $pair[1]

            Set-InlineCell $SheetXml $rowNum "A" "Y"
            Set-InlineCell $SheetXml $rowNum "B" $shop
            Set-InlineCell $SheetXml $rowNum "C" $region
            Set-InlineCell $SheetXml $rowNum "D" $price
            Set-InlineCell $SheetXml $rowNum "E" $imgWeb
            Set-InlineCell $SheetXml $rowNum "F" $link
            Set-InlineCell $SheetXml $rowNum "G" $desc
            Set-InlineCell $SheetXml $rowNum "H" ([string]$slot)

            $result += [pscustomobject]@{
                shop=$shop; region=$region; price=$price; image=$imgWeb;
                link=$link; desc=$desc; sort=$slot
            }
        }

        return $result
    }

    $vipItems = @(Fill-SheetRows $vipXml $vipRows "VIP" 0 "vip")
    $preItems = @(Fill-SheetRows $preXml $preRows "PREMIUM" 4 "premium")

    Set-ZipText $zip $vipPath $vipXml.OuterXml
    Set-ZipText $zip $prePath $preXml.OuterXml
}
finally {
    $zip.Dispose()
}

Copy-Item -LiteralPath $tempXlsx -Destination $master -Force
Remove-Item -LiteralPath $tempXlsx -Force -ErrorAction SilentlyContinue

$data = [ordered]@{
    vip = @($vipItems)
    premium = @($preItems)
}
$json = $data | ConvertTo-Json -Depth 6 -Compress

# 딱 하나의 추천 JS만 사용. 이전 임시 업체명은 이 파일에 아예 넣지 않음.
$js = @"
window.ML_MAIN_RECOMMEND = $json;

(function(){
  const D = window.ML_MAIN_RECOMMEND || {vip:[],premium:[]};
  const esc = s => String(s ?? "")
    .replace(/&/g,"&amp;").replace(/</g,"&lt;")
    .replace(/>/g,"&gt;").replace(/"/g,"&quot;");

  function card(x,badge){
    return `
      <a class="mlr-card" href="${esc(x.link || "#")}">
        <div class="mlr-thumb">
          <img src="${esc(x.image)}" alt="${esc(x.shop)}">
          <span>${esc(badge)}</span>
          <b>${esc(x.region)}</b>
        </div>
        <div class="mlr-body">
          <h3>${esc(x.shop)}</h3>
          <p>${esc(x.desc)}</p>
          <div class="mlr-foot">
            <strong>${esc(x.price)}</strong>
            <em>상세보기</em>
          </div>
        </div>
      </a>`;
  }

  function group(title,badge,items){
    if(!items || !items.length) return "";
    return `
      <section class="mlr-group">
        <h2>${esc(title)}</h2>
        <div class="mlr-grid">${items.map(x=>card(x,badge)).join("")}</div>
      </section>`;
  }

  function style(){
    if(document.getElementById("mlr-style")) return;
    const s=document.createElement("style");
    s.id="mlr-style";
    s.textContent=`
      #ml-main-recommend-final{margin:28px 0 38px}
      #ml-main-recommend-final .mlr-group{margin:0 0 34px}
      #ml-main-recommend-final .mlr-group h2{margin:0 0 14px;font-size:22px;color:#111}
      #ml-main-recommend-final .mlr-grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:10px}
      #ml-main-recommend-final .mlr-card{overflow:hidden;border:1px solid #ddd;border-radius:14px;background:#fff;text-decoration:none;color:#111;box-shadow:0 2px 8px rgba(0,0,0,.05)}
      #ml-main-recommend-final .mlr-thumb{height:120px;position:relative;overflow:hidden;background:#174f43}
      #ml-main-recommend-final .mlr-thumb img{width:100%;height:100%;object-fit:cover;display:block}
      #ml-main-recommend-final .mlr-thumb span{position:absolute;left:0;top:0;background:#efc342;color:#111;padding:14px 20px;border-radius:0 0 14px 0;font-size:11px;font-weight:900}
      #ml-main-recommend-final .mlr-thumb b{position:absolute;left:12px;bottom:10px;color:#fff;font-size:12px;text-shadow:0 1px 3px #000}
      #ml-main-recommend-final .mlr-body{padding:12px}
      #ml-main-recommend-final .mlr-body h3{margin:0 0 6px;font-size:15px}
      #ml-main-recommend-final .mlr-body p{height:38px;overflow:hidden;margin:0 0 14px;color:#777;font-size:11px;line-height:1.55}
      #ml-main-recommend-final .mlr-foot{display:flex;align-items:center;justify-content:space-between;gap:8px}
      #ml-main-recommend-final .mlr-foot strong{font-size:14px}
      #ml-main-recommend-final .mlr-foot em{font-style:normal;border:1px solid #e5ad32;border-radius:8px;padding:6px 9px;font-size:10px;font-weight:800}
      @media(max-width:760px){#ml-main-recommend-final .mlr-grid{grid-template-columns:repeat(2,minmax(0,1fr))}}
    `;
    document.head.appendChild(s);
  }

  function removePrevious(){
    document.getElementById("ml-main-recommend-auto")?.remove();
    document.getElementById("ml-main-recommend-final")?.remove();
  }

  function render(){
    removePrevious();
    style();

    const root=document.createElement("div");
    root.id="ml-main-recommend-final";
    root.innerHTML=
      group("VIP 추천","VIP",D.vip)+
      group("프리미엄 추천","PICK",D.premium);

    const main=document.querySelector("main") || document.body;
    const partner=main.querySelector(".partner");

    if(partner) main.insertBefore(root,partner);
    else main.appendChild(root);
  }

  if(document.readyState==="loading"){
    document.addEventListener("DOMContentLoaded",render);
  }else{
    render();
  }
})();
"@

[System.IO.File]::WriteAllText($finalJs,$js,[System.Text.UTF8Encoding]::new($false))

# index.html에서 기존 추천용 JS 3종 연결을 전부 제거하고 최종 JS 하나만 연결
$html = Get-Content -LiteralPath $index -Raw -Encoding UTF8

$html = [regex]::Replace(
    $html,
    '(?is)\s*<!--\s*ML-MAIN-RECOMMEND-AUTO-START\s*-->.*?<!--\s*ML-MAIN-RECOMMEND-AUTO-END\s*-->\s*',
    "`r`n"
)
$html = [regex]::Replace(
    $html,
    '(?is)\s*<!--\s*ML-MAIN-RECOMMEND-CLEANUP-START\s*-->.*?<!--\s*ML-MAIN-RECOMMEND-CLEANUP-END\s*-->\s*',
    "`r`n"
)
$html = [regex]::Replace(
    $html,
    '(?is)\s*<!--\s*ML-MAIN-RECOMMEND-FINAL-START\s*-->.*?<!--\s*ML-MAIN-RECOMMEND-FINAL-END\s*-->\s*',
    "`r`n"
)

# 마커 없이 들어간 추천 JS 태그도 제거
$html = [regex]::Replace(
    $html,
    '(?is)\s*<script\b[^>]*src=["'']assets/js/main-recommend-(?:auto|cleanup|final)\.js[^"'']*["''][^>]*>\s*</script>\s*',
    "`r`n"
)

$tag = @"
<!-- ML-MAIN-RECOMMEND-FINAL-START -->
<script src="assets/js/main-recommend-final.js?v=$stamp"></script>
<!-- ML-MAIN-RECOMMEND-FINAL-END -->
"@

$html = [regex]::Replace($html,'(?is)</body>',"$tag`r`n</body>",1)
Set-Content -LiteralPath $index -Value $html -Encoding UTF8

# 최종 검사
$indexCheck = Get-Content -LiteralPath $index -Raw -Encoding UTF8
$jsCheck    = Get-Content -LiteralPath $finalJs -Raw -Encoding UTF8

$oldLinked = $indexCheck -match 'main-recommend-auto\.js|main-recommend-cleanup\.js'
$oldNames  = $jsCheck -match '강남 골드케어|송파 밸런스룸|광교 리셋테라피|부평 온기케어'

Write-Host ""
Write-Host "=== 메인 추천 원본 직접 교체 완료 ==="
Write-Host "VIP 추천:"
$vipItems | ForEach-Object { Write-Host " -" $_.shop "|" $_.region "|" $_.price }
Write-Host "프리미엄 추천:"
$preItems | ForEach-Object { Write-Host " -" $_.shop "|" $_.region "|" $_.price }

Write-Host ""
Write-Host "기존 auto/cleanup JS 연결 남음:" $oldLinked
Write-Host "최종 JS에 예전 임시업체명 남음:" $oldNames
Write-Host "최종 JS 연결 확인:" ($indexCheck -match 'main-recommend-final\.js')
Write-Host "추천 이미지: 8개"
Write-Host "백업 폴더:" $backupRoot
Write-Host ""
Write-Host "정상이면 위 두 항목이 모두 False 입니다."
Write-Host "브라우저에서 Ctrl+F5 하세요."
