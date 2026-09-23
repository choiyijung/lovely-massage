$ErrorActionPreference = "Stop"

$root = "E:\mong-lounge-site"
$dataDir = Join-Path $root "data"
$master = Join-Path $dataDir "몽라운지_업체자동등록_마스터_v2.xlsx"
$index = Join-Path $root "index.html"
$shopsRoot = Join-Path $root "shops"
$imageDir = Join-Path $root "assets\images\main-recommend"
$jsFile = Join-Path $root "assets\js\main-recommend-final.js"
$stamp = Get-Date -Format "yyyyMMdd-HHmmss"
$backupRoot = "E:\mong-lounge-main-recommend-final-backup-$stamp"
$temp = Join-Path $env:TEMP "mong_lounge_main_final_$stamp.xlsx"

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

if (!(Test-Path -LiteralPath $master)) { throw "마스터 엑셀을 찾을 수 없습니다: $master" }
if (!(Test-Path -LiteralPath $index)) { throw "메인 index.html을 찾을 수 없습니다: $index" }

# 엑셀 열려 있으면 중단
try {
    $lock = [System.IO.File]::Open($master,[System.IO.FileMode]::Open,[System.IO.FileAccess]::ReadWrite,[System.IO.FileShare]::None)
    $lock.Close()
}
catch {
    Write-Host ""
    Write-Host "=== 작업 중단 ===" -ForegroundColor Yellow
    Write-Host "몽라운지_업체자동등록_마스터_v2.xlsx 파일을 닫은 뒤 다시 실행하세요."
    exit 1
}

New-Item -ItemType Directory -Force -Path $backupRoot,$imageDir,(Split-Path $jsFile -Parent) | Out-Null
Copy-Item -LiteralPath $master -Destination (Join-Path $backupRoot "몽라운지_업체자동등록_마스터_v2.xlsx") -Force
Copy-Item -LiteralPath $index -Destination (Join-Path $backupRoot "index.html") -Force
if (Test-Path -LiteralPath $jsFile) {
    Copy-Item -LiteralPath $jsFile -Destination (Join-Path $backupRoot "main-recommend-final.js") -Force
}
Copy-Item -LiteralPath $master -Destination $temp -Force

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
    $list = New-Object System.Collections.Generic.List[string]
    if (!$Zip.GetEntry("xl/sharedStrings.xml")) { return $list }
    [xml]$x = Get-ZipText $Zip "xl/sharedStrings.xml"
    foreach ($si in $x.SelectNodes("//*[local-name()='si']")) {
        $parts = @($si.SelectNodes(".//*[local-name()='t']") | ForEach-Object { $_.InnerText })
        $list.Add(($parts -join ""))
    }
    return $list
}

function Get-CellText {
    param([System.Xml.XmlElement]$Cell,$Shared)
    if (!$Cell) { return "" }
    $t = $Cell.GetAttribute("t")
    if ($t -eq "inlineStr") {
        return ((@($Cell.SelectNodes(".//*[local-name()='t']") | ForEach-Object{$_.InnerText})) -join "")
    }
    $v = $Cell.SelectSingleNode("./*[local-name()='v']")
    if (!$v) { return "" }
    if ($t -eq "s") {
        $i=0
        if ([int]::TryParse($v.InnerText,[ref]$i) -and $i -ge 0 -and $i -lt $Shared.Count) { return [string]$Shared[$i] }
    }
    return [string]$v.InnerText
}

function Get-SheetPath {
    param([xml]$Wb,[xml]$Rels,[string]$Name)
    $sheet = $Wb.SelectSingleNode("//*[local-name()='sheet' and @name='$Name']")
    if (!$sheet) { throw "시트 없음: $Name" }
    $rid = $sheet.GetAttribute("id","http://schemas.openxmlformats.org/officeDocument/2006/relationships")
    $rel = $Rels.SelectSingleNode("//*[local-name()='Relationship' and @Id='$rid']")
    if (!$rel) { throw "시트 관계 없음: $Name" }
    $target = $rel.GetAttribute("Target").Replace("\","/")
    if ($target.StartsWith("/")) { return $target.TrimStart("/") }
    return "xl/" + $target.TrimStart("/")
}

function Read-Table {
    param([xml]$Sheet,$Shared,[string[]]$Cols)
    $rows = @()
    foreach ($row in $Sheet.SelectNodes("//*[local-name()='row' and number(@r)>=2]")) {
        $r = [int]$row.GetAttribute("r")
        $o = [ordered]@{}
        foreach ($col in $Cols) {
            $c = $Sheet.SelectSingleNode("//*[local-name()='c' and @r='$col$r']")
            $o[$col] = (Get-CellText $c $Shared).Trim()
        }
        $rows += [pscustomobject]$o
    }
    return $rows
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
    $t = $Sheet.CreateElement("t",$ns)
    $t.InnerText = $Text
    [void]$is.AppendChild($t)
    [void]$cell.AppendChild($is)
}

function Format-Price([string[]]$values) {
    $nums = @()
    foreach ($v in $values) {
        if ([string]::IsNullOrWhiteSpace($v)) { continue }
        $ntext = ($v -replace '[^\d]','')
        if ($ntext) {
            $n = 0
            if ([int]::TryParse($ntext,[ref]$n) -and $n -gt 0) { $nums += $n }
        }
    }
    if ($nums.Count -eq 0) { return "가격 문의" }
    $min = ($nums | Measure-Object -Minimum).Minimum
    return ("{0:N0}원~" -f $min)
}

function Resolve-Link([string]$shop,[string]$region) {
    if (!(Test-Path -LiteralPath $shopsRoot)) { return "#" }
    $cands = @(Get-ChildItem -LiteralPath $shopsRoot -Recurse -File -Filter "index.html" | Where-Object {
        (Split-Path (Split-Path $_.FullName -Parent) -Leaf) -eq $shop
    })
    if ($cands.Count -eq 0) { return "#" }

    $pick = $null
    if ($region) {
        foreach ($c in $cands) {
            try {
                $txt = Get-Content -LiteralPath $c.FullName -Raw -Encoding UTF8
                $parts = @($region -split '[·,/ ]+' | Where-Object{ $_.Length -ge 2 })
                if ($parts | Where-Object{$txt -match [regex]::Escape($_)}) { $pick=$c; break }
            } catch {}
        }
    }
    if (!$pick) { $pick=$cands[0] }
    return "/" + $pick.FullName.Substring($root.Length).TrimStart("\").Replace("\","/")
}

function Write-Svg([string]$path,[string]$badge,[string]$shop,[string]$region,[string]$c1,[string]$c2) {
    $shopEsc=[System.Net.WebUtility]::HtmlEncode($shop)
    $regionEsc=[System.Net.WebUtility]::HtmlEncode($region)
    $badgeEsc=[System.Net.WebUtility]::HtmlEncode($badge)
    $svg=@"
<svg xmlns="http://www.w3.org/2000/svg" width="900" height="520" viewBox="0 0 900 520">
<defs><linearGradient id="g" x1="0" y1="0" x2="1" y2="1"><stop offset="0%" stop-color="$c1"/><stop offset="100%" stop-color="$c2"/></linearGradient></defs>
<rect width="900" height="520" fill="url(#g)"/>
<circle cx="760" cy="100" r="190" fill="#fff" opacity=".07"/><circle cx="110" cy="450" r="220" fill="#fff" opacity=".05"/>
<rect x="55" y="55" width="790" height="410" rx="34" fill="#071e1a" opacity=".73"/>
<rect x="88" y="90" width="150" height="44" rx="22" fill="#ffc928"/>
<text x="163" y="119" text-anchor="middle" font-family="Arial,sans-serif" font-size="20" font-weight="800" fill="#083f35">$badgeEsc</text>
<text x="88" y="250" font-family="Arial,sans-serif" font-size="48" font-weight="800" fill="#fff">$shopEsc</text>
<text x="90" y="310" font-family="Arial,sans-serif" font-size="24" fill="#c9f1e5">$regionEsc</text>
<line x1="90" y1="355" x2="805" y2="355" stroke="#ffc928" stroke-width="2"/>
<text x="90" y="405" font-family="Arial,sans-serif" font-size="18" font-weight="700" fill="#fff" opacity=".85">MONG LOUNGE RECOMMEND</text>
</svg>
"@
    [System.IO.File]::WriteAllText($path,$svg,[System.Text.UTF8Encoding]::new($false))
}

$zip = [System.IO.Compression.ZipFile]::Open($temp,[System.IO.Compression.ZipArchiveMode]::Update)
try {
    [xml]$wb = Get-ZipText $zip "xl/workbook.xml"
    [xml]$rels = Get-ZipText $zip "xl/_rels/workbook.xml.rels"
    $shared = Get-SharedStrings $zip

    $basicPath = Get-SheetPath $wb $rels "업체기본정보"
    $coursePath = Get-SheetPath $wb $rels "코스정보"
    $areaPath = Get-SheetPath $wb $rels "적용지역"
    $vipPath = Get-SheetPath $wb $rels "메인VIP추천"
    $prePath = Get-SheetPath $wb $rels "메인프리미엄추천"

    [xml]$basicXml = Get-ZipText $zip $basicPath
    [xml]$courseXml = Get-ZipText $zip $coursePath
    [xml]$areaXml = Get-ZipText $zip $areaPath
    [xml]$vipXml = Get-ZipText $zip $vipPath
    [xml]$preXml = Get-ZipText $zip $prePath

    $basic = @(Read-Table $basicXml $shared @("A","B","C","D"))
    $courses = @(Read-Table $courseXml $shared @("A","B","C","D","E"))
    $areas = @(Read-Table $areaXml $shared @("A","B","C"))

    $active = @($basic | Where-Object{ $_.B -and $_.A.ToUpperInvariant() -ne "N" })
    if ($active.Count -lt 8) { throw "활성 업체가 8개 미만입니다. 현재: $($active.Count)" }

    $preferred = @(
        "한국미인테라피","너무이쁜테라피","love테라피","프리미엄테라피",
        "마인드좋은테라피","S썸바디테라피","24시미녀테라피","레드테라피"
    )
    $selected = New-Object System.Collections.Generic.List[string]
    foreach ($name in $preferred) {
        if ($active.B -contains $name -and !$selected.Contains($name)) { $selected.Add($name) }
    }
    foreach ($row in $active) {
        if ($selected.Count -ge 8) { break }
        if (!$selected.Contains($row.B)) { $selected.Add($row.B) }
    }
    $selected = @($selected | Select-Object -First 8)

    $items = @()
    $pal = @(
        @("#0b594b","#07352e"),@("#6d4a1c","#2e2012"),@("#49315d","#21162c"),@("#27566a","#112d37"),
        @("#0c5145","#072c26"),@("#5e3b20","#2d1d12"),@("#43315d","#21182d"),@("#3f5961","#1f3036")
    )

    for($i=0;$i -lt 8;$i++){
        $name=$selected[$i]
        $vendorCourses=@($courses | Where-Object{ $_.B -eq $name -and $_.A.ToUpperInvariant() -ne "N" })
        $vendorAreas=@($areas | Where-Object{ $_.B -eq $name -and $_.A.ToUpperInvariant() -ne "N" } | Select-Object -ExpandProperty C -Unique)

        $region = if($vendorAreas.Count -gt 0){ ($vendorAreas | Select-Object -First 3) -join " · " } else { "추천 지역" }
        $price = Format-Price @($vendorCourses | Select-Object -ExpandProperty E)
        $kind = if($i -lt 4){"VIP"}else{"PREMIUM"}
        $num = if($i -lt 4){$i+1}else{$i-3}
        $prefix = if($i -lt 4){"vip"}else{"premium"}
        $imgWeb="/assets/images/main-recommend/$prefix-{0:00}.svg" -f $num
        $imgDisk=Join-Path $imageDir ("$prefix-{0:00}.svg" -f $num)
        $desc = if($i -lt 4){
            "$name 업체의 지역·코스·이용 정보를 한눈에 확인할 수 있는 VIP 추천입니다. 원하는 일정에 맞춰 상세 안내를 확인해보세요."
        }else{
            "$name 업체의 이용 조건과 코스 정보를 편리하게 비교할 수 있는 프리미엄 추천입니다. 상세페이지에서 필요한 내용을 확인해보세요."
        }
        $link=Resolve-Link $name $region
        Write-Svg $imgDisk $kind $name $region $pal[$i][0] $pal[$i][1]

        $items += [pscustomobject]@{
            kind=$kind.ToLower()
            shop=$name
            region=$region
            price=$price
            image=$imgWeb
            link=$link
            desc=$desc
            sort=$num
        }

        $targetXml=if($i -lt 4){$vipXml}else{$preXml}
        $row=if($i -lt 4){$i+2}else{$i-2}
        Set-InlineCell $targetXml $row "A" "Y"
        Set-InlineCell $targetXml $row "B" $name
        Set-InlineCell $targetXml $row "C" $region
        Set-InlineCell $targetXml $row "D" $price
        Set-InlineCell $targetXml $row "E" $imgWeb
        Set-InlineCell $targetXml $row "F" $link
        Set-InlineCell $targetXml $row "G" $desc
        Set-InlineCell $targetXml $row "H" ([string]$num)
    }

    Set-ZipText $zip $vipPath $vipXml.OuterXml
    Set-ZipText $zip $prePath $preXml.OuterXml
}
finally {
    $zip.Dispose()
}

Copy-Item -LiteralPath $temp -Destination $master -Force
Remove-Item -LiteralPath $temp -Force -ErrorAction SilentlyContinue

$vipItems=@($items | Where-Object{$_.kind -eq "vip"})
$preItems=@($items | Where-Object{$_.kind -eq "premium"})
$data=[ordered]@{vip=$vipItems;premium=$preItems}|ConvertTo-Json -Depth 6 -Compress

$js=@"
window.ML_MAIN_RECOMMEND_FINAL=$data;
(function(){
 const D=window.ML_MAIN_RECOMMEND_FINAL||{vip:[],premium:[]};
 const fake=["강남 골드케어","송파 밸런스룸","광교 리셋테라피","부평 온기케어"];
 const esc=s=>String(s??"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");

 function locateOld(title){
   const hs=[...document.querySelectorAll("h1,h2,h3,h4,b,strong,div")].filter(x=>!x.closest("#ml-main-recommend-final"));
   const h=hs.find(x=>(x.textContent||"").replace(/\s+/g," ").trim()===title);
   if(!h)return null;
   let n=h;
   for(let i=0;i<8 && n && n!==document.body;i++,n=n.parentElement){
     const t=(n.innerText||"").replace(/\s+/g," ");
     const hits=fake.filter(v=>t.includes(v)).length;
     const details=(t.match(/상세\s*보기/g)||[]).length;
     if(hits>=2 || details>=4)return n;
     if(n.tagName==="MAIN")break;
   }
   return null;
 }

 function hideOld(){
   ["VIP 추천","프리미엄 추천"].forEach(title=>{
      const b=locateOld(title);
      if(b && b.id!=="ml-main-recommend-final"){
        b.style.setProperty("display","none","important");
        b.dataset.mlOldRecommend="hidden";
      }
   });
   document.querySelectorAll("body *").forEach(el=>{
     if(el.closest("#ml-main-recommend-final"))return;
     const own=(el.textContent||"").replace(/\s+/g," ").trim();
     if(fake.includes(own)){
       let n=el;
       for(let i=0;i<7 && n && n!==document.body;i++,n=n.parentElement){
         const t=(n.innerText||"").replace(/\s+/g," ");
         if((t.match(/상세\s*보기/g)||[]).length>=1){
           n.style.setProperty("display","none","important");break;
         }
       }
     }
   });
 }

 function card(x,label){
   return '<a class="mlf-card" href="'+esc(x.link||"#")+'">'+
   '<div class="mlf-thumb"><img src="'+esc(x.image)+'" alt="'+esc(x.shop)+'"><span>'+esc(label)+'</span><b>'+esc(x.region)+'</b></div>'+
   '<div class="mlf-body"><h3>'+esc(x.shop)+'</h3><p>'+esc(x.desc)+'</p><div><strong>'+esc(x.price)+'</strong><em>상세보기</em></div></div></a>';
 }
 function group(title,label,arr){
   return '<section class="mlf-group"><div class="mlf-head"><h2>'+title+'</h2></div><div class="mlf-grid">'+arr.map(x=>card(x,label)).join("")+'</div></section>';
 }
 function style(){
   if(document.getElementById("mlf-style"))return;
   const s=document.createElement("style");s.id="mlf-style";
   s.textContent=`
   #ml-main-recommend-final{margin:30px 0 36px}
   #ml-main-recommend-final .mlf-group{margin:0 0 34px}
   #ml-main-recommend-final .mlf-head{display:flex;justify-content:space-between;align-items:center;margin:0 0 13px}
   #ml-main-recommend-final .mlf-head h2{margin:0;font-size:22px;color:#111}
   #ml-main-recommend-final .mlf-grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:10px}
   #ml-main-recommend-final .mlf-card{overflow:hidden;border:1px solid #ddd;border-radius:14px;background:#fff;text-decoration:none;color:#111;box-shadow:0 2px 8px rgba(0,0,0,.05)}
   #ml-main-recommend-final .mlf-thumb{height:120px;position:relative;background:#184f43;overflow:hidden}
   #ml-main-recommend-final .mlf-thumb img{width:100%;height:100%;object-fit:cover;display:block}
   #ml-main-recommend-final .mlf-thumb span{position:absolute;top:0;left:0;background:#efc342;color:#111;font-size:11px;font-weight:900;padding:15px 22px;border-radius:0 0 14px 0}
   #ml-main-recommend-final .mlf-thumb b{position:absolute;left:12px;bottom:10px;color:#fff;font-size:13px;text-shadow:0 1px 3px #000}
   #ml-main-recommend-final .mlf-body{padding:12px}
   #ml-main-recommend-final .mlf-body h3{margin:0 0 6px;font-size:15px}
   #ml-main-recommend-final .mlf-body p{height:38px;overflow:hidden;margin:0 0 14px;color:#777;font-size:11px;line-height:1.55}
   #ml-main-recommend-final .mlf-body div{display:flex;align-items:center;justify-content:space-between;gap:7px}
   #ml-main-recommend-final .mlf-body strong{font-size:14px}
   #ml-main-recommend-final .mlf-body em{font-style:normal;border:1px solid #e5ad32;border-radius:8px;padding:6px 9px;font-size:10px;font-weight:800}
   @media(max-width:760px){#ml-main-recommend-final .mlf-grid{grid-template-columns:repeat(2,minmax(0,1fr))}}
   `;
   document.head.appendChild(s);
 }
 function render(){
   hideOld();style();
   let root=document.getElementById("ml-main-recommend-final");
   if(!root){root=document.createElement("div");root.id="ml-main-recommend-final";}
   root.innerHTML=group("VIP 추천","VIP",D.vip)+group("프리미엄 추천","PICK",D.premium);
   const old=locateOld("VIP 추천")||locateOld("프리미엄 추천");
   const main=document.querySelector("main")||document.body;
   if(!root.isConnected){
      if(old&&old.parentNode)old.parentNode.insertBefore(root,old);
      else{
        const partner=main.querySelector(".partner");
        if(partner)main.insertBefore(root,partner);else main.appendChild(root);
      }
   }
   hideOld();
 }
 function boot(){
   render();
   let n=0;const timer=setInterval(()=>{render();if(++n>=20)clearInterval(timer)},250);
   const mo=new MutationObserver(()=>hideOld());mo.observe(document.body,{childList:true,subtree:true});
   setTimeout(()=>mo.disconnect(),10000);
 }
 if(document.readyState==="loading")document.addEventListener("DOMContentLoaded",boot);else boot();
})();
"@
[System.IO.File]::WriteAllText($jsFile,$js,[System.Text.UTF8Encoding]::new($false))

$html=Get-Content -LiteralPath $index -Raw -Encoding UTF8
$html=[regex]::Replace($html,'(?is)\s*<!--\s*ML-MAIN-RECOMMEND-FINAL-START\s*-->.*?<!--\s*ML-MAIN-RECOMMEND-FINAL-END\s*-->\s*',"`r`n")
$tag=@"
<!-- ML-MAIN-RECOMMEND-FINAL-START -->
<script src="assets/js/main-recommend-final.js?v=$stamp"></script>
<!-- ML-MAIN-RECOMMEND-FINAL-END -->
"@
$html=[regex]::Replace($html,'(?is)</body>',"$tag`r`n</body>",1)
Set-Content -LiteralPath $index -Value $html -Encoding UTF8

Write-Host ""
Write-Host "=== 메인 추천업체 최종 교체 완료 ==="
Write-Host "VIP:"
$vipItems | ForEach-Object{ Write-Host " -" $_.shop "|" $_.region "|" $_.price }
Write-Host "프리미엄:"
$preItems | ForEach-Object{ Write-Host " -" $_.shop "|" $_.region "|" $_.price }
Write-Host ""
Write-Host "추천 시트 자동 입력: 8개"
Write-Host "추천 이미지 생성: 8개"
Write-Host "상세페이지 링크 자동 탐색 완료"
Write-Host "기존 임시업체 숨김: 강남 골드케어 / 송파 밸런스룸 / 광교 리셋테라피 / 부평 온기케어"
Write-Host "백업 폴더:" $backupRoot
Write-Host ""
Write-Host "브라우저에서 Ctrl+F5 하세요."
