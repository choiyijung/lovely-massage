$ErrorActionPreference = "Stop"

$root      = "E:\mong-lounge-site"
$dataDir   = Join-Path $root "data"
$master    = Join-Path $dataDir "몽라운지_업체자동등록_마스터_v2.xlsx"
$index     = Join-Path $root "index.html"
$vipHtml   = Join-Path $root "vip.html"
$preHtml   = Join-Path $root "premium.html"
$shopsRoot = Join-Path $root "shops"
$imageDir  = Join-Path $root "assets\images\main-recommend"
$patchJs   = Join-Path $root "assets\js\main-recommend-existing-cards.js"

$stamp = Get-Date -Format "yyyyMMdd-HHmmss"
$backupRoot = "E:\mong-lounge-existing-recommend-cards-backup-$stamp"
$tempXlsx = Join-Path $env:TEMP "mong_lounge_recommend_cards_$stamp.xlsx"

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

if (!(Test-Path -LiteralPath $master)) { throw "마스터 엑셀 없음: $master" }

# 엑셀이 열려 있으면 안전하게 중단
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
    Write-Host "몽라운지_업체자동등록_마스터_v2.xlsx 파일을 닫은 뒤 다시 실행하세요."
    exit 1
}

New-Item -ItemType Directory -Force -Path $backupRoot,$imageDir,(Split-Path $patchJs -Parent) | Out-Null

foreach($f in @($master,$index,$vipHtml,$preHtml,$patchJs)) {
    if(Test-Path -LiteralPath $f) {
        Copy-Item -LiteralPath $f -Destination (Join-Path $backupRoot ([IO.Path]::GetFileName($f))) -Force
    }
}
Copy-Item -LiteralPath $master -Destination $tempXlsx -Force

function Get-ZipText {
    param([System.IO.Compression.ZipArchive]$Zip,[string]$Name)
    $e = $Zip.GetEntry($Name)
    if(!$e){ throw "XLSX 내부 파일 누락: $Name" }
    $s = $e.Open()
    try {
        $r = [IO.StreamReader]::new($s,[Text.Encoding]::UTF8,$true)
        try { return $r.ReadToEnd() } finally { $r.Dispose() }
    } finally { $s.Dispose() }
}

function Set-ZipText {
    param([System.IO.Compression.ZipArchive]$Zip,[string]$Name,[string]$Text)
    $old=$Zip.GetEntry($Name)
    if($old){$old.Delete()}
    $e=$Zip.CreateEntry($Name,[IO.Compression.CompressionLevel]::Optimal)
    $s=$e.Open()
    try {
        $w=[IO.StreamWriter]::new($s,[Text.UTF8Encoding]::new($false))
        try{$w.Write($Text);$w.Flush()}finally{$w.Dispose()}
    } finally {$s.Dispose()}
}

function Get-SharedStrings {
    param([System.IO.Compression.ZipArchive]$Zip)
    $list=New-Object System.Collections.Generic.List[string]
    if(!$Zip.GetEntry("xl/sharedStrings.xml")){return $list}
    [xml]$x=Get-ZipText $Zip "xl/sharedStrings.xml"
    foreach($si in $x.SelectNodes("//*[local-name()='si']")){
        $parts=@($si.SelectNodes(".//*[local-name()='t']")|ForEach-Object{$_.InnerText})
        $list.Add(($parts -join ""))
    }
    return $list
}

function Get-CellText {
    param([System.Xml.XmlElement]$Cell,$Shared)
    if(!$Cell){return ""}
    $t=$Cell.GetAttribute("t")
    if($t -eq "inlineStr"){
        return ((@($Cell.SelectNodes(".//*[local-name()='t']")|ForEach-Object{$_.InnerText})) -join "")
    }
    $v=$Cell.SelectSingleNode("./*[local-name()='v']")
    if(!$v){return ""}
    if($t -eq "s"){
        $i=0
        if([int]::TryParse($v.InnerText,[ref]$i) -and $i -ge 0 -and $i -lt $Shared.Count){
            return [string]$Shared[$i]
        }
    }
    return [string]$v.InnerText
}

function Get-SheetPath {
    param([xml]$Wb,[xml]$Rels,[string]$Name)
    $sheet=$Wb.SelectSingleNode("//*[local-name()='sheet' and @name='$Name']")
    if(!$sheet){throw "시트 없음: $Name"}
    $rid=$sheet.GetAttribute("id","http://schemas.openxmlformats.org/officeDocument/2006/relationships")
    $rel=$Rels.SelectSingleNode("//*[local-name()='Relationship' and @Id='$rid']")
    if(!$rel){throw "시트 관계 없음: $Name"}
    $target=$rel.GetAttribute("Target").Replace("\","/")
    if($target.StartsWith("/")){return $target.TrimStart("/")}
    return "xl/"+$target.TrimStart("/")
}

function Read-Rows {
    param([xml]$Sheet,$Shared,[string[]]$Cols)
    $out=@()
    foreach($row in $Sheet.SelectNodes("//*[local-name()='row' and number(@r)>=2]")){
        $r=[int]$row.GetAttribute("r")
        $o=[ordered]@{Row=$r}
        foreach($col in $Cols){
            $c=$Sheet.SelectSingleNode("//*[local-name()='c' and @r='$col$r']")
            $o[$col]=(Get-CellText $c $Shared).Trim()
        }
        $out += [pscustomobject]$o
    }
    return $out
}

function Set-InlineCell {
    param([xml]$Sheet,[int]$Row,[string]$Col,[string]$Text)
    $ns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    $sheetData=$Sheet.SelectSingleNode("/*[local-name()='worksheet']/*[local-name()='sheetData']")
    $rowNode=$Sheet.SelectSingleNode("//*[local-name()='row' and @r='$Row']")
    if(!$rowNode){
        $rowNode=$Sheet.CreateElement("row",$ns)
        [void]$rowNode.SetAttribute("r",[string]$Row)
        [void]$sheetData.AppendChild($rowNode)
    }
    $ref="$Col$Row"
    $cell=$Sheet.SelectSingleNode("//*[local-name()='c' and @r='$ref']")
    if(!$cell){
        $cell=$Sheet.CreateElement("c",$ns)
        [void]$cell.SetAttribute("r",$ref)
        [void]$rowNode.AppendChild($cell)
    }
    while($cell.HasChildNodes){[void]$cell.RemoveChild($cell.FirstChild)}
    [void]$cell.SetAttribute("t","inlineStr")
    $is=$Sheet.CreateElement("is",$ns)
    $tx=$Sheet.CreateElement("t",$ns)
    $tx.InnerText=$Text
    [void]$is.AppendChild($tx)
    [void]$cell.AppendChild($is)
}

function Min-Price {
    param($CourseRows,[string]$Shop)
    $nums=@()
    foreach($r in ($CourseRows|Where-Object{$_.B -eq $Shop -and $_.A.ToUpperInvariant() -ne "N"})){
        $ntext=($r.E -replace '[^\d]','')
        if($ntext){
            $n=0
            if([int]::TryParse($ntext,[ref]$n) -and $n -gt 0){$nums+=$n}
        }
    }
    if($nums.Count -eq 0){return "가격 문의"}
    $min=($nums|Measure-Object -Minimum).Minimum
    return ("{0:N0}원~" -f $min)
}

function Region-Text {
    param($AreaRows,[string]$Shop)
    $vals=@(
        $AreaRows |
        Where-Object{$_.B -eq $Shop -and $_.A.ToUpperInvariant() -ne "N" -and $_.C} |
        Select-Object -ExpandProperty C -Unique |
        Select-Object -First 3
    )
    if($vals.Count -eq 0){return "추천 지역"}
    return ($vals -join " · ")
}

function Resolve-Link {
    param([string]$Shop,[string]$Region)
    if(!(Test-Path -LiteralPath $shopsRoot)){return "#"}
    $cands=@(
        Get-ChildItem -LiteralPath $shopsRoot -Recurse -File -Filter "index.html" |
        Where-Object{(Split-Path (Split-Path $_.FullName -Parent) -Leaf) -eq $Shop}
    )
    if($cands.Count -eq 0){return "#"}
    $pick=$null
    $parts=@($Region -split '[·,/ ]+'|Where-Object{$_.Length -ge 2})
    foreach($c in $cands){
        try{
            $txt=Get-Content -LiteralPath $c.FullName -Raw -Encoding UTF8
            if($parts|Where-Object{$txt -match [regex]::Escape($_)}){$pick=$c;break}
        }catch{}
    }
    if(!$pick){$pick=$cands[0]}
    return "/"+$pick.FullName.Substring($root.Length).TrimStart("\").Replace("\","/")
}

function New-Svg {
    param([string]$Path,[string]$Badge,[string]$Shop,[string]$Region,[string]$C1,[string]$C2)
    $b=[Net.WebUtility]::HtmlEncode($Badge)
    $s=[Net.WebUtility]::HtmlEncode($Shop)
    $r=[Net.WebUtility]::HtmlEncode($Region)
    $svg=@"
<svg xmlns="http://www.w3.org/2000/svg" width="900" height="520" viewBox="0 0 900 520">
<defs><linearGradient id="g" x1="0" y1="0" x2="1" y2="1"><stop offset="0%" stop-color="$C1"/><stop offset="100%" stop-color="$C2"/></linearGradient></defs>
<rect width="900" height="520" fill="url(#g)"/><circle cx="760" cy="90" r="180" fill="#fff" opacity=".07"/>
<rect x="56" y="56" width="788" height="408" rx="34" fill="#061e19" opacity=".72"/>
<rect x="90" y="92" width="150" height="44" rx="22" fill="#ffc928"/>
<text x="165" y="121" text-anchor="middle" font-family="Arial,sans-serif" font-size="20" font-weight="800" fill="#083f35">$b</text>
<text x="90" y="250" font-family="Arial,sans-serif" font-size="48" font-weight="800" fill="#fff">$s</text>
<text x="92" y="310" font-family="Arial,sans-serif" font-size="24" fill="#ccefe6">$r</text>
<line x1="92" y1="355" x2="805" y2="355" stroke="#ffc928" stroke-width="2"/>
<text x="92" y="405" font-family="Arial,sans-serif" font-size="18" font-weight="700" fill="#fff" opacity=".85">MONG LOUNGE RECOMMEND</text>
</svg>
"@
    [IO.File]::WriteAllText($Path,$svg,[Text.UTF8Encoding]::new($false))
}

$zip=[IO.Compression.ZipFile]::Open($tempXlsx,[IO.Compression.ZipArchiveMode]::Update)
try{
    [xml]$wb=Get-ZipText $zip "xl/workbook.xml"
    [xml]$rels=Get-ZipText $zip "xl/_rels/workbook.xml.rels"
    $shared=Get-SharedStrings $zip

    $basicPath=Get-SheetPath $wb $rels "업체기본정보"
    $coursePath=Get-SheetPath $wb $rels "코스정보"
    $areaPath=Get-SheetPath $wb $rels "적용지역"
    $vipPath=Get-SheetPath $wb $rels "메인VIP추천"
    $prePath=Get-SheetPath $wb $rels "메인프리미엄추천"

    [xml]$basicXml=Get-ZipText $zip $basicPath
    [xml]$courseXml=Get-ZipText $zip $coursePath
    [xml]$areaXml=Get-ZipText $zip $areaPath
    [xml]$vipXml=Get-ZipText $zip $vipPath
    [xml]$preXml=Get-ZipText $zip $prePath

    $basic=@(Read-Rows $basicXml $shared @("A","B"))
    $course=@(Read-Rows $courseXml $shared @("A","B","E"))
    $area=@(Read-Rows $areaXml $shared @("A","B","C"))
    $vipRows=@(Read-Rows $vipXml $shared @("A","B","C","D","E","F","G","H"))
    $preRows=@(Read-Rows $preXml $shared @("A","B","C","D","E","F","G","H"))

    $active=@($basic|Where-Object{$_.B -and $_.A.ToUpperInvariant() -ne "N"}|Select-Object -ExpandProperty B -Unique)
    if($active.Count -lt 8){throw "활성 업체가 8개 미만입니다: $($active.Count)"}

    $used=New-Object System.Collections.Generic.HashSet[string]
    $colors=@(
        @("#0b594b","#07352e"),@("#6d4a1c","#2e2012"),@("#49315d","#21162c"),@("#27566a","#112d37"),
        @("#0c5145","#072c26"),@("#5e3b20","#2d1d12"),@("#43315d","#21182d"),@("#3f5961","#1f3036")
    )

    function Build-Items {
        param([xml]$Sheet,$Existing,[string]$Kind,[string]$Prefix,[int]$ColorStart)

        $items=@()
        for($slot=1;$slot -le 4;$slot++){
            $row=$slot+1
            $e=$Existing|Where-Object{$_.Row -eq $row}|Select-Object -First 1
            $shop=if($e){$e.B}else{""}

            if([string]::IsNullOrWhiteSpace($shop) -or !($active -contains $shop) -or $used.Contains($shop)){
                $shop=$active|Where-Object{!$used.Contains($_)}|Select-Object -First 1
            }
            if(!$shop){throw "$Kind 추천업체 자동선택 실패"}
            [void]$used.Add($shop)

            $region=if($e -and $e.C){$e.C}else{Region-Text $area $shop}
            $price=if($e -and $e.D){$e.D}else{Min-Price $course $shop}
            $link=if($e -and $e.F){$e.F}else{Resolve-Link $shop $region}
            $desc=if($e -and $e.G){$e.G}else{
                if($Kind -eq "VIP"){
                    "$shop 업체의 지역과 코스 정보를 편리하게 확인할 수 있는 VIP 추천입니다. 상세페이지에서 이용 조건을 확인해보세요."
                }else{
                    "$shop 업체의 이용 정보와 코스를 한눈에 비교할 수 있는 프리미엄 추천입니다. 상세 안내를 확인해보세요."
                }
            }

            $imgName="$Prefix-{0:00}.svg" -f $slot
            $imgWeb="/assets/images/main-recommend/$imgName"
            $imgDisk=Join-Path $imageDir $imgName
            $pair=$colors[$ColorStart+$slot-1]
            New-Svg $imgDisk $Kind $shop $region $pair[0] $pair[1]

            Set-InlineCell $Sheet $row "A" "Y"
            Set-InlineCell $Sheet $row "B" $shop
            Set-InlineCell $Sheet $row "C" $region
            Set-InlineCell $Sheet $row "D" $price
            Set-InlineCell $Sheet $row "E" $imgWeb
            Set-InlineCell $Sheet $row "F" $link
            Set-InlineCell $Sheet $row "G" $desc
            Set-InlineCell $Sheet $row "H" ([string]$slot)

            $items += [pscustomobject]@{
                shop=$shop;region=$region;price=$price;image=$imgWeb;link=$link;desc=$desc;sort=$slot
            }
        }
        return $items
    }

    $vipItems=@(Build-Items $vipXml $vipRows "VIP" "vip" 0)
    $preItems=@(Build-Items $preXml $preRows "PREMIUM" "premium" 4)

    Set-ZipText $zip $vipPath $vipXml.OuterXml
    Set-ZipText $zip $prePath $preXml.OuterXml
}
finally{$zip.Dispose()}

Copy-Item -LiteralPath $tempXlsx -Destination $master -Force
Remove-Item -LiteralPath $tempXlsx -Force -ErrorAction SilentlyContinue

# vip.html / premium.html의 기존 카드 텍스트도 실제 추천업체로 직접 교체
$legacyNames=@("강남 골드케어","송파 밸런스룸","광교 리셋테라피","부평 온기케어")

function Rewrite-RecommendPage {
    param([string]$Path,$Items)
    if(!(Test-Path -LiteralPath $Path)){return 0}

    $html=Get-Content -LiteralPath $Path -Raw -Encoding UTF8
    $before=$html

    for($i=0;$i -lt 4;$i++){
        $old=[regex]::Escape($legacyNames[$i])
        $item=$Items[$i]

        # h3 + 바로 뒤 p + b를 한 묶음으로 교체
        $pat="(?is)<h3>\s*$old\s*</h3>\s*<p>.*?</p>\s*<b>.*?</b>"
        $rep="<h3>$($item.shop)</h3><p>$($item.region) · 추천 업체</p><b>$($item.price)</b>"
        $html=[regex]::Replace($html,$pat,[System.Text.RegularExpressions.MatchEvaluator]{param($m)$rep},1)
    }

    if($html -ne $before){
        Set-Content -LiteralPath $Path -Value $html -Encoding UTF8
        return 1
    }
    return 0
}

$vipSourceChanged=Rewrite-RecommendPage $vipHtml $vipItems
$preSourceChanged=Rewrite-RecommendPage $preHtml $preItems

# 기존 화면 카드 자체를 "제자리"에서 바꾸는 JS
$data=[ordered]@{vip=$vipItems;premium=$preItems}|ConvertTo-Json -Depth 6 -Compress

$patch=@"
window.ML_EXISTING_RECOMMEND_DATA=$data;
(function(){
  const D=window.ML_EXISTING_RECOMMEND_DATA||{vip:[],premium:[]};
  const LEGACY=["강남 골드케어","송파 밸런스룸","광교 리셋테라피","부평 온기케어"];
  const OLD_REGIONS=["서울 강남구","서울 송파구","경기 수원시","인천 부평구"];

  function setCard(h,item){
    if(!h||!item)return;
    const card=h.closest("a")||h.closest("article")||h.parentElement?.parentElement||h.parentElement;
    h.textContent=item.shop;

    if(card){
      if(card.tagName==="A" && item.link && item.link!=="#") card.setAttribute("href",item.link);
      const a=card.querySelector("a");
      if(a && item.link && item.link!=="#") a.setAttribute("href",item.link);

      const img=card.querySelector("img");
      if(img && item.image){img.setAttribute("src",item.image);img.setAttribute("alt",item.shop);}

      const ps=[...card.querySelectorAll("p")];
      if(ps.length) ps[0].textContent=item.region;

      const price=[...card.querySelectorAll("b,strong,span")].find(el=>/원/.test((el.textContent||"").trim()));
      if(price) price.textContent=item.price;

      [...card.querySelectorAll("*")].forEach(el=>{
        const t=(el.textContent||"").replace(/\s+/g," ").trim();
        const ix=OLD_REGIONS.indexOf(t);
        if(ix>=0) el.textContent=item.region;
      });
    }
  }

  function apply(){
    // 1) 메인에서 예전 4개가 두 번(VIP/PICK) 나타나는 경우: 등장 순서대로 4+4 매핑
    const legacyH3=[...document.querySelectorAll("h3")].filter(h=>LEGACY.includes((h.textContent||"").trim()));
    legacyH3.forEach((h,i)=>{
      const item=i<4 ? D.vip[i] : D.premium[i-4];
      if(item) setCard(h,item);
    });

    // 2) vip.html / premium.html처럼 이미 이름만 바뀐 경우에도 링크/이미지/가격을 맞춤
    const all=[...(D.vip||[]),...(D.premium||[])];
    [...document.querySelectorAll("h3")].forEach(h=>{
      const name=(h.textContent||"").trim();
      const item=all.find(x=>x.shop===name);
      if(item) setCard(h,item);
    });
  }

  function boot(){
    apply();
    requestAnimationFrame(apply);
    let n=0;
    const timer=setInterval(()=>{apply();if(++n>=24)clearInterval(timer)},250);
    const mo=new MutationObserver(()=>apply());
    mo.observe(document.body,{childList:true,subtree:true});
    setTimeout(()=>mo.disconnect(),8000);
  }

  if(document.readyState==="loading")document.addEventListener("DOMContentLoaded",boot);
  else boot();
})();
"@

[IO.File]::WriteAllText($patchJs,$patch,[Text.UTF8Encoding]::new($false))

# index / vip / premium에 패치 JS 하나만 연결
function Ensure-PatchTag {
    param([string]$Path)
    if(!(Test-Path -LiteralPath $Path)){return}
    $html=Get-Content -LiteralPath $Path -Raw -Encoding UTF8
    $html=[regex]::Replace(
        $html,
        '(?is)\s*<!--\s*ML-EXISTING-RECOMMEND-PATCH-START\s*-->.*?<!--\s*ML-EXISTING-RECOMMEND-PATCH-END\s*-->\s*',
        "`r`n"
    )
    $tag=@"
<!-- ML-EXISTING-RECOMMEND-PATCH-START -->
<script src="assets/js/main-recommend-existing-cards.js?v=$stamp"></script>
<!-- ML-EXISTING-RECOMMEND-PATCH-END -->
"@
    $html=[regex]::Replace($html,'(?is)</body>',"$tag`r`n</body>",1)
    Set-Content -LiteralPath $Path -Value $html -Encoding UTF8
}

Ensure-PatchTag $index
Ensure-PatchTag $vipHtml
Ensure-PatchTag $preHtml

Write-Host ""
Write-Host "=== 기존 VIP/프리미엄 카드 실제업체 교체 완료 ==="
Write-Host "VIP 추천:"
$vipItems|ForEach-Object{Write-Host " -" $_.shop "|" $_.region "|" $_.price}
Write-Host "프리미엄 추천:"
$preItems|ForEach-Object{Write-Host " -" $_.shop "|" $_.region "|" $_.price}
Write-Host ""
Write-Host "vip.html 원본 수정:" $vipSourceChanged
Write-Host "premium.html 원본 수정:" $preSourceChanged
Write-Host "메인 기존카드 제자리 교체 JS:" (Test-Path -LiteralPath $patchJs)
Write-Host "추천 이미지: 8개"
Write-Host "백업 폴더:" $backupRoot
Write-Host ""
Write-Host "5512 주소에서 Ctrl+F5 하세요."
