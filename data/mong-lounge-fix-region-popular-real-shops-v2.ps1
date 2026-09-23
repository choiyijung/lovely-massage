$ErrorActionPreference = "Stop"

$root      = "E:\mong-lounge-site"
$dataDir   = Join-Path $root "data"
$master    = Join-Path $dataDir "몽라운지_업체자동등록_마스터_v2.xlsx"
$appJs     = Join-Path $root "assets\js\app.js"
$index     = Join-Path $root "index.html"
$shopsRoot = Join-Path $root "shops"

$stamp     = Get-Date -Format "yyyyMMdd-HHmmss"
$backup    = "E:\mong-lounge-region-popular-backup-$stamp"
$temp      = Join-Path $env:TEMP "mong_lounge_region_popular_$stamp.xlsx"

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

if (!(Test-Path -LiteralPath $master)) { throw "마스터 엑셀 없음: $master" }
if (!(Test-Path -LiteralPath $appJs))  { throw "app.js 없음: $appJs" }
if (!(Test-Path -LiteralPath $index))  { throw "index.html 없음: $index" }

New-Item -ItemType Directory -Force -Path $backup | Out-Null
Copy-Item -LiteralPath $appJs -Destination (Join-Path $backup "app.js") -Force
Copy-Item -LiteralPath $index -Destination (Join-Path $backup "index.html") -Force

# 엑셀이 열려 있어도 읽을 수 있게 임시복사
$inStream = [System.IO.File]::Open(
    $master,
    [System.IO.FileMode]::Open,
    [System.IO.FileAccess]::Read,
    [System.IO.FileShare]::ReadWrite
)
try {
    $outStream = [System.IO.File]::Create($temp)
    try { $inStream.CopyTo($outStream) }
    finally { $outStream.Dispose() }
}
finally { $inStream.Dispose() }

function Get-ZipText {
    param([System.IO.Compression.ZipArchive]$Zip,[string]$Name)
    $e=$Zip.GetEntry($Name)
    if(!$e){throw "XLSX 내부 파일 누락: $Name"}
    $s=$e.Open()
    try{
        $r=[IO.StreamReader]::new($s,[Text.Encoding]::UTF8,$true)
        try{return $r.ReadToEnd()}finally{$r.Dispose()}
    }finally{$s.Dispose()}
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
        $o=[ordered]@{}
        foreach($c in $Cols){
            $cell=$Sheet.SelectSingleNode("//*[local-name()='c' and @r='$c$r']")
            $o[$c]=(Get-CellText $cell $Shared).Trim()
        }
        $out += [pscustomobject]$o
    }
    return $out
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

function Resolve-Link {
    param([string]$Shop,[string]$RegionHint)

    if(!(Test-Path -LiteralPath $shopsRoot)){return "#"}

    $cands=@(
        Get-ChildItem -LiteralPath $shopsRoot -Recurse -File -Filter "index.html" |
        Where-Object{(Split-Path (Split-Path $_.FullName -Parent) -Leaf) -eq $Shop}
    )

    if($cands.Count -eq 0){return "#"}

    $pick=$null
    if($RegionHint){
        $parts=@($RegionHint -split '[·,/ ]+'|Where-Object{$_.Length -ge 2})
        foreach($c in $cands){
            try{
                $txt=Get-Content -LiteralPath $c.FullName -Raw -Encoding UTF8
                if($parts|Where-Object{$txt -match [regex]::Escape($_)}){
                    $pick=$c
                    break
                }
            }catch{}
        }
    }
    if(!$pick){$pick=$cands[0]}

    return "/"+$pick.FullName.Substring($root.Length).TrimStart("\").Replace("\","/")
}

function Category-ForArea {
    param([string]$Area)

    if(!$Area){return $null}
    $a=$Area.Replace(" ","")

    # 광주광역시는 기타지역, 경기광주시는 경기
    if($a -match '^(서울|강남|강동|강북|강서|관악|광진|구로|금천|노원|도봉|동대문|동작|마포|서대문|서초|성동|성북|송파|양천|영등포|용산|은평|종로|중구|중랑)'){
        return "서울"
    }

    if($a -match '^(경기|수원|성남|고양|용인|안양|안산|부천|화성|광명|평택|동두천|의정부|파주|하남|의왕|군포|양주|포천|안성|이천|오산|구리|시흥|여주|연천|가평|양평|김포|남양주|경기광주시|광주시)'){
        return "경기"
    }

    if($a -match '^(인천|강화|검단|계양|남동|미추홀|부평|연수|영종|옹진|제물포|서해구|인천중구)'){
        return "인천"
    }

    return "기타지역"
}

$zip=[IO.Compression.ZipFile]::OpenRead($temp)
try{
    [xml]$wb=Get-ZipText $zip "xl/workbook.xml"
    [xml]$rels=Get-ZipText $zip "xl/_rels/workbook.xml.rels"
    $shared=Get-SharedStrings $zip

    [xml]$basicXml=Get-ZipText $zip (Get-SheetPath $wb $rels "업체기본정보")
    [xml]$courseXml=Get-ZipText $zip (Get-SheetPath $wb $rels "코스정보")
    [xml]$areaXml=Get-ZipText $zip (Get-SheetPath $wb $rels "적용지역")

    $basic=@(Read-Rows $basicXml $shared @("A","B"))
    $course=@(Read-Rows $courseXml $shared @("A","B","E"))
    $area=@(Read-Rows $areaXml $shared @("A","B","C"))
}
finally{
    $zip.Dispose()
    Remove-Item -LiteralPath $temp -Force -ErrorAction SilentlyContinue
}

$active=@(
    $basic |
    Where-Object{$_.B -and $_.A.ToUpperInvariant() -ne "N"} |
    Select-Object -ExpandProperty B -Unique
)

$groups=[ordered]@{
    "서울"=@()
    "경기"=@()
    "인천"=@()
    "기타지역"=@()
}

# OrderedDictionary를 열거하면서 값을 바꾸면 "컬렉션이 수정되었습니다" 오류가 날 수 있으므로
# 키 목록을 별도 배열로 고정해서 사용합니다.
$groupKeys = @($groups.Keys)

foreach($shop in $active){
    $areas=@(
        $area |
        Where-Object{$_.B -eq $shop -and $_.A.ToUpperInvariant() -ne "N" -and $_.C} |
        Select-Object -ExpandProperty C -Unique
    )

    foreach($cat in $groupKeys){
        $matchAreas=@(
            $areas | Where-Object{(Category-ForArea $_) -eq $cat}
        )
        if($matchAreas.Count -eq 0){continue}

        $displayArea=($matchAreas|Select-Object -First 3) -join " · "
        $groups[$cat] += [pscustomobject]@{
            name=$shop
            area=$displayArea
            price=Min-Price $course $shop
            link=Resolve-Link $shop $displayArea
        }
    }
}

# 각 탭 4개. 부족하면 다른 실제 업체를 중복 없이 보충
foreach($cat in $groupKeys){
    $list=New-Object System.Collections.Generic.List[object]
    foreach($x in $groups[$cat]){
        if(!$list.Where({param($y)$y.name -eq $x.name}).Count){
            $list.Add($x)
        }
        if($list.Count -ge 4){break}
    }

    if($list.Count -lt 4){
        foreach($shop in $active){
            if($list.Where({param($y)$y.name -eq $shop}).Count){continue}
            $allAreas=@(
                $area |
                Where-Object{$_.B -eq $shop -and $_.A.ToUpperInvariant() -ne "N" -and $_.C} |
                Select-Object -ExpandProperty C -Unique |
                Select-Object -First 3
            )
            $hint=if($allAreas.Count){$allAreas -join " · "}else{"추천 지역"}
            $list.Add([pscustomobject]@{
                name=$shop
                area=$hint
                price=Min-Price $course $shop
                link=Resolve-Link $shop $hint
            })
            if($list.Count -ge 4){break}
        }
    }

    $groups[$cat]=@($list|Select-Object -First 4)
}

# JS용 객체: 기존 x[0],x[1],x[2] 구조를 유지하면서 x[3]에 링크 추가
$regionObject=[ordered]@{}
foreach($cat in $groupKeys){
    $regionObject[$cat]=@(
        $groups[$cat] | ForEach-Object{
            @($_.name,$_.area,$_.price,$_.link)
        }
    )
}
$json=$regionObject|ConvertTo-Json -Depth 6 -Compress

$js=Get-Content -LiteralPath $appJs -Raw -Encoding UTF8
$before=$js

# 기존 const regions = {...}; 전체 교체
$js=[regex]::Replace(
    $js,
    '(?is)\s*const\s+regions\s*=\s*\{.*?\};\s*(?=\s*(?:const|document\.querySelector|function|let|var))',
    "`r`n      const regions = $json;`r`n",
    1
)

# 기존 regionItems 렌더러를 링크 가능한 실제업체 행으로 교체
$renderPattern='(?is)document\.querySelector\("#regionItems"\)\.innerHTML\s*=\s*regions\[key\]\.map\(\(x,i\)=>`.*?`\)\.join\(""\);'
$renderReplacement=@'
document.querySelector("#regionItems").innerHTML = regions[key].map((x,i)=>`
    <a class="region-row" href="${x[3] || '#'}">
      <i>${i+1}</i>
      <div>
        <h3>${x[0]}</h3>
        <p>${x[1]} · 등록 업체</p>
      </div>
      <strong>${x[2]}</strong>
    </a>`).join("");
'@
$js=[regex]::Replace($js,$renderPattern,$renderReplacement,1)

if($js -eq $before){
    throw "app.js의 지역추천 구조를 찾지 못했습니다. 파일은 수정되지 않았습니다."
}

Set-Content -LiteralPath $appJs -Value $js -Encoding UTF8

# index의 부산 탭을 기타지역으로 변경
$html=Get-Content -LiteralPath $index -Raw -Encoding UTF8
$html=$html.Replace('data-key="부산">부산<','data-key="기타지역">기타지역<')
$html=$html.Replace('data-region="부산">부산<','data-region="기타지역">기타지역<')
$html=$html.Replace('data-tab="부산">부산<','data-tab="기타지역">기타지역<')

# 단순 버튼/앵커 형태도 보정
$html=[regex]::Replace(
    $html,
    '(<(?:button|a)\b[^>]*>)(\s*)부산(\s*</(?:button|a)>)',
    '$1$2기타지역$3'
)

# app.js 캐시 갱신
$html=[regex]::Replace(
    $html,
    'assets/js/app\.js(?:\?v=[^"'']*)?',
    "assets/js/app.js?v=$stamp"
)

Set-Content -LiteralPath $index -Value $html -Encoding UTF8

# 검증
$check=Get-Content -LiteralPath $appJs -Raw -Encoding UTF8
$idx=Get-Content -LiteralPath $index -Raw -Encoding UTF8

Write-Host ""
Write-Host "=== 지역별 인기 추천 실제업체 교체 완료 ==="
foreach($cat in $groupKeys){
    Write-Host ""
    Write-Host "[$cat]"
    $groups[$cat]|ForEach-Object{
        Write-Host " -" $_.name "|" $_.area "|" $_.price
    }
}
Write-Host ""
Write-Host "기존 임시 지역업체 남음:" ($check -match '부평 온기케어|연수 힐링룸|구월 밸런스|계양 리셋룸|강남 골드케어|광교 리셋테라피')
Write-Host "지역 카드 상세링크 사용:" ($check -match 'x\[3\]')
Write-Host "부산 탭 남음:" ($idx -match '>부산<')
Write-Host "기타지역 탭 확인:" ($idx -match '>기타지역<')
Write-Host "app.js 캐시 버전 갱신:" ($idx -match [regex]::Escape("app.js?v=$stamp"))
Write-Host ""
Write-Host "백업 폴더:" $backup
Write-Host "이제 5512 메인에서 Ctrl+F5 하세요."
