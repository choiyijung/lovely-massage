$ErrorActionPreference = "Stop"

$root      = "E:\mong-lounge-site"
$dataDir   = Join-Path $root "data"
$master    = Join-Path $dataDir "몽라운지_업체자동등록_마스터_v2.xlsx"
$appJs     = Join-Path $root "assets\js\app.js"
$index     = Join-Path $root "index.html"
$shopsRoot = Join-Path $root "shops"

$stamp = Get-Date -Format "yyyyMMdd-HHmmss"
$currentBackup = "E:\mong-lounge-region-popular-v3-current-backup-$stamp"

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

if (!(Test-Path -LiteralPath $master)) { throw "마스터 엑셀 없음: $master" }
if (!(Test-Path -LiteralPath $appJs))  { throw "app.js 없음: $appJs" }
if (!(Test-Path -LiteralPath $index))  { throw "index.html 없음: $index" }

# 현재 깨진 상태도 별도 백업
New-Item -ItemType Directory -Force -Path $currentBackup | Out-Null
Copy-Item -LiteralPath $appJs -Destination (Join-Path $currentBackup "app.js") -Force
Copy-Item -LiteralPath $index -Destination (Join-Path $currentBackup "index.html") -Force

# 직전 지역추천 작업 전 백업에서 정상 app.js / index.html 복구
$restoreCandidates = @(
    Get-ChildItem "E:\" -Directory -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -like "mong-lounge-region-popular-backup-*" } |
    Sort-Object LastWriteTime -Descending
)

$restore = $restoreCandidates |
    Where-Object {
        (Test-Path -LiteralPath (Join-Path $_.FullName "app.js")) -and
        (Test-Path -LiteralPath (Join-Path $_.FullName "index.html"))
    } |
    Select-Object -First 1

if (!$restore) {
    throw "복구용 mong-lounge-region-popular-backup-* 폴더에서 app.js/index.html을 찾지 못했습니다."
}

Copy-Item -LiteralPath (Join-Path $restore.FullName "app.js") -Destination $appJs -Force
Copy-Item -LiteralPath (Join-Path $restore.FullName "index.html") -Destination $index -Force

# 엑셀이 열려 있어도 읽기 가능한 복사본 생성
$temp = Join-Path $env:TEMP "mong_lounge_region_v3_$stamp.xlsx"
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
    $e = $Zip.GetEntry($Name)
    if(!$e){ throw "XLSX 내부 파일 누락: $Name" }
    $s = $e.Open()
    try {
        $r = [System.IO.StreamReader]::new($s,[System.Text.Encoding]::UTF8,$true)
        try { return $r.ReadToEnd() }
        finally { $r.Dispose() }
    }
    finally { $s.Dispose() }
}

function Get-SharedStrings {
    param([System.IO.Compression.ZipArchive]$Zip)

    $list = New-Object System.Collections.Generic.List[string]
    if(!$Zip.GetEntry("xl/sharedStrings.xml")){ return $list }

    [xml]$x = Get-ZipText $Zip "xl/sharedStrings.xml"
    foreach($si in $x.SelectNodes("//*[local-name()='si']")){
        $parts = @(
            $si.SelectNodes(".//*[local-name()='t']") |
            ForEach-Object { $_.InnerText }
        )
        $list.Add(($parts -join ""))
    }
    return $list
}

function Get-CellText {
    param([System.Xml.XmlElement]$Cell,$Shared)

    if(!$Cell){ return "" }
    $t = $Cell.GetAttribute("t")

    if($t -eq "inlineStr"){
        return (
            @(
                $Cell.SelectNodes(".//*[local-name()='t']") |
                ForEach-Object { $_.InnerText }
            ) -join ""
        )
    }

    $v = $Cell.SelectSingleNode("./*[local-name()='v']")
    if(!$v){ return "" }

    if($t -eq "s"){
        $i = 0
        if([int]::TryParse($v.InnerText,[ref]$i) -and $i -ge 0 -and $i -lt $Shared.Count){
            return [string]$Shared[$i]
        }
    }

    return [string]$v.InnerText
}

function Get-SheetPath {
    param([xml]$Wb,[xml]$Rels,[string]$Name)

    $sheet = $Wb.SelectSingleNode("//*[local-name()='sheet' and @name='$Name']")
    if(!$sheet){ throw "시트 없음: $Name" }

    $rid = $sheet.GetAttribute(
        "id",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    )
    $rel = $Rels.SelectSingleNode("//*[local-name()='Relationship' and @Id='$rid']")
    if(!$rel){ throw "시트 관계 없음: $Name" }

    $target = $rel.GetAttribute("Target").Replace("\","/")
    if($target.StartsWith("/")){ return $target.TrimStart("/") }
    return "xl/" + $target.TrimStart("/")
}

function Read-Rows {
    param([xml]$Sheet,$Shared,[string[]]$Cols)

    $out = @()
    foreach($row in $Sheet.SelectNodes("//*[local-name()='row' and number(@r)>=2]")){
        $r = [int]$row.GetAttribute("r")
        $o = [ordered]@{}

        foreach($c in $Cols){
            $cell = $Sheet.SelectSingleNode("//*[local-name()='c' and @r='$c$r']")
            $o[$c] = (Get-CellText $cell $Shared).Trim()
        }

        $out += [pscustomobject]$o
    }
    return $out
}

function Min-Price {
    param($CourseRows,[string]$Shop)

    $nums = @()
    foreach($r in ($CourseRows | Where-Object {
        $_.B -eq $Shop -and $_.A.ToUpperInvariant() -ne "N"
    })){
        $nText = ($r.E -replace '[^\d]','')
        if($nText){
            $n = 0
            if([int]::TryParse($nText,[ref]$n) -and $n -gt 0){ $nums += $n }
        }
    }

    if($nums.Count -eq 0){ return "가격 문의" }

    $min = ($nums | Measure-Object -Minimum).Minimum
    return ("{0:N0}원~" -f $min)
}

function Resolve-Link {
    param([string]$Shop,[string]$RegionHint)

    if(!(Test-Path -LiteralPath $shopsRoot)){ return "#" }

    $cands = @(
        Get-ChildItem -LiteralPath $shopsRoot -Recurse -File -Filter "index.html" |
        Where-Object {
            (Split-Path (Split-Path $_.FullName -Parent) -Leaf) -eq $Shop
        }
    )

    if($cands.Count -eq 0){ return "#" }

    $pick = $null
    $parts = @(
        $RegionHint -split '[·,/ ]+' |
        Where-Object { $_.Length -ge 2 }
    )

    foreach($c in $cands){
        try {
            $txt = Get-Content -LiteralPath $c.FullName -Raw -Encoding UTF8
            if($parts | Where-Object { $txt -match [regex]::Escape($_) }){
                $pick = $c
                break
            }
        } catch {}
    }

    if(!$pick){ $pick = $cands[0] }

    return "/" + $pick.FullName.Substring($root.Length).TrimStart("\").Replace("\","/")
}

function Category-ForArea {
    param([string]$Area)

    if(!$Area){ return $null }
    $a = $Area.Replace(" ","")

    if($a -match '^(서울|강남|강동|강북|강서|관악|광진|구로|금천|노원|도봉|동대문|동작|마포|서대문|서초|성동|성북|송파|양천|영등포|용산|은평|종로|중랑|서울중구)'){
        return "서울"
    }

    if($a -match '^(경기|수원|성남|고양|용인|안양|안산|부천|화성|광명|평택|동두천|의정부|파주|하남|의왕|군포|양주|포천|안성|이천|오산|구리|시흥|여주|연천|가평|양평|김포|남양주|경기광주시|경기도광주|광주시)'){
        return "경기"
    }

    if($a -match '^(인천|강화|검단|계양|남동|미추홀|부평|연수|영종|옹진|제물포|서해구|인천중구)'){
        return "인천"
    }

    return "기타지역"
}

$zip = [System.IO.Compression.ZipFile]::OpenRead($temp)
try {
    [xml]$wb   = Get-ZipText $zip "xl/workbook.xml"
    [xml]$rels = Get-ZipText $zip "xl/_rels/workbook.xml.rels"
    $shared    = Get-SharedStrings $zip

    [xml]$basicXml  = Get-ZipText $zip (Get-SheetPath $wb $rels "업체기본정보")
    [xml]$courseXml = Get-ZipText $zip (Get-SheetPath $wb $rels "코스정보")
    [xml]$areaXml   = Get-ZipText $zip (Get-SheetPath $wb $rels "적용지역")

    $basic  = @(Read-Rows $basicXml  $shared @("A","B"))
    $course = @(Read-Rows $courseXml $shared @("A","B","E"))
    $area   = @(Read-Rows $areaXml   $shared @("A","B","C"))
}
finally {
    $zip.Dispose()
    Remove-Item -LiteralPath $temp -Force -ErrorAction SilentlyContinue
}

$active = @(
    $basic |
    Where-Object { $_.B -and $_.A.ToUpperInvariant() -ne "N" } |
    Select-Object -ExpandProperty B -Unique
)

$groupKeys = @("서울","경기","인천","기타지역")

$groups = [ordered]@{
    "서울"      = @()
    "경기"      = @()
    "인천"      = @()
    "기타지역"  = @()
}

# 실제 적용지역 기준으로 업체 후보 생성
foreach($shop in $active){
    $shopAreas = @(
        $area |
        Where-Object {
            $_.B -eq $shop -and
            $_.A.ToUpperInvariant() -ne "N" -and
            $_.C
        } |
        Select-Object -ExpandProperty C -Unique
    )

    foreach($cat in $groupKeys){
        $matched = @(
            $shopAreas |
            Where-Object { (Category-ForArea $_) -eq $cat }
        )

        if($matched.Count -eq 0){ continue }

        $displayArea = ($matched | Select-Object -First 3) -join " · "

        $groups[$cat] += [pscustomobject]@{
            name  = $shop
            area  = $displayArea
            price = Min-Price $course $shop
            link  = Resolve-Link $shop $displayArea
        }
    }
}

# 각 탭 4개로 정리. 부족한 경우 실제 등록업체 중 다른 업체로 보충.
foreach($cat in $groupKeys){
    $picked = New-Object System.Collections.Generic.List[object]
    $seen = New-Object System.Collections.Generic.HashSet[string]

    foreach($x in @($groups[$cat])){
        if($seen.Add($x.name)){
            $picked.Add($x)
        }
        if($picked.Count -ge 4){ break }
    }

    if($picked.Count -lt 4){
        foreach($shop in $active){
            if(!$seen.Add($shop)){ continue }

            $allAreas = @(
                $area |
                Where-Object {
                    $_.B -eq $shop -and
                    $_.A.ToUpperInvariant() -ne "N" -and
                    $_.C
                } |
                Select-Object -ExpandProperty C -Unique |
                Select-Object -First 3
            )

            $hint = if($allAreas.Count){ $allAreas -join " · " } else { "추천 지역" }

            $picked.Add([pscustomobject]@{
                name  = $shop
                area  = $hint
                price = Min-Price $course $shop
                link  = Resolve-Link $shop $hint
            })

            if($picked.Count -ge 4){ break }
        }
    }

    $groups[$cat] = @($picked | Select-Object -First 4)
}

# 중요: 문자열 배열이 아니라 객체 배열로 JSON 생성
$regionObject = [ordered]@{}
foreach($cat in $groupKeys){
    $regionObject[$cat] = @(
        $groups[$cat] |
        ForEach-Object {
            [pscustomobject]@{
                name  = $_.name
                area  = $_.area
                price = $_.price
                link  = $_.link
            }
        }
    )
}

$json = $regionObject | ConvertTo-Json -Depth 8 -Compress

$js = Get-Content -LiteralPath $appJs -Raw -Encoding UTF8

$oldRegionPattern = '(?is)const\s+regions\s*=\s*\{.*?\};'
if(![regex]::IsMatch($js,$oldRegionPattern)){
    throw "복구된 app.js에서 const regions = {...}; 구조를 찾지 못했습니다."
}

$js = [regex]::Replace(
    $js,
    $oldRegionPattern,
    "const regions = $json;",
    1
)

$renderPattern = '(?is)document\.querySelector\("#regionItems"\)\.innerHTML\s*=\s*regions\[key\]\.map\(\(x,i\)=>`.*?`\)\.join\(""\);'

$renderReplacement = @'
document.querySelector("#regionItems").innerHTML = regions[key].map((x,i)=>`
    <a class="region-row" href="${x.link || '#'}" style="text-decoration:none;color:inherit">
      <i>${i+1}</i>
      <div>
        <h3>${x.name}</h3>
        <p>${x.area} · 등록 업체</p>
      </div>
      <strong>${x.price}</strong>
    </a>`).join("");
'@

if(![regex]::IsMatch($js,$renderPattern)){
    throw "복구된 app.js에서 regionItems 렌더링 코드를 찾지 못했습니다."
}

$js = [regex]::Replace(
    $js,
    $renderPattern,
    $renderReplacement,
    1
)

Set-Content -LiteralPath $appJs -Value $js -Encoding UTF8

# 부산 탭 -> 기타지역
$html = Get-Content -LiteralPath $index -Raw -Encoding UTF8

$html = [regex]::Replace(
    $html,
    '((?:data-key|data-region|data-tab)=["''])부산(["''])',
    '$1기타지역$2'
)

$html = [regex]::Replace(
    $html,
    '(<(?:button|a)\b[^>]*(?:data-key|data-region|data-tab)=["'']기타지역["''][^>]*>)\s*부산\s*(</(?:button|a)>)',
    '$1기타지역$2'
)

# app.js 캐시 버전 갱신
$html = [regex]::Replace(
    $html,
    'assets/js/app\.js(?:\?v=[^"'']*)?',
    "assets/js/app.js?v=$stamp"
)

Set-Content -LiteralPath $index -Value $html -Encoding UTF8

# 최종 검증
$checkJs = Get-Content -LiteralPath $appJs -Raw -Encoding UTF8
$checkIndex = Get-Content -LiteralPath $index -Raw -Encoding UTF8

Write-Host ""
Write-Host "=== 지역별 인기 추천 v3 복구/재적용 완료 ==="
Write-Host "복구에 사용한 백업:" $restore.FullName
Write-Host "깨진 현재상태 백업:" $currentBackup
Write-Host ""

foreach($cat in $groupKeys){
    Write-Host "[$cat] $($groups[$cat].Count)개"
    $groups[$cat] | ForEach-Object {
        Write-Host " -" $_.name "|" $_.area "|" $_.price
    }
    Write-Host ""
}

Write-Host "서울 4개 확인:" ($groups["서울"].Count -eq 4)
Write-Host "경기 4개 확인:" ($groups["경기"].Count -eq 4)
Write-Host "인천 4개 확인:" ($groups["인천"].Count -eq 4)
Write-Host "기타지역 4개 확인:" ($groups["기타지역"].Count -eq 4)
Write-Host "객체형 렌더링 확인:" ($checkJs -match 'x\.name' -and $checkJs -match 'x\.area' -and $checkJs -match 'x\.price')
Write-Host "undefined 문자열 남음:" ($checkJs -match 'undefined')
Write-Host "부산 탭 남음:" ($checkIndex -match '>부산<')
Write-Host "기타지역 탭 확인:" ($checkIndex -match '>기타지역<')
Write-Host "app.js 캐시 갱신:" ($checkIndex -match [regex]::Escape("app.js?v=$stamp"))
Write-Host ""
Write-Host "이제 5512 메인에서 Ctrl+F5 하세요."
