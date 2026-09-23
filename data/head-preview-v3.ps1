$ErrorActionPreference = "Stop"

$root = "E:\mong-lounge-site"
$dataDir = Join-Path $root "data"
$stamp = Get-Date -Format "yyyyMMdd-HHmmss"
$out = Join-Path $dataDir "head-rewrite-preview-v3-$stamp.csv"

function Clean([string]$s){
    if([string]::IsNullOrWhiteSpace($s)){ return "" }
    $x=[System.Net.WebUtility]::HtmlDecode($s)
    $x=[regex]::Replace($x,'<[^>]+>',' ')
    $x=[regex]::Replace($x,'\s+',' ')
    return $x.Trim()
}

function Get-Title([string]$html){
    $m=[regex]::Match($html,'(?is)<title\b[^>]*>(.*?)</title>')
    if($m.Success){ return Clean $m.Groups[1].Value }
    return ""
}

function Get-Description([string]$html){
    $m=[regex]::Match($html,'(?is)<meta\b(?=[^>]*\bname\s*=\s*["'']description["''])(?=[^>]*\bcontent\s*=\s*["'']([^"'']*)["''])[^>]*>')
    if($m.Success){ return Clean $m.Groups[1].Value }
    $m=[regex]::Match($html,'(?is)<meta\b(?=[^>]*\bcontent\s*=\s*["'']([^"'']*)["''])(?=[^>]*\bname\s*=\s*["'']description["''])[^>]*>')
    if($m.Success){ return Clean $m.Groups[1].Value }
    return ""
}

function HashBytes([string]$s){
    $sha=[System.Security.Cryptography.SHA256]::Create()
    try { return $sha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($s)) }
    finally { $sha.Dispose() }
}

function Strip-RegionWords([string]$s){
    $x=Clean $s
    $x=$x -replace '\s*\|\s*몽라운지\s*$',''
    $x=$x -replace '\s*출장마사지\s*',' '
    $x=$x -replace '\s*생활권·지역\s*이용\s*안내\s*',' '
    $x=$x -replace '\s*생활권\s*안내\s*',' '
    $x=$x -replace '\s*지역\s*이용\s*안내\s*',' '
    $x=$x -replace '\s*지역\s*안내\s*',' '
    $x=$x -replace '\s*이용\s*안내\s*',' '
    $x=[regex]::Replace($x,'\s+',' ')
    return $x.Trim(' ','|','·','-')
}

function Parse-Shop([string]$title,[string]$rel){
    $parts=@($title -split '\s*\|\s*')
    $vendor=""
    $loc=""

    if($parts.Count -ge 2){
        $vendor=Clean $parts[0]
        $loc=Clean $parts[1]
        $loc=$loc -replace '\s*이용\s*안내\s*$',''
        $loc=$loc -replace '\s*상세\s*안내\s*$',''
        $loc=$loc.Trim()
    }

    if([string]::IsNullOrWhiteSpace($vendor)){
        $leaf=Split-Path (Split-Path ($rel.Replace('/','\')) -Parent) -Leaf
        $vendor=$leaf
    }
    if([string]::IsNullOrWhiteSpace($loc)){
        $loc="해당 지역"
    }

    return @($vendor,$loc)
}

function Parse-Region([string]$title,[string]$rel){
    $parts=@($title -split '\s*\|\s*')
    $leaf=""
    $parent=""

    if($parts.Count -ge 1){ $leaf=Strip-RegionWords $parts[0] }
    if($parts.Count -ge 2){ $parent=Strip-RegionWords $parts[1] }

    if([string]::IsNullOrWhiteSpace($leaf)){
        $folder=Split-Path (Split-Path ($rel.Replace('/','\')) -Parent) -Leaf
        $folder=[uri]::UnescapeDataString($folder)
        $leaf=($folder -split '-출장마사지|-생활권|-지역정보')[0]
        $leaf=$leaf -replace '-',' '
        $leaf=$leaf.Trim()
    }

    if([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $leaf){
        $parent=""
    }

    return @($leaf,$parent)
}

$shopTitles = @(
'{V} | {L} 코스·가격 안내 | 몽라운지',
'{L} {V} 영업시간·코스 안내 | 몽라운지',
'{V} {L} 이용정보 | 가격·문의 안내 | 몽라운지',
'{L}에서 확인하는 {V} | 코스와 운영시간 | 몽라운지',
'{V} 이용 안내 | {L} 가격·운영정보 | 몽라운지',
'{L} {V} | 전화·문자 문의와 코스 정보 | 몽라운지',
'{V} 정보 | {L} 코스·가격·영업시간 | 몽라운지',
'{L} 지역 {V} | 이용 전 확인사항 | 몽라운지',
'{V} | {L} 운영시간과 가격 정보 | 몽라운지',
'{L} {V} 이용 안내 | 코스·문의 정보 | 몽라운지',
'{V} {L} | 예약 전 코스·운영정보 | 몽라운지',
'{L}에서 보는 {V} | 가격과 연락 방법 | 몽라운지',
'{V} 코스 안내 | {L} 영업시간·문의 | 몽라운지',
'{L} {V} 상세 안내 | 코스·가격 확인 | 몽라운지',
'{V} | {L} 이용정보와 전화·문자 문의 | 몽라운지',
'{L} {V} | 운영시간·코스·가격 안내 | 몽라운지',
'{V} {L} 출장마사지 | 코스·가격 안내 | 몽라운지',
'{L} 출장마사지 {V} | 영업시간·문의 정보 | 몽라운지',
'{V} 출장마사지 | {L} 이용 전 확인정보 | 몽라운지',
'{L} {V} 출장마사지 안내 | 코스·운영시간 | 몽라운지',
'{V} 안내 | {L} 코스 구성과 가격 | 몽라운지',
'{L} {V} 정보 | 운영시간과 문의 방법 | 몽라운지',
'{V} {L} 이용 가이드 | 가격·코스 정보 | 몽라운지',
'{L} 기준 {V} | 영업시간·가격 안내 | 몽라운지',
'{V} | {L} 코스 정보와 연락 안내 | 몽라운지',
'{L} {V} 운영 안내 | 코스·가격·문의 | 몽라운지',
'{V} 상세정보 | {L} 이용시간·코스 안내 | 몽라운지',
'{L}에서 이용 전 확인하는 {V} | 몽라운지',
'{V} {L} 안내 | 전화·문자·가격 정보 | 몽라운지',
'{L} {V} | 코스별 시간과 가격 안내 | 몽라운지',
'{V} 이용정보 | {L} 운영시간·연락 방법 | 몽라운지',
'{L} {V} 가이드 | 코스·가격·이용정보 | 몽라운지'
)

$shopOpen = @(
'{L} 지역의 {V} 정보를 정리했습니다.',
'{V}의 {L} 이용 정보를 확인할 수 있는 페이지입니다.',
'{L}에서 {V} 이용 전 확인할 내용을 모았습니다.',
'{L} 기준으로 {V}의 운영 정보를 안내합니다.',
'{V}의 {L} 지역 안내 페이지입니다.',
'{L} 생활권에서 확인할 수 있는 {V} 정보를 정리했습니다.',
'{V} 이용을 검토할 때 참고할 {L} 지역 정보를 담았습니다.',
'{L}에서 확인하는 {V} 상세 안내입니다.',
'{V}의 코스와 운영 정보를 {L} 기준으로 정리했습니다.',
'{L} 지역 {V}의 기본 이용 정보를 한곳에 모았습니다.',
'{V} 관련 {L} 이용 정보를 보기 쉽게 정리했습니다.',
'{L}에서 {V} 이용 전 필요한 기본 정보를 확인해보세요.'
)

$shopBody = @(
'코스별 시간과 가격, 영업시간, 전화·문자 문의 방법을 함께 확인할 수 있습니다.',
'운영시간과 코스 구성, 가격, 연락 방법을 비교한 뒤 이용 여부를 결정할 수 있습니다.',
'코스와 가격 정보부터 영업시간, 전화 및 문자 문의 방법까지 순서대로 확인할 수 있습니다.',
'이용 전 코스별 시간·가격과 운영시간을 살펴보고 필요한 경우 전화나 문자로 문의할 수 있습니다.',
'가격과 코스 구성, 영업시간, 연락 수단을 한 페이지에서 확인할 수 있도록 구성했습니다.',
'코스별 이용시간과 가격, 운영시간 및 문의 방법을 확인하고 필요한 내용을 미리 비교할 수 있습니다.',
'영업시간과 코스·가격을 확인한 뒤 전화 또는 문자로 세부 이용 내용을 문의할 수 있습니다.',
'코스 구성과 가격, 운영시간, 전화·문자 연락 방법 등 실제 이용 전 필요한 정보를 담았습니다.',
'이용 가능한 코스와 가격, 영업시간, 문의 수단을 확인할 수 있어 예약 전 정보 비교에 활용할 수 있습니다.',
'운영시간, 코스별 가격과 시간, 전화·문자 문의 정보를 차례로 확인할 수 있습니다.',
'코스와 가격을 먼저 확인하고 영업시간과 연락 방법까지 함께 살펴볼 수 있도록 정리했습니다.',
'방문 전 확인하기 좋은 코스·가격·운영시간과 전화·문자 문의 정보를 제공합니다.'
)

$regionTitles = @(
'{A} 생활권 안내 | {P} 주변 지역정보 | 몽라운지',
'{A} 지역 이용정보 | {P} 인접 생활권 안내 | 몽라운지',
'{A} 출장마사지 | {P} 위치·이동 범위 안내 | 몽라운지',
'{P} {A} 생활권 | 주변 지역과 이용정보 | 몽라운지',
'{A} 위치 안내 | {P} 가까운 생활권 정보 | 몽라운지',
'{A} 지역 안내 | {P} 이동 동선과 주변 정보 | 몽라운지',
'{P} {A} 이용 가이드 | 생활권·위치정보 | 몽라운지',
'{A} 생활권 정보 | {P} 인접 지역 비교 | 몽라운지',
'{A} 출장마사지 지역안내 | {P} 생활권 정보 | 몽라운지',
'{P} {A} 지역정보 | 위치·주변 생활권 안내 | 몽라운지',
'{A} 이용 전 지역정보 | {P} 생활권 안내 | 몽라운지',
'{A} 주변 생활권 안내 | {P} 위치정보 | 몽라운지',
'{P} {A} 출장마사지 | 인접 지역·이동정보 | 몽라운지',
'{A} 지역 가이드 | {P} 생활권과 이동 범위 | 몽라운지',
'{A} 생활권·위치정보 | {P} 이용 안내 | 몽라운지',
'{P} {A} 안내 | 가까운 지역과 이동 동선 | 몽라운지',
'{A} 출장마사지 이용정보 | {P} 주변 지역 안내 | 몽라운지',
'{A} 지역정보 | {P} 생활권·인접 지역 확인 | 몽라운지',
'{P} {A} 생활권 안내 | 위치와 주변 지역 | 몽라운지',
'{A} 이용 안내 | {P} 지역 범위와 위치정보 | 몽라운지',
'{A} 주변 지역정보 | {P} 생활권 비교 | 몽라운지',
'{P} {A} 지역 이용정보 | 이동 범위 안내 | 몽라운지',
'{A} 출장마사지 생활권 | {P} 위치·주변 정보 | 몽라운지',
'{A} 위치·생활권 안내 | {P} 인접 지역정보 | 몽라운지'
)

$regionOpen = @(
'{A}을 중심으로 생활권과 주변 지역 정보를 정리했습니다.',
'{A} 주변의 위치 범위와 인접 지역을 확인할 수 있도록 정리했습니다.',
'{A} 생활권에서 이용 전 살펴볼 지역 정보를 안내합니다.',
'{P}의 {A}을 기준으로 가까운 생활권과 이동 범위를 정리했습니다.',
'{A}에서 주변 지역으로 이어지는 이동 동선과 생활권 정보를 담았습니다.',
'{A}의 위치 특성과 인접 생활권을 비교할 수 있도록 구성했습니다.',
'{A} 기준의 지역 범위와 가까운 동·읍·면 정보를 정리했습니다.',
'{A} 주변 생활권을 기준으로 위치와 이동 방향을 확인할 수 있습니다.',
'{P} {A}의 생활권과 인접 지역 정보를 한눈에 볼 수 있도록 정리했습니다.',
'{A} 이용 전 현재 위치와 주변 지역 범위를 비교할 수 있도록 안내합니다.',
'{A}에서 가까운 생활권과 주요 이동 방향을 중심으로 지역 정보를 담았습니다.',
'{A}을 기준으로 주변 지역과 연결되는 생활권 범위를 정리했습니다.'
)

$regionBody = @(
'가까운 지역과 이동 범위, 주변 생활권을 비교해 현재 위치에서 확인하기 좋은 지역을 찾는 데 참고할 수 있습니다.',
'인접 지역과 생활권 흐름을 함께 살펴보면 현재 위치에서 어느 방향이 가까운지 비교하는 데 도움이 됩니다.',
'주변 동·읍·면과의 연결 관계와 이동 범위를 확인해 지역 선택 전 참고 정보로 활용할 수 있습니다.',
'현재 위치를 기준으로 가까운 생활권과 주변 지역을 비교하고 이동 방향을 확인할 때 참고할 수 있습니다.',
'인접 지역의 위치 관계와 생활권 범위를 함께 정리해 주변 지역을 비교하기 쉽도록 했습니다.',
'지역별 위치와 이동 동선을 확인하고 가까운 생활권을 비교할 수 있도록 필요한 정보만 정리했습니다.',
'주변 지역과의 거리감, 이동 방향, 생활권 연결 정보를 살펴보고 이용 범위를 판단할 때 참고하세요.',
'가까운 생활권과 인접 지역을 한 번에 비교할 수 있어 위치 선택 전 지역 범위를 확인하기 좋습니다.',
'현재 지역에서 이어지는 주변 생활권과 이동 흐름을 살펴보고 가까운 범위를 비교할 수 있습니다.',
'지역의 위치 관계와 주변 생활권 정보를 확인해 이동 범위와 인접 지역을 판단할 때 활용할 수 있습니다.',
'주변 생활권과 연결되는 방향을 중심으로 정리해 현재 위치에서 가까운 지역을 비교하기 쉽습니다.',
'인접한 지역과 생활권 범위를 함께 확인할 수 있어 이동 동선과 위치 선택을 검토할 때 참고할 수 있습니다.'
)

$rows = New-Object System.Collections.Generic.List[object]

$files=@(
    Get-ChildItem -LiteralPath $root -Recurse -File -Filter "*.html" |
    Where-Object { $_.FullName -notmatch '\\\.git\\|\\node_modules\\|\\backup\\|\\_backup\\' }
)

foreach($f in $files){
    $rel=$f.FullName.Substring($root.Length).TrimStart("\").Replace("\","/")
    $class = if($rel -like "shops/*"){"업체상세"} elseif($rel -like "massage/*"){"지역SEO"} else {""}
    if(!$class){ continue }

    try{
        $html=Get-Content -LiteralPath $f.FullName -Raw -Encoding UTF8
        $oldTitle=Get-Title $html
        $oldMeta=Get-Description $html
        $h=HashBytes $rel

        if($class -eq "업체상세"){
            $p=Parse-Shop $oldTitle $rel
            $v=$p[0]; $l=$p[1]

            $ti=[int]$h[0] % $shopTitles.Count
            $oi=[int]$h[1] % $shopOpen.Count
            $bi=[int]$h[2] % $shopBody.Count

            $newTitle=$shopTitles[$ti].Replace('{V}',$v).Replace('{L}',$l)
            $newMeta=($shopOpen[$oi].Replace('{V}',$v).Replace('{L}',$l) + " " + $shopBody[$bi]).Trim()

            $rows.Add([pscustomobject]@{
                Class=$class; File=$rel
                OldTitle=$oldTitle; NewTitle=$newTitle
                OldDescription=$oldMeta; NewDescription=$newMeta
                TitlePattern="SHOP-T$('{0:D2}' -f ($ti+1))"
                MetaPattern="SHOP-M$('{0:D2}' -f ($oi+1))-$('{0:D2}' -f ($bi+1))"
            })
        }
        else{
            $p=Parse-Region $oldTitle $rel
            $a=$p[0]; $par=$p[1]
            if([string]::IsNullOrWhiteSpace($par)){ $par="주변 지역" }

            $ti=[int]$h[0] % $regionTitles.Count
            $oi=[int]$h[1] % $regionOpen.Count
            $bi=[int]$h[2] % $regionBody.Count

            $newTitle=$regionTitles[$ti].Replace('{A}',$a).Replace('{P}',$par)
            $newMeta=($regionOpen[$oi].Replace('{A}',$a).Replace('{P}',$par) + " " + $regionBody[$bi]).Trim()

            $rows.Add([pscustomobject]@{
                Class=$class; File=$rel
                OldTitle=$oldTitle; NewTitle=$newTitle
                OldDescription=$oldMeta; NewDescription=$newMeta
                TitlePattern="REG-T$('{0:D2}' -f ($ti+1))"
                MetaPattern="REG-M$('{0:D2}' -f ($oi+1))-$('{0:D2}' -f ($bi+1))"
            })
        }
    }
    catch{}
}

$rows | Export-Csv -LiteralPath $out -NoTypeInformation -Encoding UTF8

Write-Host ""
Write-Host "=== TITLE / META 재작성 v3 미리보기 ==="
Write-Host "사이트 파일 수정: 0"
Write-Host "전체 후보:" $rows.Count
Write-Host "지역SEO:" (@($rows|Where-Object Class -eq "지역SEO")).Count
Write-Host "업체상세:" (@($rows|Where-Object Class -eq "업체상세")).Count
Write-Host "지역 TITLE 패턴 사용 수:" (@($rows|Where-Object Class -eq "지역SEO"|Select-Object -ExpandProperty TitlePattern -Unique)).Count
Write-Host "지역 META 조합 사용 수:" (@($rows|Where-Object Class -eq "지역SEO"|Select-Object -ExpandProperty MetaPattern -Unique)).Count
Write-Host "업체 TITLE 패턴 사용 수:" (@($rows|Where-Object Class -eq "업체상세"|Select-Object -ExpandProperty TitlePattern -Unique)).Count
Write-Host "업체 META 조합 사용 수:" (@($rows|Where-Object Class -eq "업체상세"|Select-Object -ExpandProperty MetaPattern -Unique)).Count

Write-Host ""
Write-Host "=== 지역SEO 샘플 8개 ==="
$rows | Where-Object Class -eq "지역SEO" | Select-Object -First 8 | ForEach-Object {
    Write-Host "OLD TITLE:" $_.OldTitle
    Write-Host "NEW TITLE:" $_.NewTitle
    Write-Host "NEW META :" $_.NewDescription
    Write-Host ""
}

Write-Host "=== 업체상세 샘플 8개 ==="
$rows | Where-Object Class -eq "업체상세" | Select-Object -First 8 | ForEach-Object {
    Write-Host "OLD TITLE:" $_.OldTitle
    Write-Host "NEW TITLE:" $_.NewTitle
    Write-Host "NEW META :" $_.NewDescription
    Write-Host ""
}

Write-Host "미리보기 CSV:" $out
