$ErrorActionPreference = "Stop"

$root = "E:\mong-lounge-site"
$dataDir = Join-Path $root "data"
$stamp = Get-Date -Format "yyyyMMdd-HHmmss"
$out = Join-Path $dataDir "head-rewrite-preview-v4-$stamp.csv"

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
function Fill([string]$s,[string]$a,[string]$b,[string]$c=""){
    return $s.Replace('{A}',$a).Replace('{B}',$b).Replace('{C}',$c)
}
function Strip-Suffix([string]$s){
    $x=Clean $s
    $x=$x -replace '\s*\|\s*몽라운지\s*$',''
    $x=$x -replace '\s*생활권·지역\s*이용\s*안내\s*$',''
    $x=$x -replace '\s*생활권\s*안내\s*$',''
    $x=$x -replace '\s*지역\s*이용\s*안내\s*$',''
    $x=$x -replace '\s*지역\s*안내\s*$',''
    $x=$x -replace '\s*이용\s*안내\s*$',''
    $x=$x -replace '\s*출장마사지\s*$',''
    return ($x -replace '\s+',' ').Trim(' ','|','·','-')
}

function Parse-RegionSimple([string]$title){
    # 긴 시나리오형 제목은 유지: "지역 출장마사지 - ..."
    if($title -match '\s출장마사지\s+-\s+' -or $title.Length -ge 90){
        return $null
    }

    $parts=@($title -split '\s*\|\s*' | ForEach-Object { Clean $_ })
    if($parts.Count -lt 2){ return $null }

    $first=Strip-Suffix $parts[0]
    $second=Strip-Suffix $parts[1]

    if([string]::IsNullOrWhiteSpace($first)){ return $null }

    # 단순 지역 안내형만 변경
    $simple = (
        $title -match '출장마사지\s*\|' -or
        $title -match '생활권\s*안내\s*\|' -or
        $title -match '지역\s*이용\s*안내\s*\|' -or
        $title -match '생활권·지역\s*이용\s*안내'
    )
    if(!$simple){ return $null }

    if([string]::IsNullOrWhiteSpace($second) -or $second -eq "몽라운지"){
        $second="주변 지역"
    }

    return @($first,$second)
}

function Parse-Shop([string]$title){
    $parts=@($title -split '\s*\|\s*' | ForEach-Object { Clean $_ })
    if($parts.Count -lt 2){ return $null }

    $vendor=$parts[0]
    $loc=$parts[1]
    $loc=$loc -replace '\s*이용\s*안내\s*$',''
    $loc=$loc -replace '\s*업체\s*상세정보\s*$',''
    $loc=$loc -replace '\s*상세정보\s*$',''
    $loc=$loc.Trim()

    if([string]::IsNullOrWhiteSpace($vendor) -or [string]::IsNullOrWhiteSpace($loc)){
        return $null
    }
    return @($vendor,$loc)
}

$regionTitles=@(
'{A} 생활권 안내 | {B} 주변 위치정보 | 몽라운지',
'{A} 지역정보 | {B} 인접 생활권 안내 | 몽라운지',
'{A} 출장마사지 | {B} 위치·이동 범위 | 몽라운지',
'{B} {A} 생활권 | 주변 지역정보 | 몽라운지',
'{A} 위치 안내 | {B} 가까운 지역정보 | 몽라운지',
'{A} 지역 가이드 | {B} 이동 범위 안내 | 몽라운지',
'{B} {A} 이용정보 | 생활권·위치 안내 | 몽라운지',
'{A} 생활권 정보 | {B} 인접 지역 비교 | 몽라운지',
'{A} 출장마사지 지역안내 | {B} 생활권 정보 | 몽라운지',
'{B} {A} 지역정보 | 위치·주변 생활권 | 몽라운지',
'{A} 이용 전 지역정보 | {B} 생활권 안내 | 몽라운지',
'{A} 주변 생활권 | {B} 위치정보 안내 | 몽라운지',
'{B} {A} 출장마사지 | 인접 지역·이동정보 | 몽라운지',
'{A} 지역 안내 | {B} 생활권과 이동 범위 | 몽라운지',
'{A} 생활권·위치정보 | {B} 주변 지역 안내 | 몽라운지',
'{B} {A} 안내 | 가까운 지역과 이동 동선 | 몽라운지',
'{A} 출장마사지 이용정보 | {B} 주변 지역 | 몽라운지',
'{A} 지역정보 | {B} 인접 생활권 확인 | 몽라운지'
)

$regionOpen=@(
'{A} 중심의 생활권과 주변 지역 정보를 정리했습니다.',
'{A} 주변 위치 범위와 인접 지역을 비교할 수 있도록 구성했습니다.',
'{A} 생활권에서 확인할 수 있는 지역 정보를 정리했습니다.',
'{B} 안에서 {A} 위치와 가까운 생활권을 확인할 수 있습니다.',
'{A} 기준으로 주변 지역과 이어지는 이동 범위를 정리했습니다.',
'{A} 위치와 인접 생활권을 한눈에 비교할 수 있도록 안내합니다.',
'{A} 주변의 동·읍·면과 연결되는 생활권 정보를 모았습니다.',
'{B} {A} 지역의 위치 관계와 주변 생활권을 정리했습니다.',
'{A} 이용 전 참고하기 좋은 주변 지역 정보를 담았습니다.',
'{A}에서 가까운 지역과 주요 이동 방향을 중심으로 정리했습니다.',
'{A} 주변 생활권과 지역 범위를 비교할 수 있도록 구성했습니다.',
'{B} 내 {A} 위치를 기준으로 인접 지역 정보를 정리했습니다.'
)

$regionBody=@(
'가까운 지역과 이동 범위를 비교해 현재 위치에서 확인하기 좋은 생활권을 찾을 때 참고할 수 있습니다.',
'인접 지역의 위치 관계와 이동 흐름을 함께 살펴보면 주변 생활권을 비교하기 쉽습니다.',
'주변 동·읍·면과 연결되는 방향을 확인해 지역 선택 전 참고 정보로 활용할 수 있습니다.',
'현재 위치를 기준으로 가까운 생활권과 주변 지역을 비교할 때 활용할 수 있습니다.',
'지역별 위치와 이동 동선을 확인하고 인접 생활권을 비교할 수 있도록 필요한 정보만 담았습니다.',
'주변 지역과의 연결 관계, 이동 방향, 생활권 범위를 살펴볼 때 참고할 수 있습니다.',
'가까운 생활권과 인접 지역을 함께 확인할 수 있어 위치 범위를 비교하기 좋습니다.',
'현재 지역에서 이어지는 주변 생활권과 이동 흐름을 확인할 수 있도록 정리했습니다.',
'지역의 위치 관계와 주변 생활권 정보를 확인해 이동 범위를 판단할 때 참고할 수 있습니다.',
'인접한 지역과 생활권 범위를 함께 확인해 주변 위치를 비교할 수 있습니다.',
'현재 위치에서 가까운 지역과 생활권 연결 방향을 살펴볼 때 활용할 수 있습니다.',
'주변 생활권과 지역 범위를 비교해 이동 동선과 위치 선택을 검토할 때 참고할 수 있습니다.'
)

$shopTitles=@(
'{A} | {B} 코스·가격 안내 | 몽라운지',
'{B} {A} | 영업시간·코스 정보 | 몽라운지',
'{A} {B} 이용정보 | 가격·문의 안내 | 몽라운지',
'{B} {A} | 코스와 운영시간 | 몽라운지',
'{A} 이용 안내 | {B} 가격·운영정보 | 몽라운지',
'{B} {A} | 전화·문자 문의·코스 정보 | 몽라운지',
'{A} 정보 | {B} 코스·가격·영업시간 | 몽라운지',
'{B} {A} | 이용 전 확인사항 | 몽라운지',
'{A} | {B} 운영시간·가격 정보 | 몽라운지',
'{B} {A} 이용 안내 | 코스·문의 정보 | 몽라운지',
'{A} {B} | 코스·가격 확인 | 몽라운지',
'{B} {A} | 가격과 연락 방법 | 몽라운지',
'{A} 코스 안내 | {B} 영업시간·문의 | 몽라운지',
'{B} {A} 상세 안내 | 코스·가격 | 몽라운지',
'{A} | {B} 이용정보·전화·문자 문의 | 몽라운지',
'{B} {A} | 운영시간·코스·가격 | 몽라운지',
'{A} {B} 출장마사지 | 코스·가격 안내 | 몽라운지',
'{B} 출장마사지 {A} | 영업시간·문의 | 몽라운지',
'{A} 출장마사지 | {B} 이용 전 확인정보 | 몽라운지',
'{B} {A} 출장마사지 | 코스·운영시간 | 몽라운지',
'{A} 안내 | {B} 코스 구성과 가격 | 몽라운지',
'{B} {A} 정보 | 운영시간과 문의 방법 | 몽라운지',
'{A} {B} 이용 가이드 | 가격·코스 | 몽라운지',
'{B} 기준 {A} | 영업시간·가격 안내 | 몽라운지',
'{A} | {B} 코스 정보·연락 안내 | 몽라운지',
'{B} {A} 운영 안내 | 코스·가격·문의 | 몽라운지',
'{A} 상세정보 | {B} 이용시간·코스 | 몽라운지',
'{B} {A} | 이용 전 확인정보 | 몽라운지',
'{A} {B} 안내 | 전화·문자·가격 정보 | 몽라운지',
'{B} {A} | 코스별 시간·가격 | 몽라운지',
'{A} 이용정보 | {B} 운영시간·연락 방법 | 몽라운지',
'{B} {A} 가이드 | 코스·가격·이용정보 | 몽라운지',
'{A} | {B} 영업시간과 코스 구성 | 몽라운지',
'{B} {A} | 가격·운영시간 확인 | 몽라운지',
'{A} {B} | 코스와 전화·문자 문의 | 몽라운지',
'{B}에서 확인하는 {A} | 코스·가격 | 몽라운지',
'{A} 운영정보 | {B} 가격·문의 방법 | 몽라운지',
'{B} {A} | 영업시간·이용정보 | 몽라운지',
'{A} {B} 코스정보 | 가격·운영시간 | 몽라운지',
'{B} {A} | 예약 전 확인정보 | 몽라운지'
)

$shopOpen=@(
'{B} 지역의 {A} 정보를 정리했습니다.',
'{A} 관련 {B} 이용 정보를 확인할 수 있습니다.',
'{B}에서 {A} 이용 전 확인할 내용을 모았습니다.',
'{B} 기준으로 {A} 운영 정보를 안내합니다.',
'{A}의 {B} 지역 이용 정보를 정리했습니다.',
'{B} 생활권에서 확인할 수 있는 {A} 정보를 담았습니다.',
'{A} 이용을 검토할 때 참고할 {B} 정보를 정리했습니다.',
'{B}에서 확인하는 {A} 상세 안내입니다.',
'{A} 코스와 운영 정보를 {B} 기준으로 정리했습니다.',
'{B} 지역 {A} 기본 이용 정보를 한곳에 모았습니다.',
'{A} 관련 {B} 이용 정보를 보기 쉽게 구성했습니다.',
'{B}에서 {A} 이용 전 필요한 기본 정보를 확인할 수 있습니다.',
'{A}의 운영 정보를 {B} 위치 기준으로 정리했습니다.',
'{B}에서 확인할 수 있는 {A} 이용 정보를 안내합니다.',
'{A} 이용 전 살펴볼 {B} 지역 정보를 정리했습니다.',
'{B} 기준 {A} 코스와 운영 정보를 확인할 수 있습니다.'
)

$shopBody=@(
'코스별 시간과 가격, 영업시간, 전화·문자 문의 방법을 함께 확인할 수 있습니다.',
'운영시간과 코스 구성, 가격, 연락 방법을 비교한 뒤 이용 여부를 결정할 수 있습니다.',
'코스와 가격 정보부터 영업시간, 전화 및 문자 문의 방법까지 순서대로 확인할 수 있습니다.',
'이용 전 코스별 시간·가격과 운영시간을 살펴보고 필요한 경우 전화나 문자로 문의할 수 있습니다.',
'가격과 코스 구성, 영업시간, 연락 수단을 한 페이지에서 확인할 수 있도록 구성했습니다.',
'코스별 이용시간과 가격, 운영시간 및 문의 방법을 확인할 수 있습니다.',
'영업시간과 코스·가격을 확인한 뒤 전화 또는 문자로 세부 내용을 문의할 수 있습니다.',
'코스 구성과 가격, 운영시간, 전화·문자 연락 방법 등 이용 전 필요한 정보를 담았습니다.',
'이용 가능한 코스와 가격, 영업시간, 문의 수단을 확인해 예약 전 비교에 활용할 수 있습니다.',
'운영시간, 코스별 가격과 시간, 전화·문자 문의 정보를 차례로 확인할 수 있습니다.',
'코스와 가격을 먼저 확인하고 영업시간과 연락 방법까지 함께 살펴볼 수 있습니다.',
'코스·가격·운영시간과 전화·문자 문의 정보를 한 번에 확인할 수 있습니다.',
'가격, 코스별 시간, 운영시간과 연락 방법을 비교하기 쉽게 정리했습니다.',
'이용 전 확인하기 좋은 코스 구성, 가격, 영업시간과 문의 방법을 제공합니다.',
'코스별 가격과 운영시간을 확인하고 전화나 문자로 필요한 내용을 문의할 수 있습니다.',
'영업시간과 가격, 코스 구성, 연락 수단을 한 페이지에서 확인할 수 있습니다.'
)

$rows=New-Object System.Collections.Generic.List[object]
$preservedRegion=0

$files=@(
    Get-ChildItem -LiteralPath $root -Recurse -File -Filter "*.html" |
    Where-Object { $_.FullName -notmatch '\\\.git\\|\\node_modules\\|\\backup\\|\\_backup\\' }
)

foreach($f in $files){
    $rel=$f.FullName.Substring($root.Length).TrimStart("\").Replace("\","/")
    $class=if($rel -like "shops/*"){"업체상세"} elseif($rel -like "massage/*"){"지역SEO"} else {""}
    if(!$class){continue}

    try{
        $html=Get-Content -LiteralPath $f.FullName -Raw -Encoding UTF8
        $oldTitle=Get-Title $html
        $oldMeta=Get-Description $html
        $h=HashBytes $rel

        if($class -eq "지역SEO"){
            $p=Parse-RegionSimple $oldTitle
            if($null -eq $p){
                $preservedRegion++
                continue
            }
            $a=$p[0]; $b=$p[1]
            $ti=[int]$h[0] % $regionTitles.Count
            $oi=[int]$h[1] % $regionOpen.Count
            $bi=[int]$h[2] % $regionBody.Count

            $newTitle=Fill $regionTitles[$ti] $a $b
            $newMeta=(Fill $regionOpen[$oi] $a $b) + " " + $regionBody[$bi]

            $rows.Add([pscustomobject]@{
                Class=$class; File=$rel; Action="변경후보"
                OldTitle=$oldTitle; NewTitle=$newTitle
                OldDescription=$oldMeta; NewDescription=$newMeta
                TitlePattern="REG-T$('{0:D2}' -f ($ti+1))"
                MetaPattern="REG-M$('{0:D2}' -f ($oi+1))-$('{0:D2}' -f ($bi+1))"
            })
        }
        else{
            $p=Parse-Shop $oldTitle
            if($null -eq $p){continue}
            $a=$p[0]; $b=$p[1]
            $ti=[int]$h[0] % $shopTitles.Count
            $oi=[int]$h[1] % $shopOpen.Count
            $bi=[int]$h[2] % $shopBody.Count

            $newTitle=Fill $shopTitles[$ti] $a $b
            $newMeta=(Fill $shopOpen[$oi] $a $b) + " " + $shopBody[$bi]

            $rows.Add([pscustomobject]@{
                Class=$class; File=$rel; Action="변경후보"
                OldTitle=$oldTitle; NewTitle=$newTitle
                OldDescription=$oldMeta; NewDescription=$newMeta
                TitlePattern="SHOP-T$('{0:D2}' -f ($ti+1))"
                MetaPattern="SHOP-M$('{0:D2}' -f ($oi+1))-$('{0:D2}' -f ($bi+1))"
            })
        }
    } catch {}
}

$rows | Export-Csv -LiteralPath $out -NoTypeInformation -Encoding UTF8

$region=@($rows|Where-Object Class -eq "지역SEO")
$shop=@($rows|Where-Object Class -eq "업체상세")

Write-Host ""
Write-Host "=== TITLE / META 재작성 v4 미리보기 ==="
Write-Host "사이트 파일 수정: 0"
Write-Host "지역SEO 변경후보:" $region.Count
Write-Host "지역SEO 긴 고유제목 보존:" $preservedRegion
Write-Host "업체상세 변경후보:" $shop.Count
Write-Host "지역 TITLE 패턴 사용 수:" (@($region|Select-Object -ExpandProperty TitlePattern -Unique)).Count
Write-Host "지역 META 조합 사용 수:" (@($region|Select-Object -ExpandProperty MetaPattern -Unique)).Count
Write-Host "업체 TITLE 패턴 사용 수:" (@($shop|Select-Object -ExpandProperty TitlePattern -Unique)).Count
Write-Host "업체 META 조합 사용 수:" (@($shop|Select-Object -ExpandProperty MetaPattern -Unique)).Count
Write-Host "지역 NEW TITLE 출장마사지 포함:" (@($region|Where-Object {$_.NewTitle -match '출장마사지'})).Count
Write-Host "업체 NEW TITLE 출장마사지 포함:" (@($shop|Where-Object {$_.NewTitle -match '출장마사지'})).Count
Write-Host ""

Write-Host "=== 지역SEO 변경 샘플 10개 ==="
$region | Select-Object -First 10 | ForEach-Object {
    Write-Host "OLD TITLE:" $_.OldTitle
    Write-Host "NEW TITLE:" $_.NewTitle
    Write-Host "NEW META :" $_.NewDescription
    Write-Host ""
}

Write-Host "=== 업체상세 변경 샘플 10개 ==="
$shop | Select-Object -First 10 | ForEach-Object {
    Write-Host "OLD TITLE:" $_.OldTitle
    Write-Host "NEW TITLE:" $_.NewTitle
    Write-Host "NEW META :" $_.NewDescription
    Write-Host ""
}

Write-Host "미리보기 CSV:" $out
