$ErrorActionPreference = "Stop"

$root = "E:\mong-lounge-site"
$dataDir = Join-Path $root "data"
$stamp = Get-Date -Format "yyyyMMdd-HHmmss"
$out = Join-Path $dataDir "head-rewrite-preview-v7-$stamp.csv"

function Clean([string]$s){
    if([string]::IsNullOrWhiteSpace($s)){ return "" }
    $x=[System.Net.WebUtility]::HtmlDecode($s)
    $x=[regex]::Replace($x,'<script\b[^>]*>.*?</script>',' ','IgnoreCase,Singleline')
    $x=[regex]::Replace($x,'<style\b[^>]*>.*?</style>',' ','IgnoreCase,Singleline')
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
    foreach($p in @(
        '(?is)<meta\b(?=[^>]*\bname\s*=\s*["'']description["''])(?=[^>]*\bcontent\s*=\s*["'']([^"'']*)["''])[^>]*>',
        '(?is)<meta\b(?=[^>]*\bcontent\s*=\s*["'']([^"'']*)["''])(?=[^>]*\bname\s*=\s*["'']description["''])[^>]*>'
    )){
        $m=[regex]::Match($html,$p)
        if($m.Success){ return Clean $m.Groups[1].Value }
    }
    return ""
}

function HashBytes([string]$s){
    $sha=[System.Security.Cryptography.SHA256]::Create()
    try { return $sha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($s)) }
    finally { $sha.Dispose() }
}

function Extract-PrimaryRegion([string]$title){
    $x=$title -replace '\s*\|\s*몽라운지\s*$',''
    $first=($x -split '\s*\|\s*')[0]
    if($first -match '^(.+?)\s+출장마사지(?:\s*-\s*.*)?$'){ return $matches[1].Trim() }
    $first=$first -replace '\s+(생활권 안내|지역 이용 안내|지역 안내|이용 안내|위치 안내|지역정보)\s*$',''
    $first=$first -replace '\s+출장마사지\s*$',''
    return $first.Trim()
}

function Is-LongUniqueRegionTitle([string]$title){
    if($title -match '\s출장마사지\s+-\s+'){ return $true }
    if($title -match '방문\s*홈케어|오피스|프리랜서|육아|집안일|업무|피로|뻐근|무거움|통근|교대|미용|숙소|호텔'){ return $true }
    if($title.Length -ge 82){ return $true }
    return $false
}

function Extract-RegionAnchors([string]$html,[string]$primary,[string]$title){
    $want = if($title -match '25개\s*구'){ '구' }
            elseif($title -match '31개\s*시.?군'){ '시군' }
            elseif($title -match '동별|읍.?면.?동|생활권'){ '하위' }
            else { '하위' }

    $seen=New-Object System.Collections.Generic.HashSet[string]([StringComparer]::OrdinalIgnoreCase)
    $out=New-Object System.Collections.Generic.List[string]

    foreach($m in [regex]::Matches($html,'(?is)<a\b[^>]*>(.*?)</a>')){
        $t=Clean $m.Groups[1].Value
        $t=$t -replace '^[←→]\s*',''
        $t=$t -replace '\s*(출장마사지|바로가기|전체보기|상세보기|업체보기|지역보기)\s*',' '
        $t=[regex]::Replace($t,'\s+',' ').Trim(' ','·','|','-')

        if(!$t -or $t.Length -gt 14){ continue }
        if($t -eq $primary){ continue }
        if($t -match '전체\s*지역|전지역|지역\s*전체|홈|메인|로그인|회원가입|FAQ|이용약관|개인정보|전화하기|문자하기'){ continue }
        if($t -match '^(서울|경기|인천|대전|대구|광주|부산|충남|충북|전남|전북|경남|경북)\s'){ continue }

        $ok=$false
        if($want -eq '구'){ $ok=($t -match '^[가-힣A-Za-z0-9]+구$') }
        elseif($want -eq '시군'){ $ok=($t -match '^[가-힣A-Za-z0-9]+(?:시|군)$') }
        else{ $ok=($t -match '^[가-힣A-Za-z0-9]+(?:동|읍|면|역)$') }

        if(!$ok){ continue }
        if($seen.Add($t)){
            $out.Add($t)
            if($out.Count -ge 4){ break }
        }
    }
    return @($out)
}

function Parse-Shop([string]$title){
    $p=@($title -split '\s*\|\s*' | ForEach-Object { Clean $_ })
    if($p.Count -lt 2){ return $null }
    $vendor=$p[0].Trim()
    $loc=$p[1]
    $loc=$loc -replace '\s+(업체\s*)?상세정보\s*$',''
    $loc=$loc -replace '\s+이용\s*안내\s*$',''
    $loc=$loc.Trim()
    if(!$vendor -or !$loc){ return $null }
    return @($vendor,$loc)
}

function Extract-ShopFacts([string]$html){
    $text=Clean $html
    $dur=@()
    foreach($m in [regex]::Matches($text,'(?<!\d)(\d{2,3})\s*분')){
        $n=[int]$m.Groups[1].Value
        if($n -ge 30 -and $n -le 240 -and $dur -notcontains $n){ $dur += $n }
        if($dur.Count -ge 3){ break }
    }

    $prices=@()
    foreach($m in [regex]::Matches($text,'(?<!\d)(\d{1,3}(?:,\d{3})+|\d{4,6})\s*원')){
        $raw=($m.Groups[1].Value -replace ',','')
        $n=0
        if([int]::TryParse($raw,[ref]$n)){
            if($n -ge 10000 -and $n -le 1000000 -and $prices -notcontains $n){ $prices += $n }
        }
        if($prices.Count -ge 5){ break }
    }

    $durText=""
    if($dur.Count -gt 0){
        $durText=(($dur | Sort-Object | Select-Object -First 2 | ForEach-Object {"${_}분"}) -join '·')
    }

    $priceText=""
    if($prices.Count -gt 0){
        $min=($prices | Measure-Object -Minimum).Minimum
        $priceText=("{0:N0}원부터" -f $min)
    }
    return @($durText,$priceText)
}

function Join-Nearby([string[]]$arr){
    if(!$arr -or $arr.Count -eq 0){ return "주변 지역" }
    return (($arr | Select-Object -First 3) -join '·')
}

$regionTitleOpen=@(
'{R} 출장마사지',
'{R} 생활권 안내',
'{R} 지역정보',
'{R} 위치 안내',
'{R} 이용 전 지역정보',
'{R} 지역 가이드',
'{R} 출장마사지 지역안내',
'{R} 생활권 정보',
'{R} 위치·생활권 안내',
'{R} 출장마사지 이용정보',
'{R} 주변 지역 안내',
'{R} 지역 이용 가이드'
)

$regionTitleClose=@(
'{N} 주변 생활권',
'{N} 인접 지역정보',
'{N} 위치·이동 범위',
'{N} 가까운 지역',
'{N} 생활권 정보',
'{N} 이동 동선',
'{N} 주변 위치',
'{N} 인접 생활권',
'{N} 지역 이용정보',
'{N} 생활권 비교'
)

$regionMetaOpen=@(
'{R} 지역에서 {N} 등 주변 생활권을 함께 확인할 수 있습니다.',
'{R} 기준으로 {N} 등 인접 지역 정보를 정리했습니다.',
'{R} 주변의 {N} 등 연결 지역을 중심으로 위치 정보를 구성했습니다.',
'{R} 생활권과 {N} 등 가까운 지역의 관계를 정리했습니다.',
'{R} 위치를 기준으로 {N} 등 주변 지역과의 연결 관계를 확인할 수 있습니다.',
'{R} 지역과 {N} 등 인접 생활권을 함께 살펴볼 수 있도록 구성했습니다.',
'{R}에서 {N} 등 주변 지역으로 이어지는 범위를 확인할 수 있습니다.',
'{R} 생활권을 중심으로 {N} 등 가까운 지역 정보를 정리했습니다.',
'{R}의 위치 특성과 {N} 등 인접 지역의 연결 방향을 함께 안내합니다.',
'{R} 지역에서 {N} 등 주변 생활권의 위치 관계를 확인할 수 있습니다.',
'{R} 기준 생활권과 {N} 등 인접 지역의 범위를 함께 정리했습니다.',
'{R} 주변 지역을 확인할 때 {N} 등 가까운 생활권을 비교할 수 있습니다.'
)

$regionMetaClose=@(
'현재 위치에서 가까운 지역을 비교하고 이동 범위를 살펴볼 때 참고할 수 있습니다.',
'생활권 연결 방향과 위치 범위를 비교해 주변 지역을 확인할 때 활용할 수 있습니다.',
'인접 생활권과 이동 방향을 비교하고 지역 범위를 확인하는 데 참고할 수 있습니다.',
'주변 위치와 이동 동선을 한눈에 비교할 수 있도록 필요한 정보만 담았습니다.',
'가까운 생활권과 이동 범위를 비교해 지역 선택 전 참고할 수 있습니다.',
'현재 위치에서 이어지는 주변 지역을 살펴보고 이동 방향을 판단할 때 활용할 수 있습니다.',
'인접 지역의 위치 관계와 생활권 범위를 비교할 때 참고할 수 있습니다.',
'주변 동·읍·면과의 연결 범위를 살펴보고 가까운 지역을 찾는 데 활용할 수 있습니다.',
'생활권별 위치와 이동 흐름을 비교해 주변 지역을 확인하기 쉽도록 구성했습니다.',
'가까운 지역과 연결 방향을 비교해 현재 위치 기준 생활권을 살펴볼 수 있습니다.',
'지역별 위치 관계와 이동 범위를 함께 비교할 때 참고할 수 있습니다.',
'인접 생활권과 주변 위치를 확인해 이동 동선을 검토할 때 활용할 수 있습니다.'
)

$shopTitleOpen=@(
'{V}',
'{L} {V}',
'{V} {L}',
'{L}에서 확인하는 {V}',
'{V} 이용 안내',
'{L} 기준 {V}',
'{V} 상세정보',
'{L} {V} 이용 가이드',
'{V} 운영정보',
'{V} 코스 안내',
'{L} 출장마사지 {V}',
'{V} {L} 출장마사지'
)

$shopTitleClose=@(
'코스·가격·운영정보',
'코스와 영업시간 안내',
'가격·전화·문자 문의',
'코스·가격 안내',
'운영시간·코스 정보',
'가격·운영시간·문의 정보',
'코스와 이용시간',
'코스·연락 방법',
'이용 전 확인정보',
'코스별 시간·가격 안내',
'가격·문의 방법',
'영업시간과 코스 구성',
'전화·문자·가격 정보',
'코스·운영시간 안내',
'코스 구성과 이용정보',
'가격·코스·연락 안내'
)

$shopMetaOpen=@(
'{L} 지역의 {V} 이용 정보를 정리했습니다.',
'{V}의 {L} 이용 안내입니다.',
'{L}에서 {V} 이용 전 확인할 정보를 모았습니다.',
'{V} 관련 {L} 정보를 확인할 수 있습니다.',
'{L} 기준으로 {V} 운영 정보를 정리했습니다.',
'{V} 이용을 검토할 때 참고할 {L} 상세 정보입니다.',
'{L} 생활권에서 확인할 수 있는 {V} 정보를 담았습니다.',
'{V}의 코스와 운영 정보를 {L} 기준으로 안내합니다.',
'{L}에서 이용 가능한 {V} 정보를 보기 쉽게 정리했습니다.',
'{V} 이용 전 살펴볼 {L} 지역 정보를 한곳에 모았습니다.',
'{L} 기준 {V}의 코스와 이용 정보를 확인할 수 있습니다.',
'{V}의 {L} 운영 정보와 문의 방법을 정리했습니다.'
)

$shopFactFormats=@(
'{D} 코스와 {P} 확인되는 가격 정보를 포함해 ',
'{P} 시작 가격과 {D} 코스 정보를 바탕으로 ',
'{D} 이용시간과 {P} 가격 정보를 함께 확인하며 ',
'{D} 코스 구성과 {P} 가격대를 확인한 뒤 ',
'{P}부터 확인되는 가격과 {D} 코스 정보를 참고해 ',
'{D} 코스·{P} 가격 정보를 포함해 ',
'{D} 시간 구성과 {P} 시작 가격을 확인하면서 ',
'{P} 가격 정보와 {D} 코스 시간을 함께 살펴보고 '
)

$shopMetaClose=@(
'코스와 가격, 영업시간, 전화·문자 문의 방법을 함께 확인할 수 있습니다.',
'운영시간과 코스 구성, 가격, 연락 방법을 한 페이지에서 확인할 수 있습니다.',
'코스별 시간과 가격, 운영시간, 전화·문자 문의 내용을 살펴볼 수 있습니다.',
'가격과 코스 구성, 영업시간 및 연락 수단을 비교해 이용 전 참고할 수 있습니다.',
'코스·가격·영업시간과 전화 또는 문자 문의 방법을 확인할 수 있습니다.',
'코스 구성과 가격, 운영시간, 연락 방법을 한곳에서 확인할 수 있습니다.',
'이용 가능한 코스와 가격, 영업시간, 전화·문자 문의 정보를 확인할 수 있습니다.',
'가격과 이용시간, 연락 방법을 확인해 필요한 내용을 미리 비교할 수 있습니다.',
'영업시간과 코스별 가격을 확인하고 전화나 문자로 세부 내용을 문의할 수 있습니다.',
'코스별 이용시간과 가격, 운영시간 및 연락 방법을 비교할 수 있습니다.',
'가격, 코스 구성, 영업시간과 문의 수단을 순서대로 확인할 수 있습니다.',
'이용 전 필요한 코스·가격·운영시간과 연락 정보를 한 번에 살펴볼 수 있습니다.'
)

$rows=New-Object System.Collections.Generic.List[object]
$preserved=0
$regionContext=0
$shopFacts=0

$files=@(
    Get-ChildItem -LiteralPath $root -Recurse -File -Filter "*.html" |
    Where-Object { $_.FullName -notmatch '\\\.git\\|\\node_modules\\|\\backup\\|\\_backup\\' }
)

foreach($f in $files){
    $rel=$f.FullName.Substring($root.Length).TrimStart("\").Replace("\","/")
    $class=if($rel -like "shops/*"){"업체상세"} elseif($rel -like "massage/*"){"지역SEO"} else {""}
    if(!$class){ continue }

    try{
        $html=Get-Content -LiteralPath $f.FullName -Raw -Encoding UTF8
        $oldTitle=Get-Title $html
        $oldMeta=Get-Description $html
        $h=HashBytes $rel

        if($class -eq "지역SEO"){
            if(Is-LongUniqueRegionTitle $oldTitle){
                $preserved++
                continue
            }

            $r=Extract-PrimaryRegion $oldTitle
            if(!$r){ continue }

            $near=Extract-RegionAnchors $html $r $oldTitle
            $n=Join-Nearby $near
            if($n -ne "주변 지역"){ $regionContext++ }

            $a=[int]$h[0] % $regionTitleOpen.Count
            $b=[int]$h[1] % $regionTitleClose.Count
            $c=[int]$h[2] % $regionMetaOpen.Count
            $d=[int]$h[3] % $regionMetaClose.Count

            $newTitle=$regionTitleOpen[$a].Replace('{R}',$r) + ' | ' + $regionTitleClose[$b].Replace('{N}',$n) + ' | 몽라운지'
            $newMeta=$regionMetaOpen[$c].Replace('{R}',$r).Replace('{N}',$n) + ' ' + $regionMetaClose[$d]

            $rows.Add([pscustomobject]@{
                Class=$class; File=$rel; Context=$n
                OldTitle=$oldTitle; NewTitle=$newTitle
                OldDescription=$oldMeta; NewDescription=$newMeta
                TitlePattern="R-$('{0:D2}' -f ($a+1))-$('{0:D2}' -f ($b+1))"
                MetaPattern="R-$('{0:D2}' -f ($c+1))-$('{0:D2}' -f ($d+1))"
            })
        }
        else{
            $p=Parse-Shop $oldTitle
            if($null -eq $p){ continue }
            $v=$p[0]; $l=$p[1]

            $facts=Extract-ShopFacts $html
            $dur=$facts[0]; $price=$facts[1]

            $a=[int]$h[0] % $shopTitleOpen.Count
            $b=[int]$h[1] % $shopTitleClose.Count
            $c=[int]$h[2] % $shopMetaOpen.Count
            $d=[int]$h[3] % $shopMetaClose.Count
            $e=[int]$h[4] % $shopFactFormats.Count

            $fact=""
            if($dur -and $price){
                $fact=$shopFactFormats[$e].Replace('{D}',$dur).Replace('{P}',$price)
                $shopFacts++
            } elseif($dur){
                $fact="$dur 코스 정보를 포함해 "
                $shopFacts++
            } elseif($price){
                $fact="$price 가격 정보를 포함해 "
                $shopFacts++
            }

            $newTitle=$shopTitleOpen[$a].Replace('{V}',$v).Replace('{L}',$l) + ' | ' + $shopTitleClose[$b] + ' | 몽라운지'
            $newMeta=$shopMetaOpen[$c].Replace('{V}',$v).Replace('{L}',$l) + ' ' + $fact + $shopMetaClose[$d]

            $rows.Add([pscustomobject]@{
                Class=$class; File=$rel; Context=("$dur $price").Trim()
                OldTitle=$oldTitle; NewTitle=$newTitle
                OldDescription=$oldMeta; NewDescription=$newMeta
                TitlePattern="S-$('{0:D2}' -f ($a+1))-$('{0:D2}' -f ($b+1))"
                MetaPattern="S-$('{0:D2}' -f ($c+1))-$('{0:D2}' -f ($e+1))-$('{0:D2}' -f ($d+1))"
            })
        }
    } catch {}
}

$rows | Export-Csv -LiteralPath $out -NoTypeInformation -Encoding UTF8

$region=@($rows|Where-Object Class -eq "지역SEO")
$shop=@($rows|Where-Object Class -eq "업체상세")

Write-Host ""
Write-Host "=== TITLE / META v7 고분산 미리보기 ==="
Write-Host "사이트 파일 수정: 0"
Write-Host "지역SEO 변경후보:" $region.Count
Write-Host "지역SEO 긴 고유제목 보존:" $preserved
Write-Host "정상 주변지역명 활용 페이지:" $regionContext
Write-Host "업체상세 변경후보:" $shop.Count
Write-Host "코스/가격 정보 활용 페이지:" $shopFacts
Write-Host "실제 지역 TITLE 조합 수:" (@($region|Select-Object -ExpandProperty TitlePattern -Unique)).Count
Write-Host "실제 지역 META 조합 수:" (@($region|Select-Object -ExpandProperty MetaPattern -Unique)).Count
Write-Host "실제 업체 TITLE 조합 수:" (@($shop|Select-Object -ExpandProperty TitlePattern -Unique)).Count
Write-Host "실제 업체 META 조합 수:" (@($shop|Select-Object -ExpandProperty MetaPattern -Unique)).Count
Write-Host ""

Write-Host "=== 지역SEO 샘플 12개 ==="
$region | Select-Object -First 12 | ForEach-Object {
    Write-Host "CONTEXT  :" $_.Context
    Write-Host "OLD TITLE:" $_.OldTitle
    Write-Host "NEW TITLE:" $_.NewTitle
    Write-Host "NEW META :" $_.NewDescription
    Write-Host ""
}

Write-Host "=== 업체상세 샘플 12개 ==="
$shop | Select-Object -First 12 | ForEach-Object {
    Write-Host "CONTEXT  :" $_.Context
    Write-Host "OLD TITLE:" $_.OldTitle
    Write-Host "NEW TITLE:" $_.NewTitle
    Write-Host "NEW META :" $_.NewDescription
    Write-Host ""
}

Write-Host "미리보기 CSV:" $out
