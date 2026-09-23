$ErrorActionPreference = "Stop"

$root = "E:\mong-lounge-site"
$dataDir = Join-Path $root "data"
$stamp = Get-Date -Format "yyyyMMdd-HHmmss"
$out = Join-Path $dataDir "head-rewrite-preview-v5-$stamp.csv"

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

function HashIndex([string]$s,[int]$mod,[int]$offset=0){
    $sha=[System.Security.Cryptography.SHA256]::Create()
    try{
        $b=$sha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($s))
        return ([int]$b[$offset % $b.Length]) % $mod
    } finally { $sha.Dispose() }
}

function Extract-PrimaryRegion([string]$title){
    $x=$title
    $x=$x -replace '\s*\|\s*몽라운지\s*$',''
    $x=($x -split '\s*\|\s*')[0]
    if($x -match '^(.+?)\s+출장마사지(?:\s*-\s*.*)?$'){
        return $matches[1].Trim()
    }
    $x=$x -replace '\s+(?:생활권 안내|지역 이용 안내|지역 안내|이용 안내)\s*$',''
    $x=$x -replace '\s+출장마사지\s*$',''
    return $x.Trim()
}

function Extract-ShopParts([string]$title){
    $p=@($title -split '\s*\|\s*' | ForEach-Object { Clean $_ })
    if($p.Count -lt 2){ return $null }
    $vendor=$p[0].Trim()
    $loc=$p[1]
    $loc=$loc -replace '\s+(?:업체\s*)?상세정보\s*$',''
    $loc=$loc -replace '\s+이용\s*안내\s*$',''
    $loc=$loc.Trim()
    if(!$vendor -or !$loc){ return $null }
    return @($vendor,$loc)
}

function Extract-LocalAnchors([string]$html,[string]$primary){
    $set=New-Object System.Collections.Generic.HashSet[string]([StringComparer]::OrdinalIgnoreCase)
    $out=New-Object System.Collections.Generic.List[string]
    foreach($m in [regex]::Matches($html,'(?is)<a\b[^>]*>(.*?)</a>')){
        $t=Clean $m.Groups[1].Value
        $t=$t -replace '\s*(?:출장마사지|바로가기|전체보기|상세보기|업체보기|지역보기)\s*',' '
        $t=[regex]::Replace($t,'\s+',' ').Trim(' ','·','|','-')
        if(!$t -or $t.Length -gt 18){ continue }
        if($t -eq $primary){ continue }
        if($t -match '^(홈|메인|로그인|회원가입|FAQ|이용약관|개인정보처리방침|전화하기|문자하기)$'){ continue }
        if($t -notmatch '(동|읍|면|구|시|군|역)$'){ continue }
        if($set.Add($t)){
            $out.Add($t)
            if($out.Count -ge 4){ break }
        }
    }
    return @($out)
}

function Extract-ShopFacts([string]$html){
    $text=Clean $html
    $facts=New-Object System.Collections.Generic.List[string]
    $seen=New-Object System.Collections.Generic.HashSet[string]([StringComparer]::OrdinalIgnoreCase)

    foreach($m in [regex]::Matches($text,'(?<!\d)(?:30|40|50|60|70|80|90|100|120|150|180)\s*분')){
        $v=$m.Value -replace '\s+',''
        if($seen.Add($v)){ $facts.Add($v) }
        if($facts.Count -ge 2){ break }
    }

    foreach($m in [regex]::Matches($text,'(?<!\d)(?:\d{1,3}(?:,\d{3})+|\d{4,6})\s*원')){
        $v=$m.Value -replace '\s+',''
        if($seen.Add($v)){ $facts.Add($v) }
        if($facts.Count -ge 4){ break }
    }

    if($text -match '24\s*시간'){
        if($seen.Add('24시간')){ $facts.Add('24시간') }
    }

    return @($facts | Select-Object -First 4)
}

function Join-Nearby([string[]]$arr){
    if(!$arr -or $arr.Count -eq 0){ return "" }
    return ($arr | Select-Object -First 3) -join '·'
}

$regionTitlePatterns=@(
    '{R} 출장마사지 | {N} 주변 생활권 안내 | 몽라운지',
    '{R} 생활권 안내 | {N} 인접 지역정보 | 몽라운지',
    '{R} 지역정보 | {N} 위치·이동 범위 | 몽라운지',
    '{R} 출장마사지 지역안내 | {N} 주변 지역 | 몽라운지',
    '{R} 위치 안내 | {N} 생활권 정보 | 몽라운지',
    '{R} 이용 전 지역정보 | {N} 인접 생활권 | 몽라운지',
    '{R} 지역 가이드 | {N} 이동 동선 안내 | 몽라운지',
    '{R} 출장마사지 | {N} 지역 이용정보 | 몽라운지',
    '{R} 생활권 정보 | {N} 주변 위치 안내 | 몽라운지',
    '{R} 지역 안내 | {N} 가까운 생활권 | 몽라운지',
    '{R} 출장마사지 이용정보 | {N} 인접 지역 | 몽라운지',
    '{R} 위치·생활권 안내 | {N} 지역정보 | 몽라운지'
)

$regionMetaPatterns=@(
    '{R}을 기준으로 {N} 등 페이지에서 함께 확인할 수 있는 주변 생활권과 이동 범위를 정리했습니다. 현재 위치에서 가까운 지역을 비교하고 이용 범위를 확인할 때 참고할 수 있습니다.',
    '{R} 주변의 {N} 등 인접 지역 정보를 함께 확인할 수 있습니다. 생활권 연결 방향과 위치 범위를 비교해 현재 위치에서 가까운 지역을 살펴볼 때 참고하세요.',
    '{R} 지역에서 {N} 등 주변 생활권으로 이어지는 위치 정보를 정리했습니다. 인접 지역과 이동 방향을 비교하고 지역 범위를 확인하는 데 활용할 수 있습니다.',
    '{R} 생활권과 함께 {N} 등 연결된 지역 정보를 확인할 수 있도록 구성했습니다. 주변 위치와 이동 범위를 비교해 가까운 지역을 찾을 때 참고할 수 있습니다.',
    '{R} 위치를 중심으로 {N} 등 페이지에 연결된 주변 지역을 정리했습니다. 인접 생활권과 이동 방향을 한눈에 비교할 수 있도록 안내합니다.',
    '{R} 이용 전 {N} 등 주변 지역의 위치 관계를 확인할 수 있습니다. 가까운 생활권과 이동 범위를 비교해 지역 선택에 참고할 수 있도록 정리했습니다.'
)

$shopTitlePatterns=@(
    '{V} | {L} 코스·가격·운영정보 | 몽라운지',
    '{L} {V} | 코스와 영업시간 안내 | 몽라운지',
    '{V} {L} 이용정보 | 가격·전화·문자 문의 | 몽라운지',
    '{L}에서 확인하는 {V} | 코스·가격 안내 | 몽라운지',
    '{V} 이용 안내 | {L} 운영시간·코스 정보 | 몽라운지',
    '{L} {V} | 가격·운영시간·문의 정보 | 몽라운지',
    '{V} 상세정보 | {L} 코스와 이용시간 | 몽라운지',
    '{L} {V} 이용 가이드 | 코스·연락 방법 | 몽라운지',
    '{V} {L} 출장마사지 | 코스·가격 안내 | 몽라운지',
    '{L} 출장마사지 {V} | 영업시간·문의 | 몽라운지',
    '{V} | {L} 이용 전 확인정보 | 몽라운지',
    '{L} {V} | 코스별 시간·가격 안내 | 몽라운지',
    '{V} 운영정보 | {L} 가격·문의 방법 | 몽라운지',
    '{L} {V} | 영업시간과 코스 구성 | 몽라운지',
    '{V} {L} 안내 | 전화·문자·가격 정보 | 몽라운지',
    '{L} 기준 {V} | 코스·운영시간 안내 | 몽라운지'
)

$shopMetaPatterns=@(
    '{L} 지역의 {V} 이용 정보를 정리했습니다. {F} 코스와 가격, 영업시간, 전화·문자 문의 방법을 확인하고 이용 전 필요한 내용을 비교할 수 있습니다.',
    '{V}의 {L} 이용 안내입니다. {F} 운영시간과 코스 구성, 가격, 연락 방법을 한 페이지에서 확인할 수 있습니다.',
    '{L}에서 {V} 이용 전 확인할 정보를 모았습니다. {F} 코스별 시간과 가격, 운영시간, 전화·문자 문의 내용을 함께 살펴볼 수 있습니다.',
    '{V} 관련 {L} 정보를 확인할 수 있습니다. {F} 가격과 코스 구성, 영업시간 및 연락 수단을 비교해 이용 전 참고할 수 있습니다.',
    '{L} 기준으로 {V} 운영 정보를 정리했습니다. {F} 코스·가격·영업시간과 전화 또는 문자 문의 방법을 확인할 수 있습니다.',
    '{V} 이용을 검토할 때 참고할 {L} 상세 정보입니다. {F} 코스 구성과 가격, 운영시간, 연락 방법을 한곳에서 확인할 수 있습니다.',
    '{L} 생활권에서 확인할 수 있는 {V} 정보를 담았습니다. {F} 이용 가능한 코스와 가격, 영업시간, 전화·문자 문의 정보를 확인할 수 있습니다.',
    '{V}의 코스와 운영 정보를 {L} 기준으로 안내합니다. {F} 가격과 이용시간, 연락 방법을 확인해 필요한 내용을 미리 비교할 수 있습니다.'
)

$rows=New-Object System.Collections.Generic.List[object]
$regionContext=0
$shopFactsCount=0
$preservedLong=0

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

        if($class -eq "지역SEO"){
            if($oldTitle -match '\s출장마사지\s+-\s+' -or $oldTitle.Length -ge 90){
                $preservedLong++
                continue
            }

            $r=Extract-PrimaryRegion $oldTitle
            if(!$r){ continue }

            $near=Extract-LocalAnchors $html $r
            $nearText=Join-Nearby $near
            if(!$nearText){ $nearText="주변 지역"; }
            else { $regionContext++ }

            $ti=HashIndex $rel $regionTitlePatterns.Count 0
            $mi=HashIndex $rel $regionMetaPatterns.Count 1

            $newTitle=$regionTitlePatterns[$ti].Replace('{R}',$r).Replace('{N}',$nearText)
            $newMeta=$regionMetaPatterns[$mi].Replace('{R}',$r).Replace('{N}',$nearText)

            $rows.Add([pscustomobject]@{
                Class=$class; File=$rel; Context=$nearText
                OldTitle=$oldTitle; NewTitle=$newTitle
                OldDescription=$oldMeta; NewDescription=$newMeta
                TitlePattern="REG-$('{0:D2}' -f ($ti+1))"
                MetaPattern="REG-$('{0:D2}' -f ($mi+1))"
            })
        }
        else{
            $p=Extract-ShopParts $oldTitle
            if($null -eq $p){ continue }
            $v=$p[0]; $l=$p[1]

            $facts=Extract-ShopFacts $html
            $factText=""
            if($facts.Count -gt 0){
                $shopFactsCount++
                $factText=(($facts | Select-Object -First 3) -join '·') + " 정보를 포함해, "
            }

            $ti=HashIndex $rel $shopTitlePatterns.Count 0
            $mi=HashIndex $rel $shopMetaPatterns.Count 1

            $newTitle=$shopTitlePatterns[$ti].Replace('{V}',$v).Replace('{L}',$l)
            $newMeta=$shopMetaPatterns[$mi].Replace('{V}',$v).Replace('{L}',$l).Replace('{F}',$factText)

            $rows.Add([pscustomobject]@{
                Class=$class; File=$rel; Context=if($facts.Count){$facts -join '·'}else{""}
                OldTitle=$oldTitle; NewTitle=$newTitle
                OldDescription=$oldMeta; NewDescription=$newMeta
                TitlePattern="SHOP-$('{0:D2}' -f ($ti+1))"
                MetaPattern="SHOP-$('{0:D2}' -f ($mi+1))"
            })
        }
    } catch {}
}

$rows | Export-Csv -LiteralPath $out -NoTypeInformation -Encoding UTF8

$region=@($rows|Where-Object Class -eq "지역SEO")
$shop=@($rows|Where-Object Class -eq "업체상세")

Write-Host ""
Write-Host "=== TITLE / META v5 실제 페이지정보 기반 미리보기 ==="
Write-Host "사이트 파일 수정: 0"
Write-Host "지역SEO 변경후보:" $region.Count
Write-Host "지역SEO 긴 고유제목 보존:" $preservedLong
Write-Host "주변 지역명 활용 페이지:" $regionContext
Write-Host "업체상세 변경후보:" $shop.Count
Write-Host "코스/가격/시간 정보 감지 업체페이지:" $shopFactsCount
Write-Host "지역 TITLE 패턴:" (@($region|Select-Object -ExpandProperty TitlePattern -Unique)).Count
Write-Host "업체 TITLE 패턴:" (@($shop|Select-Object -ExpandProperty TitlePattern -Unique)).Count
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
