$ErrorActionPreference = "Stop"

$root = "E:\mong-lounge-site"
$dataDir = Join-Path $root "data"
$stamp = Get-Date -Format "yyyyMMdd-HHmmss"

$missingReport = Join-Path $dataDir "missing-description-classification-$stamp.csv"
$grammarReport = Join-Path $dataDir "meta-grammar-check-$stamp.csv"
$shopOgReport  = Join-Path $dataDir "shop-og-missing-$stamp.csv"

function Clean([string]$s){
    if([string]::IsNullOrWhiteSpace($s)){ return "" }
    $s=[System.Net.WebUtility]::HtmlDecode($s)
    $s=[regex]::Replace($s,'\s+',' ')
    return $s.Trim()
}

function Get-Title([string]$html){
    $m=[regex]::Match($html,'(?is)<title\b[^>]*>(.*?)</title>')
    if($m.Success){ return Clean $m.Groups[1].Value }
    return ""
}

function Get-Meta([string]$html,[string]$attr,[string]$name){
    foreach($p in @(
        "(?is)<meta\b(?=[^>]*\b$attr\s*=\s*[""']$([regex]::Escape($name))[""'])(?=[^>]*\bcontent\s*=\s*[""']([^""']*)[""'])[^>]*>",
        "(?is)<meta\b(?=[^>]*\bcontent\s*=\s*[""']([^""']*)[""'])(?=[^>]*\b$attr\s*=\s*[""']$([regex]::Escape($name))[""'])[^>]*>"
    )){
        $m=[regex]::Match($html,$p)
        if($m.Success){ return Clean $m.Groups[1].Value }
    }
    return ""
}

function Get-Canonical([string]$html){
    foreach($p in @(
        '(?is)<link\b(?=[^>]*\brel\s*=\s*["'']canonical["''])(?=[^>]*\bhref\s*=\s*["'']([^"'']+)["''])[^>]*>',
        '(?is)<link\b(?=[^>]*\bhref\s*=\s*["'']([^"'']+)["''])(?=[^>]*\brel\s*=\s*["'']canonical["''])[^>]*>'
    )){
        $m=[regex]::Match($html,$p)
        if($m.Success){ return $m.Groups[1].Value.Trim() }
    }
    return ""
}

function Classify-Missing([string]$rel,[string]$title,[string]$html){
    $name=[IO.Path]::GetFileName($rel).ToLowerInvariant()

    if($rel -match '(^|/)(blank-|login\.html$|signup\.html$|register\.html$|mypage\.html$|privacy\.html$|terms\.html$|search\.html$|notice\.html$|partner\.html$|event\.html$|dong-check\.html$)'){
        return "기능/보조페이지"
    }

    if($title -match '상세페이지 이동'){
        return "이동/연결페이지"
    }

    if($html -match '(?is)<meta[^>]+http-equiv\s*=\s*["'']refresh["'']'){
        return "리다이렉트페이지"
    }

    if($rel -like 'shops/*'){
        return "업체상세"
    }

    if($rel -like 'massage/*'){
        return "지역SEO페이지"
    }

    return "기타"
}

$all = @(Get-ChildItem -LiteralPath $root -Recurse -File -Filter "*.html" |
    Where-Object { $_.FullName -notmatch '\\\.git\\|\\node_modules\\|\\backup\\|\\_backup\\' })

$missing = New-Object System.Collections.Generic.List[object]
$grammar = New-Object System.Collections.Generic.List[object]
$shopOg  = New-Object System.Collections.Generic.List[object]

foreach($f in $all){
    $rel=$f.FullName.Substring($root.Length).TrimStart("\").Replace("\","/")
    $html=Get-Content -LiteralPath $f.FullName -Raw -Encoding UTF8

    $title=Get-Title $html
    $desc=Get-Meta $html "name" "description"
    $ogTitle=Get-Meta $html "property" "og:title"
    $ogDesc=Get-Meta $html "property" "og:description"
    $canonical=Get-Canonical $html
    $noindex=($html -match '(?is)<meta[^>]+name\s*=\s*["'']robots["''][^>]+content\s*=\s*["''][^"'']*noindex')

    if([string]::IsNullOrWhiteSpace($desc)){
        $missing.Add([pscustomobject]@{
            File=$rel
            Title=$title
            Class=Classify-Missing $rel $title $html
            Noindex=$noindex
            HasCanonical=(-not [string]::IsNullOrWhiteSpace($canonical))
            HasOGTitle=(-not [string]::IsNullOrWhiteSpace($ogTitle))
            HasOGDescription=(-not [string]::IsNullOrWhiteSpace($ogDesc))
        })
    }

    # 현재 확인된 대표 문법 이상 + 자주 생기는 조사 오류만 읽기 전용 검사
    foreach($pattern in @(
        '피로을',
        '통증을을',
        '긴장을을',
        '붓기을',
        '피로가을',
        '불편을을'
    )){
        if($title -match [regex]::Escape($pattern) -or $desc -match [regex]::Escape($pattern)){
            $grammar.Add([pscustomobject]@{
                File=$rel
                Pattern=$pattern
                Title=$title
                Description=$desc
            })
        }
    }

    if($rel -like 'shops/*'){
        if([string]::IsNullOrWhiteSpace($ogTitle) -or [string]::IsNullOrWhiteSpace($ogDesc)){
            $shopOg.Add([pscustomobject]@{
                File=$rel
                Title=$title
                Description=$desc
                OGTitleMissing=[string]::IsNullOrWhiteSpace($ogTitle)
                OGDescriptionMissing=[string]::IsNullOrWhiteSpace($ogDesc)
            })
        }
    }
}

$missing | Export-Csv -LiteralPath $missingReport -NoTypeInformation -Encoding UTF8
$grammar | Export-Csv -LiteralPath $grammarReport -NoTypeInformation -Encoding UTF8
$shopOg  | Export-Csv -LiteralPath $shopOgReport  -NoTypeInformation -Encoding UTF8

Write-Host ""
Write-Host "=== HEAD 후속 읽기 전용 점검 완료 ==="
Write-Host "사이트 파일 수정: 0"
Write-Host "전체 HTML:" $all.Count
Write-Host "DESCRIPTION 누락:" $missing.Count
Write-Host "업체상세 OG 누락:" $shopOg.Count
Write-Host "조사/문법 의심 페이지:" $grammar.Count
Write-Host ""

Write-Host "=== DESCRIPTION 누락 분류 ==="
$missing | Group-Object Class | Sort-Object Count -Descending | ForEach-Object {
    Write-Host "$($_.Name): $($_.Count)"
}

Write-Host ""
Write-Host "=== 업체상세 OG 누락 샘플 5개 ==="
$shopOg | Select-Object -First 5 | ForEach-Object {
    Write-Host $_.File
    Write-Host " TITLE:" $_.Title
    Write-Host " META :" $_.Description
    Write-Host "---"
}

Write-Host ""
Write-Host "=== 문법 의심 샘플 10개 ==="
$grammar | Select-Object -First 10 | ForEach-Object {
    Write-Host $_.File
    Write-Host " 패턴:" $_.Pattern
    Write-Host " META :" $_.Description
    Write-Host "---"
}

Write-Host ""
Write-Host "누락 분류 보고서:" $missingReport
Write-Host "문법 점검 보고서:" $grammarReport
Write-Host "업체 OG 보고서:" $shopOgReport
