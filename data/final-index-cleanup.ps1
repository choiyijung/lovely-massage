$ErrorActionPreference = "Stop"

$root = "E:\mong-lounge-site"
$baseUrl = "https://1004테라피.shop"
$stamp = Get-Date -Format "yyyyMMdd-HHmmss"
$backup = "E:\mong-lounge-final-index-cleanup-backup-$stamp"

$indexableFolders = @(
    "premium/index.html",
    "community/index.html",
    "blog/index.html",
    "terms/index.html",
    "region/index.html",
    "special/index.html",
    "subway/index.html",
    "privacy/index.html",
    "coupon/index.html",
    "vip/index.html",
    "faq/index.html",
    "event/index.html",
    "partner/index.html",
    "notice/index.html"
)

$noindexPages = @(
    "shop-songpa.html",
    "event.html",
    "privacy.html",
    "partner.html",
    "premium.html",
    "shop-gangnam.html",
    "subway.html",
    "blog.html",
    "terms.html",
    "notice.html",
    "community.html",
    "shop-suwon.html",
    "faq.html",
    "coupon.html",
    "vip.html",
    "special.html",
    "shop-bupyeong.html",
    "region.html",
    "data/shop-detail-template.html",
    "mypage/index.html",
    "join/index.html",
    "login/index.html"
)

$flatToClean = @{
    "event.html"     = "/event/"
    "privacy.html"   = "/privacy/"
    "partner.html"   = "/partner/"
    "premium.html"   = "/premium/"
    "subway.html"    = "/subway/"
    "blog.html"      = "/blog/"
    "terms.html"     = "/terms/"
    "notice.html"    = "/notice/"
    "community.html" = "/community/"
    "faq.html"       = "/faq/"
    "coupon.html"    = "/coupon/"
    "vip.html"       = "/vip/"
    "special.html"   = "/special/"
    "region.html"    = "/region/"
}

New-Item -ItemType Directory -Force -Path $backup | Out-Null

function Backup-File([string]$rel){
    $src = Join-Path $root ($rel.Replace('/','\'))
    if(!(Test-Path -LiteralPath $src)){ return }
    $dest = Join-Path $backup ($rel.Replace('/','\'))
    New-Item -ItemType Directory -Force -Path (Split-Path $dest -Parent) | Out-Null
    Copy-Item -LiteralPath $src -Destination $dest -Force
}

function Set-Noindex([string]$html){
    $pattern='(?is)<meta\b(?=[^>]*\bname\s*=\s*["'']robots["''])[^>]*>'
    $tag='<meta name="robots" content="noindex,follow">'
    if([regex]::IsMatch($html,$pattern)){
        return [regex]::Replace($html,$pattern,$tag,1)
    }
    return [regex]::Replace($html,'(?is)</head>',$tag + "`r`n</head>",1)
}

function Set-Canonical([string]$html,[string]$url){
    $pattern='(?is)\s*<link\b(?=[^>]*\brel\s*=\s*["'']canonical["''])[^>]*>\s*'
    $html=[regex]::Replace($html,$pattern,"`r`n")
    $tag='<link rel="canonical" href="' + $url + '">'
    if([regex]::IsMatch($html,'(?is)</title>')){
        return [regex]::Replace($html,'(?is)</title>','</title>' + "`r`n" + $tag,1)
    }
    return [regex]::Replace($html,'(?is)</head>',$tag + "`r`n</head>",1)
}

$modified = 0
$missing = New-Object System.Collections.Generic.List[string]

# 1) 공개 보조페이지 14개: index 유지 + self canonical
foreach($rel in $indexableFolders){
    $file = Join-Path $root ($rel.Replace('/','\'))
    if(!(Test-Path -LiteralPath $file)){
        $missing.Add($rel)
        continue
    }

    Backup-File $rel
    $html=Get-Content -LiteralPath $file -Raw -Encoding UTF8

    $path="/" + (($rel -replace '/index\.html$','').Trim('/')) + "/"
    $url=$baseUrl.TrimEnd('/') + $path

    $new=Set-Canonical $html $url

    # 혹시 robots noindex가 있었다면 공개페이지에서는 제거
    $new=[regex]::Replace(
        $new,
        '(?is)\s*<meta\b(?=[^>]*\bname\s*=\s*["'']robots["''])(?=[^>]*\bcontent\s*=\s*["''][^"'']*noindex[^"'']*["''])[^>]*>\s*',
        "`r`n"
    )

    if($new -ne $html){
        Set-Content -LiteralPath $file -Value $new -Encoding UTF8
        $modified++
    }
}

# 2) 중복/기능/템플릿 22개: noindex,follow
foreach($rel in $noindexPages){
    $file = Join-Path $root ($rel.Replace('/','\'))
    if(!(Test-Path -LiteralPath $file)){
        $missing.Add($rel)
        continue
    }

    Backup-File $rel
    $html=Get-Content -LiteralPath $file -Raw -Encoding UTF8
    $new=Set-Noindex $html

    # flat 중복은 clean 폴더 버전으로 canonical
    if($flatToClean.ContainsKey($rel)){
        $new=Set-Canonical $new ($baseUrl.TrimEnd('/') + $flatToClean[$rel])
    }

    if($new -ne $html){
        Set-Content -LiteralPath $file -Value $new -Encoding UTF8
        $modified++
    }
}

# 3) sitemap.xml: 기존 URL 유지 + 공개 보조페이지 14개 추가, noindex URL 제거
$sitemapPath=Join-Path $root "sitemap.xml"
if(!(Test-Path -LiteralPath $sitemapPath)){
    throw "sitemap.xml이 없습니다."
}
Backup-File "sitemap.xml"

$xml=Get-Content -LiteralPath $sitemapPath -Raw -Encoding UTF8

$urls=New-Object System.Collections.Generic.HashSet[string]([StringComparer]::OrdinalIgnoreCase)
foreach($m in [regex]::Matches($xml,'(?is)<loc>\s*(.*?)\s*</loc>')){
    [void]$urls.Add([System.Net.WebUtility]::HtmlDecode($m.Groups[1].Value.Trim()))
}

# noindex 대상 URL 제거
foreach($rel in $noindexPages){
    if($rel -match '/index\.html$'){
        $p="/" + (($rel -replace '/index\.html$','').Trim('/')) + "/"
        [void]$urls.Remove($baseUrl.TrimEnd('/') + $p)
    } else {
        [void]$urls.Remove($baseUrl.TrimEnd('/') + "/" + $rel)
        [void]$urls.Remove($baseUrl.TrimEnd('/') + "/" + $rel + "/")
    }
}

# 공개 보조페이지 14개 추가
foreach($rel in $indexableFolders){
    $p="/" + (($rel -replace '/index\.html$','').Trim('/')) + "/"
    [void]$urls.Add($baseUrl.TrimEnd('/') + $p)
}

$sorted=@($urls | Sort-Object)

$sb=New-Object System.Text.StringBuilder
[void]$sb.AppendLine('<?xml version="1.0" encoding="UTF-8"?>')
[void]$sb.AppendLine('<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">')
foreach($u in $sorted){
    $esc=[System.Security.SecurityElement]::Escape($u)
    [void]$sb.AppendLine('  <url>')
    [void]$sb.AppendLine('    <loc>' + $esc + '</loc>')
    [void]$sb.AppendLine('  </url>')
}
[void]$sb.AppendLine('</urlset>')

$utf8NoBom=New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllText($sitemapPath,$sb.ToString(),$utf8NoBom)

# 4) 검증
$indexableOk=0
$noindexOk=0
$canonicalOk=0

foreach($rel in $indexableFolders){
    $file=Join-Path $root ($rel.Replace('/','\'))
    if(!(Test-Path -LiteralPath $file)){ continue }
    $html=Get-Content -LiteralPath $file -Raw -Encoding UTF8

    $hasNoindex=$html -match '(?is)<meta\b(?=[^>]*\bname\s*=\s*["'']robots["''])(?=[^>]*\bcontent\s*=\s*["''][^"'']*noindex'
    if(!$hasNoindex){ $indexableOk++ }

    $path="/" + (($rel -replace '/index\.html$','').Trim('/')) + "/"
    $url=$baseUrl.TrimEnd('/') + $path
    if($html -match [regex]::Escape('href="' + $url + '"')){ $canonicalOk++ }
}

foreach($rel in $noindexPages){
    $file=Join-Path $root ($rel.Replace('/','\'))
    if(!(Test-Path -LiteralPath $file)){ continue }
    $html=Get-Content -LiteralPath $file -Raw -Encoding UTF8
    if($html -match '(?is)<meta\b(?=[^>]*\bname\s*=\s*["'']robots["''])(?=[^>]*\bcontent\s*=\s*["''][^"'']*noindex'){
        $noindexOk++
    }
}

$xmlCheck=Get-Content -LiteralPath $sitemapPath -Raw -Encoding UTF8
$locs=@([regex]::Matches($xmlCheck,'(?is)<loc>\s*(.*?)\s*</loc>'))
$locSet=New-Object System.Collections.Generic.HashSet[string]([StringComparer]::OrdinalIgnoreCase)
foreach($m in $locs){ [void]$locSet.Add([System.Net.WebUtility]::HtmlDecode($m.Groups[1].Value.Trim())) }

$publicInSitemap=0
foreach($rel in $indexableFolders){
    $p="/" + (($rel -replace '/index\.html$','').Trim('/')) + "/"
    if($locSet.Contains($baseUrl.TrimEnd('/') + $p)){ $publicInSitemap++ }
}

$noindexInSitemap=0
foreach($rel in $noindexPages){
    $u1=$baseUrl.TrimEnd('/') + "/" + $rel
    $u2=$u1 + "/"
    if($locSet.Contains($u1) -or $locSet.Contains($u2)){ $noindexInSitemap++ }
}

Write-Host ""
Write-Host "=== 보조페이지 색인 정리 완료 ==="
Write-Host "실제 수정 HTML:" $modified
Write-Host "누락 파일:" $missing.Count
Write-Host "공개 보조페이지 index 상태 정상:" $indexableOk "/ 14"
Write-Host "공개 보조페이지 self canonical 정상:" $canonicalOk "/ 14"
Write-Host "noindex 대상 정상:" $noindexOk "/ 22"
Write-Host "공개 보조페이지 sitemap 포함:" $publicInSitemap "/ 14"
Write-Host "noindex 페이지 sitemap 포함:" $noindexInSitemap
Write-Host "sitemap 전체 URL:" $locs.Count
Write-Host "sitemap 고유 URL:" $locSet.Count
Write-Host "sitemap 중복 URL:" ($locs.Count - $locSet.Count)
Write-Host "백업 폴더:" $backup

if($missing.Count -gt 0){
    Write-Host ""
    Write-Host "=== 누락 파일 ==="
    $missing | ForEach-Object { Write-Host $_ }
}
