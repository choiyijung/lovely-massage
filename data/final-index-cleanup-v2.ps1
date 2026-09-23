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

function Has-Noindex([string]$html){
    foreach($m in [regex]::Matches($html, '(?is)<meta\b[^>]*>')){
        $tag = $m.Value
        $isRobots = (
            $tag -match '(?is)\bname\s*=\s*"robots"' -or
            $tag -match "(?is)\bname\s*=\s*'robots'"
        )
        $hasNoindex = (
            $tag -match '(?is)\bcontent\s*=\s*"[^"]*noindex[^"]*"' -or
            $tag -match "(?is)\bcontent\s*=\s*'[^']*noindex[^']*'"
        )
        if($isRobots -and $hasNoindex){ return $true }
    }
    return $false
}

function Backup-File([string]$rel){
    $src = Join-Path $root ($rel.Replace('/','\'))
    if(!(Test-Path -LiteralPath $src)){ return }
    $dest = Join-Path $backup ($rel.Replace('/','\'))
    New-Item -ItemType Directory -Force -Path (Split-Path $dest -Parent) | Out-Null
    Copy-Item -LiteralPath $src -Destination $dest -Force
}

function Set-Noindex([string]$html){
    $tags = [regex]::Matches($html, '(?is)<meta\b[^>]*>')
    foreach($m in $tags){
        $tag = $m.Value
        $isRobots = (
            $tag -match '(?is)\bname\s*=\s*"robots"' -or
            $tag -match "(?is)\bname\s*=\s*'robots'"
        )
        if($isRobots){
            return $html.Replace($tag, '<meta name="robots" content="noindex,follow">')
        }
    }

    return [regex]::Replace(
        $html,
        '(?is)</head>',
        '<meta name="robots" content="noindex,follow">' + "`r`n</head>",
        1
    )
}

function Remove-Noindex-Robots([string]$html){
    $result = $html
    $tags = @([regex]::Matches($html, '(?is)<meta\b[^>]*>'))

    foreach($m in $tags){
        $tag = $m.Value
        $isRobots = (
            $tag -match '(?is)\bname\s*=\s*"robots"' -or
            $tag -match "(?is)\bname\s*=\s*'robots'"
        )
        $hasNoindex = (
            $tag -match '(?is)\bcontent\s*=\s*"[^"]*noindex[^"]*"' -or
            $tag -match "(?is)\bcontent\s*=\s*'[^']*noindex[^']*'"
        )

        if($isRobots -and $hasNoindex){
            $result = $result.Replace($tag, '')
        }
    }

    return $result
}

function Set-Canonical([string]$html,[string]$url){
    $result = $html

    $links = @([regex]::Matches($result, '(?is)<link\b[^>]*>'))
    foreach($m in $links){
        $tag = $m.Value
        $isCanonical = (
            $tag -match '(?is)\brel\s*=\s*"canonical"' -or
            $tag -match "(?is)\brel\s*=\s*'canonical'"
        )
        if($isCanonical){
            $result = $result.Replace($tag, '')
        }
    }

    $tag = '<link rel="canonical" href="' + $url + '">'

    if([regex]::IsMatch($result, '(?is)</title>')){
        return [regex]::Replace(
            $result,
            '(?is)</title>',
            '</title>' + "`r`n" + $tag,
            1
        )
    }

    return [regex]::Replace(
        $result,
        '(?is)</head>',
        $tag + "`r`n</head>",
        1
    )
}

New-Item -ItemType Directory -Force -Path $backup | Out-Null

$modified = 0
$missing = New-Object System.Collections.Generic.List[string]

# 1. 공개 보조페이지 14개: index 유지 + self canonical
foreach($rel in $indexableFolders){
    $file = Join-Path $root ($rel.Replace('/','\'))

    if(!(Test-Path -LiteralPath $file)){
        $missing.Add($rel)
        continue
    }

    $html = Get-Content -LiteralPath $file -Raw -Encoding UTF8
    $before = $html

    $path = "/" + (($rel -replace '/index\.html$','').Trim('/')) + "/"
    $url = $baseUrl.TrimEnd('/') + $path

    $html = Remove-Noindex-Robots $html
    $html = Set-Canonical $html $url

    if($html -ne $before){
        Backup-File $rel
        Set-Content -LiteralPath $file -Value $html -Encoding UTF8
        $modified++
    }
}

# 2. 중복/기능/템플릿 22개: noindex,follow
foreach($rel in $noindexPages){
    $file = Join-Path $root ($rel.Replace('/','\'))

    if(!(Test-Path -LiteralPath $file)){
        $missing.Add($rel)
        continue
    }

    $html = Get-Content -LiteralPath $file -Raw -Encoding UTF8
    $before = $html

    $html = Set-Noindex $html

    if($flatToClean.ContainsKey($rel)){
        $html = Set-Canonical $html ($baseUrl.TrimEnd('/') + $flatToClean[$rel])
    }

    if($html -ne $before){
        Backup-File $rel
        Set-Content -LiteralPath $file -Value $html -Encoding UTF8
        $modified++
    }
}

# 3. sitemap 갱신
$sitemapPath = Join-Path $root "sitemap.xml"
if(!(Test-Path -LiteralPath $sitemapPath)){
    throw "sitemap.xml이 없습니다."
}

Backup-File "sitemap.xml"

$xml = Get-Content -LiteralPath $sitemapPath -Raw -Encoding UTF8
$urls = New-Object System.Collections.Generic.HashSet[string]([StringComparer]::OrdinalIgnoreCase)

foreach($m in [regex]::Matches($xml, '(?is)<loc>\s*(.*?)\s*</loc>')){
    [void]$urls.Add([System.Net.WebUtility]::HtmlDecode($m.Groups[1].Value.Trim()))
}

# noindex URL 제거
foreach($rel in $noindexPages){
    if($rel -match '/index\.html$'){
        $p = "/" + (($rel -replace '/index\.html$','').Trim('/')) + "/"
        [void]$urls.Remove($baseUrl.TrimEnd('/') + $p)
    }
    else {
        $u = $baseUrl.TrimEnd('/') + "/" + $rel
        [void]$urls.Remove($u)
        [void]$urls.Remove($u + "/")
    }
}

# 공개 보조페이지 14개 sitemap 추가
foreach($rel in $indexableFolders){
    $p = "/" + (($rel -replace '/index\.html$','').Trim('/')) + "/"
    [void]$urls.Add($baseUrl.TrimEnd('/') + $p)
}

$sorted = @($urls | Sort-Object)

$sb = New-Object System.Text.StringBuilder
[void]$sb.AppendLine('<?xml version="1.0" encoding="UTF-8"?>')
[void]$sb.AppendLine('<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">')

foreach($u in $sorted){
    $esc = [System.Security.SecurityElement]::Escape($u)
    [void]$sb.AppendLine('  <url>')
    [void]$sb.AppendLine('    <loc>' + $esc + '</loc>')
    [void]$sb.AppendLine('  </url>')
}

[void]$sb.AppendLine('</urlset>')

$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllText($sitemapPath, $sb.ToString(), $utf8NoBom)

# 4. 검증
$indexableOk = 0
$canonicalOk = 0
$noindexOk = 0

foreach($rel in $indexableFolders){
    $file = Join-Path $root ($rel.Replace('/','\'))
    if(!(Test-Path -LiteralPath $file)){ continue }

    $html = Get-Content -LiteralPath $file -Raw -Encoding UTF8

    if(!(Has-Noindex $html)){ $indexableOk++ }

    $path = "/" + (($rel -replace '/index\.html$','').Trim('/')) + "/"
    $url = $baseUrl.TrimEnd('/') + $path

    if($html -match [regex]::Escape('href="' + $url + '"')){
        $canonicalOk++
    }
}

foreach($rel in $noindexPages){
    $file = Join-Path $root ($rel.Replace('/','\'))
    if(!(Test-Path -LiteralPath $file)){ continue }

    $html = Get-Content -LiteralPath $file -Raw -Encoding UTF8
    if(Has-Noindex $html){ $noindexOk++ }
}

$xmlCheck = Get-Content -LiteralPath $sitemapPath -Raw -Encoding UTF8
$locs = @([regex]::Matches($xmlCheck, '(?is)<loc>\s*(.*?)\s*</loc>'))

$locSet = New-Object System.Collections.Generic.HashSet[string]([StringComparer]::OrdinalIgnoreCase)
foreach($m in $locs){
    [void]$locSet.Add([System.Net.WebUtility]::HtmlDecode($m.Groups[1].Value.Trim()))
}

$publicInSitemap = 0
foreach($rel in $indexableFolders){
    $p = "/" + (($rel -replace '/index\.html$','').Trim('/')) + "/"
    $u = $baseUrl.TrimEnd('/') + $p
    if($locSet.Contains($u)){ $publicInSitemap++ }
}

$noindexInSitemap = 0
foreach($rel in $noindexPages){
    if($rel -match '/index\.html$'){
        $p = "/" + (($rel -replace '/index\.html$','').Trim('/')) + "/"
        $u = $baseUrl.TrimEnd('/') + $p
        if($locSet.Contains($u)){ $noindexInSitemap++ }
    }
    else {
        $u = $baseUrl.TrimEnd('/') + "/" + $rel
        if($locSet.Contains($u) -or $locSet.Contains($u + "/")){
            $noindexInSitemap++
        }
    }
}

Write-Host ""
Write-Host "=== 보조페이지 색인 정리 v2 완료 ==="
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
