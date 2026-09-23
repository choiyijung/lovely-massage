$ErrorActionPreference = "Stop"

$root = "E:\mong-lounge-site"
$baseUrl = "https://1004테라피.shop"
$stamp = Get-Date -Format "yyyyMMdd-HHmmss"
$backup = "E:\mong-lounge-sitemap-backup-$stamp"

$robotsPath = Join-Path $root "robots.txt"
$sitemapPath = Join-Path $root "sitemap.xml"

New-Item -ItemType Directory -Force -Path $backup | Out-Null

if(Test-Path -LiteralPath $robotsPath){
    Copy-Item -LiteralPath $robotsPath -Destination (Join-Path $backup "robots.txt") -Force
}
if(Test-Path -LiteralPath $sitemapPath){
    Copy-Item -LiteralPath $sitemapPath -Destination (Join-Path $backup "sitemap.xml") -Force
}

function Has-Noindex([string]$html){
    foreach($m in [regex]::Matches($html,'(?is)<meta\b[^>]*>')){
        $tag = $m.Value
        if(
            ($tag -match '(?is)\bname\s*=\s*["''](?:robots|googlebot|naverbot|yeti)["'']') -and
            ($tag -match '(?is)\bcontent\s*=\s*["''][^"'']*noindex[^"'']*["'']')
        ){
            return $true
        }
    }
    return $false
}

function Get-Canonical([string]$html){
    $m = [regex]::Match(
        $html,
        '(?is)<link\b(?=[^>]*\brel\s*=\s*["'']canonical["''])(?=[^>]*\bhref\s*=\s*["'']([^"'']+)["''])[^>]*>'
    )
    if($m.Success){ return [System.Net.WebUtility]::HtmlDecode($m.Groups[1].Value.Trim()) }
    return ""
}

function Xml-Escape([string]$s){
    return [System.Security.SecurityElement]::Escape($s)
}

# 핵심 색인 대상만 sitemap에 포함:
# 메인 1 + 지역SEO + 업체상세
$targetFiles = New-Object System.Collections.Generic.List[System.IO.FileInfo]

$main = Get-Item -LiteralPath (Join-Path $root "index.html")
$targetFiles.Add($main)

Get-ChildItem -LiteralPath (Join-Path $root "massage") -Recurse -File -Filter "index.html" |
    ForEach-Object { $targetFiles.Add($_) }

Get-ChildItem -LiteralPath (Join-Path $root "shops") -Recurse -File -Filter "index.html" |
    ForEach-Object { $targetFiles.Add($_) }

$urls = New-Object System.Collections.Generic.List[string]
$seen = New-Object System.Collections.Generic.HashSet[string]([StringComparer]::OrdinalIgnoreCase)

$noindexSkipped = 0
$canonicalMissing = 0

foreach($f in $targetFiles){
    $html = Get-Content -LiteralPath $f.FullName -Raw -Encoding UTF8

    if(Has-Noindex $html){
        $noindexSkipped++
        continue
    }

    $url = Get-Canonical $html

    if([string]::IsNullOrWhiteSpace($url)){
        $canonicalMissing++

        # 메인 또는 예외 페이지용 fallback
        $rel = $f.FullName.Substring($root.Length).TrimStart("\").Replace("\","/")
        if($rel -eq "index.html"){
            $url = $baseUrl.TrimEnd("/") + "/"
        } else {
            $dir = $rel -replace '/index\.html$',''
            $url = $baseUrl.TrimEnd("/") + "/" + $dir.Trim("/") + "/"
        }
    }

    if($seen.Add($url)){
        $urls.Add($url)
    }
}

# sitemap.xml 생성
$sb = New-Object System.Text.StringBuilder
[void]$sb.AppendLine('<?xml version="1.0" encoding="UTF-8"?>')
[void]$sb.AppendLine('<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">')

foreach($u in $urls){
    [void]$sb.AppendLine('  <url>')
    [void]$sb.AppendLine('    <loc>' + (Xml-Escape $u) + '</loc>')
    [void]$sb.AppendLine('  </url>')
}

[void]$sb.AppendLine('</urlset>')

$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllText($sitemapPath, $sb.ToString(), $utf8NoBom)

# robots.txt: 기존 규칙 유지 + Sitemap 선언 1개로 통일
$robots = ""
if(Test-Path -LiteralPath $robotsPath){
    $robots = Get-Content -LiteralPath $robotsPath -Raw -Encoding UTF8
}

if([string]::IsNullOrWhiteSpace($robots)){
    $robots = "User-agent: *`r`nAllow: /`r`n"
}

$robots = [regex]::Replace(
    $robots,
    '(?im)^\s*Sitemap\s*:\s*\S+\s*$',
    ''
)

$robots = $robots.TrimEnd() + "`r`n`r`nSitemap: $baseUrl/sitemap.xml`r`n"
[System.IO.File]::WriteAllText($robotsPath, $robots, $utf8NoBom)

# 검증
$robotsCheck = Get-Content -LiteralPath $robotsPath -Raw -Encoding UTF8
$sitemapCheck = Get-Content -LiteralPath $sitemapPath -Raw -Encoding UTF8

$sitemapLines = @(
    [regex]::Matches($robotsCheck,'(?im)^\s*Sitemap\s*:\s*(\S+)\s*$')
)

$locMatches = @(
    [regex]::Matches($sitemapCheck,'(?is)<loc>\s*(.*?)\s*</loc>')
)

$locSet = New-Object System.Collections.Generic.HashSet[string]([StringComparer]::OrdinalIgnoreCase)
foreach($m in $locMatches){
    [void]$locSet.Add([System.Net.WebUtility]::HtmlDecode($m.Groups[1].Value.Trim()))
}

$duplicateCount = $locMatches.Count - $locSet.Count

Write-Host ""
Write-Host "=== ROBOTS + SITEMAP 생성 완료 ==="
Write-Host "핵심 대상 파일:" $targetFiles.Count
Write-Host "noindex 제외:" $noindexSkipped
Write-Host "canonical fallback 사용:" $canonicalMissing
Write-Host "sitemap URL 수:" $locMatches.Count
Write-Host "sitemap 고유 URL 수:" $locSet.Count
Write-Host "sitemap 중복 URL:" $duplicateCount
Write-Host "robots Sitemap 선언 수:" $sitemapLines.Count
if($sitemapLines.Count -gt 0){
    Write-Host "robots Sitemap:" $sitemapLines[0].Groups[1].Value
}
Write-Host "sitemap 파일:" $sitemapPath
Write-Host "백업 폴더:" $backup
