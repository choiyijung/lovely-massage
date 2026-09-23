$ErrorActionPreference = "Stop"

$root = "E:\mong-lounge-site"
$baseUrl = "https://1004테라피.shop"
$dataDir = Join-Path $root "data"
$stamp = Get-Date -Format "yyyyMMdd-HHmmss"
$detailReport = Join-Path $dataDir "robots-sitemap-index-audit-$stamp.csv"

function Normalize-Url([string]$u){
    if([string]::IsNullOrWhiteSpace($u)){ return "" }
    $x = [System.Net.WebUtility]::HtmlDecode($u.Trim())

    try {
        $uri = [uri]$x
        $path = [uri]::UnescapeDataString($uri.AbsolutePath)
        $path = $path.Replace('\','/')
        if($path -eq "" -or $path -eq "/"){
            return $baseUrl.TrimEnd('/') + "/"
        }
        $path = "/" + $path.Trim('/')
        return $baseUrl.TrimEnd('/') + $path + "/"
    } catch {
        $x = [uri]::UnescapeDataString($x)
        $x = $x.Replace('\','/')
        $x = $x -replace '^https?://[^/]+',''
        if($x -eq "" -or $x -eq "/"){
            return $baseUrl.TrimEnd('/') + "/"
        }
        return $baseUrl.TrimEnd('/') + "/" + $x.Trim('/') + "/"
    }
}

function File-To-Url([string]$file){
    $rel = $file.Substring($root.Length).TrimStart('\').Replace('\','/')
    if($rel -eq "index.html"){ return $baseUrl.TrimEnd('/') + "/" }
    if($rel -match '/index\.html$'){
        $dir = $rel -replace '/index\.html$',''
        return $baseUrl.TrimEnd('/') + "/" + $dir.Trim('/') + "/"
    }
    return $baseUrl.TrimEnd('/') + "/" + $rel
}

function Page-Class([string]$rel){
    $p = $rel.Replace('\','/')
    if($p -like "massage/*"){ return "지역SEO" }
    if($p -like "shops/*"){ return "업체상세" }
    if($p -eq "index.html"){ return "메인" }
    return "기타"
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

$robotsPath = Join-Path $root "robots.txt"
$robotsExists = Test-Path -LiteralPath $robotsPath
$robotsText = if($robotsExists){ Get-Content -LiteralPath $robotsPath -Raw -Encoding UTF8 } else { "" }

$robotsSitemapLines = @()
if($robotsExists){
    $robotsSitemapLines = @(
        [regex]::Matches($robotsText,'(?im)^\s*Sitemap\s*:\s*(\S+)\s*$') |
        ForEach-Object { $_.Groups[1].Value.Trim() }
    )
}

$robotsDisallowAll = $false
if($robotsExists){
    # User-agent:* 섹션에서 Disallow:/ 여부를 단순 안전 점검
    $blocks = [regex]::Split($robotsText,'(?im)(?=^\s*User-agent\s*:)')
    foreach($b in $blocks){
        if($b -match '(?im)^\s*User-agent\s*:\s*\*\s*$' -and
           $b -match '(?im)^\s*Disallow\s*:\s*/\s*$'){
            $robotsDisallowAll = $true
        }
    }
}

# sitemap XML 수집
$sitemapFiles = @(
    Get-ChildItem -LiteralPath $root -Recurse -File -Filter "sitemap*.xml" |
    Where-Object {
        $_.FullName -notmatch '\\\.git\\|\\node_modules\\|\\backup\\|\\_backup\\|\\data\\'
    }
)

$sitemapSet = New-Object System.Collections.Generic.HashSet[string]([StringComparer]::OrdinalIgnoreCase)
$sitemapRawCount = 0
$sitemapParseFailures = New-Object System.Collections.Generic.List[string]

foreach($sf in $sitemapFiles){
    try{
        $xmlText = Get-Content -LiteralPath $sf.FullName -Raw -Encoding UTF8

        # namespace 유무와 관계없이 loc 텍스트 추출
        foreach($m in [regex]::Matches($xmlText,'(?is)<loc\b[^>]*>\s*(.*?)\s*</loc>')){
            $loc = [System.Net.WebUtility]::HtmlDecode($m.Groups[1].Value.Trim())
            if(!$loc){ continue }

            # sitemap index의 sitemap 파일 URL은 페이지 URL 집합에서 제외
            if($loc -match '(?i)\.xml(?:$|\?)'){ continue }

            $sitemapRawCount++
            [void]$sitemapSet.Add((Normalize-Url $loc))
        }
    }
    catch{
        $sitemapParseFailures.Add("$($sf.FullName) | $($_.Exception.Message)")
    }
}

$htmlFiles = @(
    Get-ChildItem -LiteralPath $root -Recurse -File -Filter "*.html" |
    Where-Object {
        $_.FullName -notmatch '\\\.git\\|\\node_modules\\|\\backup\\|\\_backup\\'
    }
)

$rows = New-Object System.Collections.Generic.List[object]

foreach($f in $htmlFiles){
    try{
        $rel = $f.FullName.Substring($root.Length).TrimStart('\').Replace('\','/')
        $html = Get-Content -LiteralPath $f.FullName -Raw -Encoding UTF8
        $noindex = Has-Noindex $html
        $url = Normalize-Url (File-To-Url $f.FullName)
        $inSitemap = $sitemapSet.Contains($url)
        $class = Page-Class $rel

        $issue = New-Object System.Collections.Generic.List[string]

        if(($class -eq "지역SEO" -or $class -eq "업체상세") -and $noindex){
            $issue.Add("SEO_NOIDEX")
        }
        if(!$noindex -and !$inSitemap){
            $issue.Add("INDEXABLE_NOT_IN_SITEMAP")
        }
        if($noindex -and $inSitemap){
            $issue.Add("NOINDEX_IN_SITEMAP")
        }

        $rows.Add([pscustomobject]@{
            Class = $class
            File = $rel
            Url = $url
            Noindex = $noindex
            InSitemap = $inSitemap
            Issue = ($issue -join ",")
        })
    }
    catch{}
}

$rows | Export-Csv -LiteralPath $detailReport -NoTypeInformation -Encoding UTF8

$allUrls = New-Object System.Collections.Generic.HashSet[string]([StringComparer]::OrdinalIgnoreCase)
foreach($r in $rows){ [void]$allUrls.Add($r.Url) }

$extraSitemap = @(
    $sitemapSet | Where-Object { -not $allUrls.Contains($_) }
)

$region = @($rows | Where-Object Class -eq "지역SEO")
$shops  = @($rows | Where-Object Class -eq "업체상세")
$main   = @($rows | Where-Object Class -eq "메인")
$other  = @($rows | Where-Object Class -eq "기타")

$seoNoindex = @($rows | Where-Object {
    ($_.Class -eq "지역SEO" -or $_.Class -eq "업체상세") -and $_.Noindex
})

$indexableMissing = @($rows | Where-Object { -not $_.Noindex -and -not $_.InSitemap })
$noindexInSitemap = @($rows | Where-Object { $_.Noindex -and $_.InSitemap })

Write-Host ""
Write-Host "=== ROBOTS / SITEMAP / 색인 대상 읽기 전용 최종 점검 ==="
Write-Host "사이트 파일 수정: 0"
Write-Host "전체 HTML:" $rows.Count
Write-Host "지역SEO:" $region.Count
Write-Host "업체상세:" $shops.Count
Write-Host "메인:" $main.Count
Write-Host "기타:" $other.Count
Write-Host ""

Write-Host "=== ROBOTS.TXT ==="
Write-Host "robots.txt 존재:" $robotsExists
Write-Host "Sitemap 선언 수:" $robotsSitemapLines.Count
if($robotsSitemapLines.Count -gt 0){
    foreach($s in $robotsSitemapLines){ Write-Host "  Sitemap:" $s }
}
Write-Host "User-agent:* 전체 차단(Disallow:/):" $robotsDisallowAll
Write-Host ""

Write-Host "=== SITEMAP ==="
Write-Host "sitemap 파일 수:" $sitemapFiles.Count
Write-Host "sitemap 원시 URL 수:" $sitemapRawCount
Write-Host "sitemap 고유 페이지 URL 수:" $sitemapSet.Count
Write-Host "sitemap 파싱 실패:" $sitemapParseFailures.Count
Write-Host "사이트에 없는 sitemap URL:" $extraSitemap.Count
Write-Host ""

Write-Host "=== NOINDEX / 색인 대상 ==="
Write-Host "전체 noindex:" (@($rows | Where-Object Noindex)).Count
Write-Host "지역SEO noindex:" (@($region | Where-Object Noindex)).Count
Write-Host "업체상세 noindex:" (@($shops | Where-Object Noindex)).Count
Write-Host "기타 noindex:" (@($other | Where-Object Noindex)).Count
Write-Host "SEO 페이지 noindex 문제:" $seoNoindex.Count
Write-Host "색인가능 페이지 sitemap 누락:" $indexableMissing.Count
Write-Host "noindex 페이지 sitemap 포함:" $noindexInSitemap.Count
Write-Host ""

Write-Host "=== 문제 샘플 ==="
if($seoNoindex.Count -gt 0){
    Write-Host "[SEO인데 noindex]"
    $seoNoindex | Select-Object -First 10 | ForEach-Object { Write-Host "  $($_.File)" }
}
if($indexableMissing.Count -gt 0){
    Write-Host "[색인가능인데 sitemap 누락]"
    $indexableMissing | Select-Object -First 10 | ForEach-Object { Write-Host "  $($_.File)" }
}
if($noindexInSitemap.Count -gt 0){
    Write-Host "[noindex인데 sitemap 포함]"
    $noindexInSitemap | Select-Object -First 10 | ForEach-Object { Write-Host "  $($_.File)" }
}
if($extraSitemap.Count -gt 0){
    Write-Host "[사이트에 없는 sitemap URL]"
    $extraSitemap | Select-Object -First 10 | ForEach-Object { Write-Host "  $_" }
}

Write-Host ""
Write-Host "상세 보고서:" $detailReport
