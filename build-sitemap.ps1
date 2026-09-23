param(
  [Parameter(Mandatory=$true)]
  [string]$Domain
)

$ErrorActionPreference = "Stop"
$Domain = $Domain.TrimEnd("/")
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$Paths = Get-Content (Join-Path $Root "sitemap-paths.txt") -Encoding UTF8 | Where-Object { $_.Trim() -ne "" }

$urls = foreach ($p in $Paths) {
  "  <url><loc>$Domain$p</loc></url>"
}

$xml = @(
  '<?xml version="1.0" encoding="UTF-8"?>'
  '<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">'
  $urls
  '</urlset>'
) -join "`r`n"

Set-Content -Path (Join-Path $Root "sitemap.xml") -Value $xml -Encoding UTF8

$robotsPath = Join-Path $Root "robots.txt"
$robots = Get-Content $robotsPath -Raw -Encoding UTF8
$robots = $robots -replace '(?m)^Sitemap:.*\r?\n?', ''
$robots = $robots.TrimEnd() + "`r`n`r`nSitemap: $Domain/sitemap.xml`r`n"
Set-Content -Path $robotsPath -Value $robots -Encoding UTF8

Write-Host "sitemap.xml 생성 완료"
Write-Host "도메인: $Domain"
Write-Host "URL 개수: $($Paths.Count)"
Write-Host "robots.txt Sitemap 주소 반영 완료"
