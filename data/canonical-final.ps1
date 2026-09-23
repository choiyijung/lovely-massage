$ErrorActionPreference = "Stop"

$root = "E:\mong-lounge-site"
$baseUrl = "https://1004테라피.shop"
$stamp = Get-Date -Format "yyyyMMdd-HHmmss"
$backup = "E:\mong-lounge-canonical-final-backup-$stamp"

New-Item -ItemType Directory -Force -Path $backup | Out-Null

function Get-CanonicalUrl([string]$filePath){
    $rel = $filePath.Substring($root.Length).TrimStart("\").Replace("\","/")
    $dir = $rel -replace '/index\.html$',''
    if($dir -eq $rel){ return $null }
    return $baseUrl.TrimEnd('/') + '/' + $dir.Trim('/') + '/'
}

$targets = @(
    Get-ChildItem -LiteralPath (Join-Path $root "massage") -Recurse -File -Filter "index.html"
    Get-ChildItem -LiteralPath (Join-Path $root "shops")   -Recurse -File -Filter "index.html"
)

$modified = 0
$failed = New-Object System.Collections.Generic.List[string]

foreach($f in $targets){
    try{
        $html = Get-Content -LiteralPath $f.FullName -Raw -Encoding UTF8
        $before = $html
        $url = Get-CanonicalUrl $f.FullName
        if(!$url){ continue }

        # 기존 canonical 모두 제거
        $html = [regex]::Replace(
            $html,
            '(?is)\s*<link\b(?=[^>]*\brel\s*=\s*["'']canonical["''])[^>]*>\s*',
            "`r`n"
        )

        $tag = '<link rel="canonical" href="' + $url + '">'

        # </title> 바로 뒤에 삽입, 없으면 </head> 직전
        if([regex]::IsMatch($html,'(?is)</title>')){
            $html = [regex]::Replace(
                $html,
                '(?is)</title>',
                '</title>' + "`r`n" + $tag,
                1
            )
        }
        elseif([regex]::IsMatch($html,'(?is)</head>')){
            $html = [regex]::Replace(
                $html,
                '(?is)</head>',
                $tag + "`r`n</head>",
                1
            )
        }
        else{
            throw "HEAD 닫기 태그를 찾지 못했습니다."
        }

        if($html -ne $before){
            $rel = $f.FullName.Substring($root.Length).TrimStart("\")
            $dest = Join-Path $backup $rel
            New-Item -ItemType Directory -Force -Path (Split-Path $dest -Parent) | Out-Null
            Copy-Item -LiteralPath $f.FullName -Destination $dest -Force
            Set-Content -LiteralPath $f.FullName -Value $html -Encoding UTF8
            $modified++
        }
    }
    catch{
        $failed.Add("$($f.FullName) | $($_.Exception.Message)")
    }
}

# 검증
$missing = 0
$multi = 0
$relative = 0
$wrongHost = 0
$selfMismatch = 0
$map = @{}

foreach($f in $targets){
    $html = Get-Content -LiteralPath $f.FullName -Raw -Encoding UTF8
    $matches = [regex]::Matches(
        $html,
        '(?is)<link\b(?=[^>]*\brel\s*=\s*["'']canonical["''])(?=[^>]*\bhref\s*=\s*["'']([^"'']+)["''])[^>]*>'
    )

    if($matches.Count -eq 0){ $missing++; continue }
    if($matches.Count -gt 1){ $multi++ }

    $href = $matches[0].Groups[1].Value
    if($href -notmatch '^https://'){ $relative++ }

    if($href -notmatch '^https://1004테라피\.shop/'){ $wrongHost++ }

    $expected = Get-CanonicalUrl $f.FullName
    if($href -ne $expected){ $selfMismatch++ }

    if(!$map.ContainsKey($href)){ $map[$href] = 0 }
    $map[$href]++
}

$dupGroups = @($map.GetEnumerator() | Where-Object { $_.Value -gt 1 })
$dupPages = ($dupGroups | Measure-Object -Property Value -Sum).Sum
if($null -eq $dupPages){ $dupPages = 0 }

Write-Host ""
Write-Host "=== CANONICAL 최종 적용 완료 ==="
Write-Host "운영 도메인:" $baseUrl
Write-Host "적용 대상:" $targets.Count
Write-Host "실제 수정:" $modified
Write-Host "실패:" $failed.Count
Write-Host ""
Write-Host "canonical 누락:" $missing
Write-Host "canonical 태그 2개 이상:" $multi
Write-Host "상대경로 canonical:" $relative
Write-Host "다른 호스트 canonical:" $wrongHost
Write-Host "자기 경로 불일치:" $selfMismatch
Write-Host "canonical 중복 그룹:" $dupGroups.Count
Write-Host "canonical 중복 페이지:" $dupPages
Write-Host "백업 폴더:" $backup

if($failed.Count -gt 0){
    $report = Join-Path $root "data\canonical-final-fail-$stamp.txt"
    $failed | Set-Content -LiteralPath $report -Encoding UTF8
    Write-Host "실패 보고서:" $report
}
