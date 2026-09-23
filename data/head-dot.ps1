$ErrorActionPreference = "Stop"

$root = "E:\mong-lounge-site"
$stamp = Get-Date -Format "yyyyMMdd-HHmmss"
$backup = "E:\mong-lounge-head-dot-backup-$stamp"

New-Item -ItemType Directory -Force -Path $backup | Out-Null

$files = @(
    Get-ChildItem -LiteralPath $root -Recurse -File -Filter "*.html" |
    Where-Object { $_.FullName -notmatch '\\\.git\\|\\node_modules\\|\\backup\\|\\_backup\\' }
)

$modified = 0
$entityCount = 0
$remaining = 0
$failed = New-Object System.Collections.Generic.List[string]

function Replace-DotEntity([string]$value){
    if($null -eq $value){ return $value }
    return $value.
        Replace('&#183;','·').
        Replace('&#xB7;','·').
        Replace('&#xb7;','·').
        Replace('&middot;','·')
}

foreach($f in $files){
    try{
        $html = Get-Content -LiteralPath $f.FullName -Raw -Encoding UTF8
        $before = $html
        $changedThis = $false

        # TITLE
        $html = [regex]::Replace(
            $html,
            '(?is)<title\b([^>]*)>(.*?)</title>',
            [System.Text.RegularExpressions.MatchEvaluator]{
                param($m)
                $v = $m.Groups[2].Value
                $new = Replace-DotEntity $v
                if($new -ne $v){ $script:entityCount += 1; $script:changedThis = $true }
                return "<title$($m.Groups[1].Value)>$new</title>"
            }
        )

        # META description / og:title / og:description
        $html = [regex]::Replace(
            $html,
            '(?is)<meta\b([^>]*?)content\s*=\s*(["''])(.*?)\2([^>]*)>',
            [System.Text.RegularExpressions.MatchEvaluator]{
                param($m)

                $full = $m.Value
                $attrs = $m.Groups[1].Value + " " + $m.Groups[4].Value

                $isTarget =
                    ($attrs -match '(?is)\bname\s*=\s*["'']description["'']') -or
                    ($attrs -match '(?is)\bproperty\s*=\s*["'']og:title["'']') -or
                    ($attrs -match '(?is)\bproperty\s*=\s*["'']og:description["'']')

                if(!$isTarget){ return $full }

                $v = $m.Groups[3].Value
                $new = Replace-DotEntity $v
                if($new -ne $v){
                    $script:entityCount += 1
                    $script:changedThis = $true
                }

                return '<meta' + $m.Groups[1].Value + 'content=' + $m.Groups[2].Value + $new + $m.Groups[2].Value + $m.Groups[4].Value + '>'
            }
        )

        if($changedThis -and $html -ne $before){
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
foreach($f in $files){
    $html = Get-Content -LiteralPath $f.FullName -Raw -Encoding UTF8

    $headMatch = [regex]::Match($html,'(?is)<head\b[^>]*>(.*?)</head>')
    if(!$headMatch.Success){ continue }

    $head = $headMatch.Groups[1].Value

    $targets = @()

    $t = [regex]::Match($head,'(?is)<title\b[^>]*>(.*?)</title>')
    if($t.Success){ $targets += $t.Groups[1].Value }

    foreach($m in [regex]::Matches($head,'(?is)<meta\b[^>]*>')){
        $tag = $m.Value
        if(
            $tag -match '(?is)\bname\s*=\s*["'']description["'']' -or
            $tag -match '(?is)\bproperty\s*=\s*["'']og:title["'']' -or
            $tag -match '(?is)\bproperty\s*=\s*["'']og:description["'']'
        ){
            $targets += $tag
        }
    }

    foreach($x in $targets){
        if($x -match '&#183;|&#xB7;|&#xb7;|&middot;'){
            $remaining++
        }
    }
}

Write-Host ""
Write-Host "=== TITLE / META 가운데점 entity 정리 완료 ==="
Write-Host "전체 HTML:" $files.Count
Write-Host "수정 페이지:" $modified
Write-Host "변경 감지 항목:" $entityCount
Write-Host "남은 entity 항목:" $remaining
Write-Host "실패:" $failed.Count
Write-Host "백업 폴더:" $backup

if($failed.Count -gt 0){
    $report = Join-Path $root "data\head-dot-fail-$stamp.txt"
    $failed | Set-Content -LiteralPath $report -Encoding UTF8
    Write-Host "실패 보고서:" $report
}
