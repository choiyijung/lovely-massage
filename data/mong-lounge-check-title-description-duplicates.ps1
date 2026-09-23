$ErrorActionPreference = "Stop"

$root = "E:\mong-lounge-site"
$dataDir = Join-Path $root "data"
$stamp = Get-Date -Format "yyyyMMdd-HHmmss"

$detailReport = Join-Path $dataDir "title-description-duplicate-detail-$stamp.csv"
$summaryReport = Join-Path $dataDir "title-description-duplicate-summary-$stamp.csv"

function Clean-Text([string]$text) {
    if ($null -eq $text) { return "" }
    $x = [System.Net.WebUtility]::HtmlDecode($text)
    $x = [regex]::Replace($x, '\s+', ' ')
    return $x.Trim()
}

function Get-Title([string]$html) {
    $m = [regex]::Match($html, '(?is)<title\b[^>]*>(.*?)</title>')
    if ($m.Success) { return Clean-Text $m.Groups[1].Value }
    return ""
}

function Get-MetaDescription([string]$html) {
    $patterns = @(
        '(?is)<meta\b(?=[^>]*\bname\s*=\s*["'']description["''])(?=[^>]*\bcontent\s*=\s*["'']([^"'']*)["''])[^>]*>',
        '(?is)<meta\b(?=[^>]*\bcontent\s*=\s*["'']([^"'']*)["''])(?=[^>]*\bname\s*=\s*["'']description["''])[^>]*>'
    )
    foreach ($p in $patterns) {
        $m = [regex]::Match($html, $p)
        if ($m.Success) { return Clean-Text $m.Groups[1].Value }
    }
    return ""
}

$files = @(
    Get-ChildItem -LiteralPath $root -Recurse -File -Filter "*.html" |
    Where-Object {
        $_.FullName -notmatch '\\node_modules\\|\\\.git\\|\\backup\\|\\_backup\\'
    }
)

$rows = New-Object System.Collections.Generic.List[object]

foreach ($f in $files) {
    try {
        $html = Get-Content -LiteralPath $f.FullName -Raw -Encoding UTF8
        $title = Get-Title $html
        $desc = Get-MetaDescription $html
        $rel = $f.FullName.Substring($root.Length).TrimStart("\").Replace("\","/")

        $rows.Add([pscustomobject]@{
            File = $rel
            FullPath = $f.FullName
            Title = $title
            Description = $desc
            TitleLength = $title.Length
            DescriptionLength = $desc.Length
            TitleMissing = [string]::IsNullOrWhiteSpace($title)
            DescriptionMissing = [string]::IsNullOrWhiteSpace($desc)
        })
    }
    catch {
        $rows.Add([pscustomobject]@{
            File = $f.FullName
            FullPath = $f.FullName
            Title = ""
            Description = ""
            TitleLength = 0
            DescriptionLength = 0
            TitleMissing = $true
            DescriptionMissing = $true
        })
    }
}

$titleDupGroups = @(
    $rows |
    Where-Object { -not $_.TitleMissing } |
    Group-Object Title |
    Where-Object { $_.Count -gt 1 } |
    Sort-Object Count -Descending
)

$descDupGroups = @(
    $rows |
    Where-Object { -not $_.DescriptionMissing } |
    Group-Object Description |
    Where-Object { $_.Count -gt 1 } |
    Sort-Object Count -Descending
)

$detail = New-Object System.Collections.Generic.List[object]

$groupNo = 0
foreach ($g in $titleDupGroups) {
    $groupNo++
    foreach ($r in $g.Group) {
        $detail.Add([pscustomobject]@{
            Type = "TITLE"
            Group = $groupNo
            DuplicateCount = $g.Count
            Value = $g.Name
            File = $r.File
        })
    }
}

$groupNo = 0
foreach ($g in $descDupGroups) {
    $groupNo++
    foreach ($r in $g.Group) {
        $detail.Add([pscustomobject]@{
            Type = "DESCRIPTION"
            Group = $groupNo
            DuplicateCount = $g.Count
            Value = $g.Name
            File = $r.File
        })
    }
}

$detail | Export-Csv -LiteralPath $detailReport -NoTypeInformation -Encoding UTF8

$summary = @(
    [pscustomobject]@{Item="전체 HTML"; Count=$rows.Count},
    [pscustomobject]@{Item="TITLE 누락"; Count=@($rows | Where-Object {$_.TitleMissing}).Count},
    [pscustomobject]@{Item="DESCRIPTION 누락"; Count=@($rows | Where-Object {$_.DescriptionMissing}).Count},
    [pscustomobject]@{Item="중복 TITLE 그룹"; Count=$titleDupGroups.Count},
    [pscustomobject]@{Item="중복 TITLE 페이지"; Count=@($detail | Where-Object {$_.Type -eq "TITLE"}).Count},
    [pscustomobject]@{Item="중복 DESCRIPTION 그룹"; Count=$descDupGroups.Count},
    [pscustomobject]@{Item="중복 DESCRIPTION 페이지"; Count=@($detail | Where-Object {$_.Type -eq "DESCRIPTION"}).Count}
)
$summary | Export-Csv -LiteralPath $summaryReport -NoTypeInformation -Encoding UTF8

Write-Host ""
Write-Host "=== TITLE / META DESCRIPTION 중복 점검 완료 ==="
Write-Host "전체 HTML:" $rows.Count
Write-Host "TITLE 누락:" @($rows | Where-Object {$_.TitleMissing}).Count
Write-Host "DESCRIPTION 누락:" @($rows | Where-Object {$_.DescriptionMissing}).Count
Write-Host ""
Write-Host "중복 TITLE 그룹:" $titleDupGroups.Count
Write-Host "중복 TITLE 페이지:" @($detail | Where-Object {$_.Type -eq "TITLE"}).Count
Write-Host "중복 DESCRIPTION 그룹:" $descDupGroups.Count
Write-Host "중복 DESCRIPTION 페이지:" @($detail | Where-Object {$_.Type -eq "DESCRIPTION"}).Count

if ($titleDupGroups.Count -gt 0) {
    Write-Host ""
    Write-Host "=== 중복 TITLE 상위 10개 ==="
    $titleDupGroups | Select-Object -First 10 | ForEach-Object {
        Write-Host ("{0}개 | {1}" -f $_.Count, $_.Name)
    }
}

if ($descDupGroups.Count -gt 0) {
    Write-Host ""
    Write-Host "=== 중복 DESCRIPTION 상위 10개 ==="
    $descDupGroups | Select-Object -First 10 | ForEach-Object {
        Write-Host ("{0}개 | {1}" -f $_.Count, $_.Name)
    }
}

Write-Host ""
Write-Host "사이트 파일 수정: 0"
Write-Host "상세 보고서:" $detailReport
Write-Host "요약 보고서:" $summaryReport
