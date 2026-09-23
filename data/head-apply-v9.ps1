$ErrorActionPreference = "Stop"

$root = "E:\mong-lounge-site"
$dataDir = Join-Path $root "data"
$stamp = Get-Date -Format "yyyyMMdd-HHmmss"
$backup = "E:\mong-lounge-head-v9-backup-$stamp"

$csv = Get-ChildItem -LiteralPath $dataDir -File -Filter "head-rewrite-preview-v9-*.csv" |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1

if(!$csv){
    throw "v9 미리보기 CSV를 찾지 못했습니다."
}

$rows = Import-Csv -LiteralPath $csv.FullName
New-Item -ItemType Directory -Force -Path $backup | Out-Null

function HtmlAttrEncode([string]$s){
    if($null -eq $s){ return "" }
    return $s.Replace('&','&amp;').Replace('"','&quot;').Replace('<','&lt;').Replace('>','&gt;')
}

$modified = 0
$failed = New-Object System.Collections.Generic.List[string]
$titleOgMismatch = 0
$metaOgMismatch = 0
$missingTitle = 0
$missingMeta = 0
$missingOgTitle = 0
$missingOgDesc = 0

foreach($r in $rows){
    try{
        $file = Join-Path $root ($r.File.Replace('/','\'))
        if(!(Test-Path -LiteralPath $file)){
            $failed.Add("$($r.File) | 파일 없음")
            continue
        }

        $html = Get-Content -LiteralPath $file -Raw -Encoding UTF8
        $before = $html

        $newTitle = [string]$r.NewTitle
        $newMeta  = [string]$r.NewDescription
        $encTitle = HtmlAttrEncode $newTitle
        $encMeta  = HtmlAttrEncode $newMeta

        # TITLE
        if([regex]::IsMatch($html,'(?is)<title\b[^>]*>.*?</title>')){
            $html = [regex]::Replace(
                $html,
                '(?is)<title\b[^>]*>.*?</title>',
                '<title>' + $newTitle + '</title>',
                1
            )
        } else {
            $html = [regex]::Replace(
                $html,
                '(?is)<head\b[^>]*>',
                '$0' + "`r`n<title>" + $newTitle + "</title>",
                1
            )
        }

        # META DESCRIPTION
        $metaDescPattern = '(?is)<meta\b(?=[^>]*\bname\s*=\s*["'']description["''])[^>]*>'
        if([regex]::IsMatch($html,$metaDescPattern)){
            $html = [regex]::Replace(
                $html,
                $metaDescPattern,
                '<meta name="description" content="' + $encMeta + '">',
                1
            )
        } else {
            $html = [regex]::Replace(
                $html,
                '(?is)</title>',
                '$0' + "`r`n<meta name=`"description`" content=`"" + $encMeta + "`">",
                1
            )
        }

        # OG TITLE
        $ogTitlePattern = '(?is)<meta\b(?=[^>]*\bproperty\s*=\s*["'']og:title["''])[^>]*>'
        if([regex]::IsMatch($html,$ogTitlePattern)){
            $html = [regex]::Replace(
                $html,
                $ogTitlePattern,
                '<meta property="og:title" content="' + $encTitle + '">',
                1
            )
        } else {
            $html = [regex]::Replace(
                $html,
                '(?is)</head>',
                '<meta property="og:title" content="' + $encTitle + '">' + "`r`n</head>",
                1
            )
        }

        # OG DESCRIPTION
        $ogDescPattern = '(?is)<meta\b(?=[^>]*\bproperty\s*=\s*["'']og:description["''])[^>]*>'
        if([regex]::IsMatch($html,$ogDescPattern)){
            $html = [regex]::Replace(
                $html,
                $ogDescPattern,
                '<meta property="og:description" content="' + $encMeta + '">',
                1
            )
        } else {
            $html = [regex]::Replace(
                $html,
                '(?is)</head>',
                '<meta property="og:description" content="' + $encMeta + '">' + "`r`n</head>",
                1
            )
        }

        if($html -ne $before){
            $rel = $file.Substring($root.Length).TrimStart("\")
            $dest = Join-Path $backup $rel
            New-Item -ItemType Directory -Force -Path (Split-Path $dest -Parent) | Out-Null
            Copy-Item -LiteralPath $file -Destination $dest -Force
            Set-Content -LiteralPath $file -Value $html -Encoding UTF8
            $modified++
        }
    }
    catch{
        $failed.Add("$($r.File) | $($_.Exception.Message)")
    }
}

# 검증
foreach($r in $rows){
    $file = Join-Path $root ($r.File.Replace('/','\'))
    if(!(Test-Path -LiteralPath $file)){ continue }

    $html = Get-Content -LiteralPath $file -Raw -Encoding UTF8

    $tm = [regex]::Match($html,'(?is)<title\b[^>]*>(.*?)</title>')
    $mm = [regex]::Match($html,'(?is)<meta\b(?=[^>]*\bname\s*=\s*["'']description["''])(?=[^>]*\bcontent\s*=\s*["'']([^"'']*)["''])[^>]*>')
    $ot = [regex]::Match($html,'(?is)<meta\b(?=[^>]*\bproperty\s*=\s*["'']og:title["''])(?=[^>]*\bcontent\s*=\s*["'']([^"'']*)["''])[^>]*>')
    $od = [regex]::Match($html,'(?is)<meta\b(?=[^>]*\bproperty\s*=\s*["'']og:description["''])(?=[^>]*\bcontent\s*=\s*["'']([^"'']*)["''])[^>]*>')

    if(!$tm.Success){ $missingTitle++; continue }
    if(!$mm.Success){ $missingMeta++; continue }
    if(!$ot.Success){ $missingOgTitle++; continue }
    if(!$od.Success){ $missingOgDesc++; continue }

    $title = [System.Net.WebUtility]::HtmlDecode($tm.Groups[1].Value).Trim()
    $meta  = [System.Net.WebUtility]::HtmlDecode($mm.Groups[1].Value).Trim()
    $ogt   = [System.Net.WebUtility]::HtmlDecode($ot.Groups[1].Value).Trim()
    $ogd   = [System.Net.WebUtility]::HtmlDecode($od.Groups[1].Value).Trim()

    if($title -ne $ogt){ $titleOgMismatch++ }
    if($meta -ne $ogd){ $metaOgMismatch++ }
}

Write-Host ""
Write-Host "=== TITLE / META / OG v9 실제 적용 완료 ==="
Write-Host "사용 CSV:" $csv.FullName
Write-Host "적용 대상:" $rows.Count
Write-Host "실제 수정:" $modified
Write-Host "실패:" $failed.Count
Write-Host "TITLE 누락:" $missingTitle
Write-Host "META DESCRIPTION 누락:" $missingMeta
Write-Host "OG TITLE 누락:" $missingOgTitle
Write-Host "OG DESCRIPTION 누락:" $missingOgDesc
Write-Host "TITLE != OG TITLE:" $titleOgMismatch
Write-Host "META != OG DESCRIPTION:" $metaOgMismatch
Write-Host "백업 폴더:" $backup

if($failed.Count -gt 0){
    $report = Join-Path $dataDir "head-v9-apply-fail-$stamp.txt"
    $failed | Set-Content -LiteralPath $report -Encoding UTF8
    Write-Host "실패 보고서:" $report
}
