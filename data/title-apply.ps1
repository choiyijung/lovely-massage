$ErrorActionPreference = "Stop"

$root = "E:\mong-lounge-site"
$dataDir = Join-Path $root "data"
$stamp = Get-Date -Format "yyyyMMdd-HHmmss"

$preview = Get-ChildItem -LiteralPath $dataDir -File |
    Where-Object { $_.Name -like "title-meta-rewrite-preview-*.csv" } |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1

if(!$preview){
    throw "title-meta-rewrite-preview-*.csv 파일을 찾지 못했습니다."
}

$rows = @(Import-Csv -LiteralPath $preview.FullName | Where-Object { $_.Action -eq "변경후보" })

if($rows.Count -eq 0){
    throw "변경후보가 없습니다."
}

$backup = "E:\mong-lounge-title-meta-backup-$stamp"
New-Item -ItemType Directory -Force -Path $backup | Out-Null

function HtmlEncodeAttr([string]$s){
    if($null -eq $s){ return "" }
    return [System.Net.WebUtility]::HtmlEncode($s)
}

function Set-Title([string]$html,[string]$value){
    $encoded = HtmlEncodeAttr $value
    if([regex]::IsMatch($html,'(?is)<title\b[^>]*>.*?</title>')){
        return [regex]::Replace($html,'(?is)<title\b[^>]*>.*?</title>',"<title>$encoded</title>",1)
    }
    return $html
}

function Set-Meta([string]$html,[string]$attr,[string]$name,[string]$value){
    $encoded = HtmlEncodeAttr $value

    $p1 = "(?is)<meta\b(?=[^>]*\b$attr\s*=\s*[""']$([regex]::Escape($name))[""'])(?=[^>]*\bcontent\s*=\s*[""'][^""']*[""'])[^>]*>"
    $p2 = "(?is)<meta\b(?=[^>]*\bcontent\s*=\s*[""'][^""']*[""'])(?=[^>]*\b$attr\s*=\s*[""']$([regex]::Escape($name))[""'])[^>]*>"

    $replacement = "<meta $attr=`"$name`" content=`"$encoded`">"

    if([regex]::IsMatch($html,$p1)){
        return [regex]::Replace($html,$p1,$replacement,1)
    }
    if([regex]::IsMatch($html,$p2)){
        return [regex]::Replace($html,$p2,$replacement,1)
    }

    # meta description/OG가 없으면 </head> 직전에 추가
    if($html -match '(?is)</head>'){
        return [regex]::Replace($html,'(?is)</head>',"    $replacement`r`n</head>",1)
    }

    return $html
}

$modified = 0
$failed = New-Object System.Collections.Generic.List[string]

foreach($r in $rows){
    try{
        $path = Join-Path $root ($r.File.Replace("/","\"))
        if(!(Test-Path -LiteralPath $path)){
            $failed.Add("파일없음: $($r.File)")
            continue
        }

        $relDir = Split-Path $r.File -Parent
        $backupDir = if($relDir){ Join-Path $backup $relDir } else { $backup }
        New-Item -ItemType Directory -Force -Path $backupDir | Out-Null
        Copy-Item -LiteralPath $path -Destination (Join-Path $backupDir (Split-Path $path -Leaf)) -Force

        $html = Get-Content -LiteralPath $path -Raw -Encoding UTF8
        $before = $html

        $html = Set-Title $html $r.ProposedTitle
        $html = Set-Meta $html "name" "description" $r.ProposedDescription

        # OG는 TITLE/META와 동일하게 맞춤
        $html = Set-Meta $html "property" "og:title" $r.ProposedTitle
        $html = Set-Meta $html "property" "og:description" $r.ProposedDescription

        if($html -ne $before){
            Set-Content -LiteralPath $path -Value $html -Encoding UTF8
            $modified++
        }
    }
    catch{
        $failed.Add("$($r.File) | $($_.Exception.Message)")
    }
}

# 검증
$titleKw0 = 0
$titleKw1 = 0
$descKw0 = 0
$descKw1 = 0
$titleOgMismatch = 0
$descOgMismatch = 0

function Decode-Clean([string]$text) {
    if ([string]::IsNullOrWhiteSpace($text)) { return "" }
    $x = [System.Net.WebUtility]::HtmlDecode($text)
    $x = [regex]::Replace($x, '\s+', ' ')
    return $x.Trim()
}

function Get-Title([string]$html){
    $m=[regex]::Match($html,'(?is)<title\b[^>]*>(.*?)</title>')
    if($m.Success){ return Decode-Clean $m.Groups[1].Value }
    return ""
}
function Get-MetaValue([string]$html,[string]$attr,[string]$name){
    foreach($p in @(
        "(?is)<meta\b(?=[^>]*\b$attr\s*=\s*[""']$([regex]::Escape($name))[""'])(?=[^>]*\bcontent\s*=\s*[""']([^""']*)[""'])[^>]*>",
        "(?is)<meta\b(?=[^>]*\bcontent\s*=\s*[""']([^""']*)[""'])(?=[^>]*\b$attr\s*=\s*[""']$([regex]::Escape($name))[""'])[^>]*>"
    )){
        $m=[regex]::Match($html,$p)
        if($m.Success){ return Decode-Clean $m.Groups[1].Value }
    }
    return ""
}

foreach($r in $rows){
    $path = Join-Path $root ($r.File.Replace("/","\"))
    if(!(Test-Path -LiteralPath $path)){ continue }

    $html=Get-Content -LiteralPath $path -Raw -Encoding UTF8
    $t=Get-Title $html
    $d=Get-MetaValue $html "name" "description"
    $ot=Get-MetaValue $html "property" "og:title"
    $od=Get-MetaValue $html "property" "og:description"

    $tk=([regex]::Matches($t,'출장마사지')).Count
    $dk=([regex]::Matches($d,'출장마사지')).Count

    if($tk -eq 0){$titleKw0++}
    elseif($tk -eq 1){$titleKw1++}

    if($dk -eq 0){$descKw0++}
    elseif($dk -eq 1){$descKw1++}

    if($t -ne $ot){$titleOgMismatch++}
    if($d -ne $od){$descOgMismatch++}
}

Write-Host ""
Write-Host "=== TITLE / META 실제 적용 완료 ==="
Write-Host "사용한 미리보기:" $preview.FullName
Write-Host "변경후보:" $rows.Count
Write-Host "실제 수정:" $modified
Write-Host "실패:" $failed.Count
Write-Host ""
Write-Host "TITLE 출장마사지 0회:" $titleKw0
Write-Host "TITLE 출장마사지 1회:" $titleKw1
Write-Host "META 출장마사지 0회:" $descKw0
Write-Host "META 출장마사지 1회:" $descKw1
Write-Host "TITLE != OG TITLE:" $titleOgMismatch
Write-Host "META != OG DESCRIPTION:" $descOgMismatch
Write-Host ""
Write-Host "백업 폴더:" $backup

if($failed.Count -gt 0){
    $failReport=Join-Path $dataDir "title-meta-apply-fail-$stamp.txt"
    $failed | Set-Content -LiteralPath $failReport -Encoding UTF8
    Write-Host "실패 보고서:" $failReport
}
