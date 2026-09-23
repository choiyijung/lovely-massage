$ErrorActionPreference = "Stop"

$root = "E:\mong-lounge-site"
$dataDir = Join-Path $root "data"
$stamp = Get-Date -Format "yyyyMMdd-HHmmss"
$backup = "E:\mong-lounge-head-cleanup-backup-$stamp"

New-Item -ItemType Directory -Force -Path $backup | Out-Null

function HtmlEncodeAttr([string]$s){
    if($null -eq $s){ return "" }
    return [System.Net.WebUtility]::HtmlEncode($s)
}

function DecodeClean([string]$s){
    if([string]::IsNullOrWhiteSpace($s)){ return "" }
    $x=[System.Net.WebUtility]::HtmlDecode($s)
    $x=[regex]::Replace($x,'\s+',' ')
    return $x.Trim()
}

function Get-Title([string]$html){
    $m=[regex]::Match($html,'(?is)<title\b[^>]*>(.*?)</title>')
    if($m.Success){ return DecodeClean $m.Groups[1].Value }
    return ""
}

function Get-Meta([string]$html,[string]$attr,[string]$name){
    foreach($p in @(
        "(?is)<meta\b(?=[^>]*\b$attr\s*=\s*[""']$([regex]::Escape($name))[""'])(?=[^>]*\bcontent\s*=\s*[""']([^""']*)[""'])[^>]*>",
        "(?is)<meta\b(?=[^>]*\bcontent\s*=\s*[""']([^""']*)[""'])(?=[^>]*\b$attr\s*=\s*[""']$([regex]::Escape($name))[""'])[^>]*>"
    )){
        $m=[regex]::Match($html,$p)
        if($m.Success){ return DecodeClean $m.Groups[1].Value }
    }
    return ""
}

function Set-Meta([string]$html,[string]$attr,[string]$name,[string]$value){
    $encoded = HtmlEncodeAttr $value
    $replacement = "<meta $attr=`"$name`" content=`"$encoded`">"

    $patterns=@(
        "(?is)<meta\b(?=[^>]*\b$attr\s*=\s*[""']$([regex]::Escape($name))[""'])(?=[^>]*\bcontent\s*=\s*[""'][^""']*[""'])[^>]*>",
        "(?is)<meta\b(?=[^>]*\bcontent\s*=\s*[""'][^""']*[""'])(?=[^>]*\b$attr\s*=\s*[""']$([regex]::Escape($name))[""'])[^>]*>"
    )

    foreach($p in $patterns){
        if([regex]::IsMatch($html,$p)){
            return [regex]::Replace($html,$p,$replacement,1)
        }
    }

    if($html -match '(?is)</head>'){
        return [regex]::Replace($html,'(?is)</head>',"    $replacement`r`n</head>",1)
    }

    return $html
}

function Backup-File([string]$fullPath){
    $rel=$fullPath.Substring($root.Length).TrimStart("\")
    $dest=Join-Path $backup $rel
    $destDir=Split-Path $dest -Parent
    New-Item -ItemType Directory -Force -Path $destDir | Out-Null
    Copy-Item -LiteralPath $fullPath -Destination $dest -Force
}

$grammarMap=[ordered]@{
    "피로을"="피로를"
    "통증을을"="통증을"
    "긴장을을"="긴장을"
    "붓기을"="붓기를"
    "피로가을"="피로를"
    "불편을을"="불편을"
}

$grammarModified=0
$grammarHits=0

$massageFiles=@(
    Get-ChildItem -LiteralPath (Join-Path $root "massage") -Recurse -File -Filter "index.html"
)

foreach($f in $massageFiles){
    $html=Get-Content -LiteralPath $f.FullName -Raw -Encoding UTF8
    $title=Get-Title $html
    $desc=Get-Meta $html "name" "description"

    $newTitle=$title
    $newDesc=$desc

    foreach($k in $grammarMap.Keys){
        if($newTitle.Contains($k)){ $grammarHits++; $newTitle=$newTitle.Replace($k,$grammarMap[$k]) }
        if($newDesc.Contains($k)){ $grammarHits++; $newDesc=$newDesc.Replace($k,$grammarMap[$k]) }
    }

    if($newTitle -ne $title -or $newDesc -ne $desc){
        Backup-File $f.FullName
        $before=$html

        if($newTitle -ne $title){
            $encoded=[System.Net.WebUtility]::HtmlEncode($newTitle)
            $html=[regex]::Replace($html,'(?is)<title\b[^>]*>.*?</title>',"<title>$encoded</title>",1)
            $html=Set-Meta $html "property" "og:title" $newTitle
        }

        if($newDesc -ne $desc){
            $html=Set-Meta $html "name" "description" $newDesc
            $html=Set-Meta $html "property" "og:description" $newDesc
        }

        if($html -ne $before){
            Set-Content -LiteralPath $f.FullName -Value $html -Encoding UTF8
            $grammarModified++
        }
    }
}

$shopFiles=@(
    Get-ChildItem -LiteralPath (Join-Path $root "shops") -Recurse -File -Filter "index.html"
)

$shopModified=0
$shopSkipped=0

foreach($f in $shopFiles){
    $html=Get-Content -LiteralPath $f.FullName -Raw -Encoding UTF8
    $title=Get-Title $html
    $desc=Get-Meta $html "name" "description"
    $ogTitle=Get-Meta $html "property" "og:title"
    $ogDesc=Get-Meta $html "property" "og:description"

    if([string]::IsNullOrWhiteSpace($title) -or [string]::IsNullOrWhiteSpace($desc)){
        $shopSkipped++
        continue
    }

    if($ogTitle -ne $title -or $ogDesc -ne $desc){
        Backup-File $f.FullName
        $html=Set-Meta $html "property" "og:title" $title
        $html=Set-Meta $html "property" "og:description" $desc
        Set-Content -LiteralPath $f.FullName -Value $html -Encoding UTF8
        $shopModified++
    }
}

# 최종 검증
$grammarRemaining=0
$titleOgMismatch=0
$descOgMismatch=0
$shopCount=0

foreach($f in $massageFiles){
    $html=Get-Content -LiteralPath $f.FullName -Raw -Encoding UTF8
    $title=Get-Title $html
    $desc=Get-Meta $html "name" "description"

    foreach($k in $grammarMap.Keys){
        if($title.Contains($k) -or $desc.Contains($k)){ $grammarRemaining++ }
    }
}

foreach($f in $shopFiles){
    $shopCount++
    $html=Get-Content -LiteralPath $f.FullName -Raw -Encoding UTF8
    $title=Get-Title $html
    $desc=Get-Meta $html "name" "description"
    $ogTitle=Get-Meta $html "property" "og:title"
    $ogDesc=Get-Meta $html "property" "og:description"

    if($title -ne $ogTitle){ $titleOgMismatch++ }
    if($desc -ne $ogDesc){ $descOgMismatch++ }
}

Write-Host ""
Write-Host "=== HEAD 후속 정리 완료 ==="
Write-Host "문법 수정 페이지:" $grammarModified
Write-Host "문법 의심 남음:" $grammarRemaining
Write-Host ""
Write-Host "업체상세 전체:" $shopCount
Write-Host "업체 OG 수정:" $shopModified
Write-Host "업체 OG 건너뜀:" $shopSkipped
Write-Host "TITLE != OG TITLE:" $titleOgMismatch
Write-Host "META != OG DESCRIPTION:" $descOgMismatch
Write-Host ""
Write-Host "DESCRIPTION 누락 88개는 이동/연결 69 + 기능/보조 19라 이번 작업에서 유지"
Write-Host "백업 폴더:" $backup
