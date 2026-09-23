$ErrorActionPreference = "Stop"

$root = "E:\mong-lounge-site"
$dataDir = Join-Path $root "data"
$stamp = Get-Date -Format "yyyyMMdd-HHmmss"

$detailReport  = Join-Path $dataDir "canonical-audit-detail-$stamp.csv"
$summaryReport = Join-Path $dataDir "canonical-audit-summary-$stamp.csv"

function Clean([string]$s){
    if([string]::IsNullOrWhiteSpace($s)){ return "" }
    return ([System.Net.WebUtility]::HtmlDecode($s)).Trim()
}

function Get-Title([string]$html){
    $m=[regex]::Match($html,'(?is)<title\b[^>]*>(.*?)</title>')
    if($m.Success){
        $t=[System.Net.WebUtility]::HtmlDecode($m.Groups[1].Value)
        return ([regex]::Replace($t,'\s+',' ')).Trim()
    }
    return ""
}

function Get-Canonical([string]$html){
    $patterns=@(
        '(?is)<link\b(?=[^>]*\brel\s*=\s*["'']canonical["''])(?=[^>]*\bhref\s*=\s*["'']([^"'']+)["''])[^>]*>',
        '(?is)<link\b(?=[^>]*\bhref\s*=\s*["'']([^"'']+)["''])(?=[^>]*\brel\s*=\s*["'']canonical["''])[^>]*>'
    )
    foreach($p in $patterns){
        $m=[regex]::Match($html,$p)
        if($m.Success){ return Clean $m.Groups[1].Value }
    }
    return ""
}

function Get-CanonicalCount([string]$html){
    return ([regex]::Matches(
        $html,
        '(?is)<link\b[^>]*\brel\s*=\s*["'']canonical["''][^>]*>'
    )).Count
}

function Local-ExpectedPath([string]$rel){
    $p=$rel.Replace("\","/")
    if($p -eq "index.html"){ return "/" }
    if($p.EndsWith("/index.html")){
        return "/" + $p.Substring(0,$p.Length-"index.html".Length)
    }
    return "/" + $p
}

function Norm-Path([string]$path){
    if([string]::IsNullOrWhiteSpace($path)){ return "" }
    try { $path=[uri]::UnescapeDataString($path) } catch {}
    $path=$path.Replace("\","/")
    $path=[regex]::Replace($path,'/{2,}','/')
    if(!$path.StartsWith("/")){ $path="/"+$path }
    if($path -ne "/" -and !$path.EndsWith("/") -and !$path.EndsWith(".html")){
        $path+="/"
    }
    return $path
}

function Classify([string]$rel,[string]$title,[string]$html){
    if($rel -like "shops/*"){ return "업체상세" }
    if($rel -like "massage/*"){ return "지역SEO" }

    if($title -match '상세페이지 이동' -or
       $html -match '(?is)<meta[^>]+http-equiv\s*=\s*["'']refresh["'']'){
        return "이동/연결"
    }

    if($rel -match '(^|/)(blank-|login\.html$|signup\.html$|register\.html$|mypage\.html$|privacy\.html$|terms\.html$|search\.html$|notice\.html$|partner\.html$|event\.html$|dong-check\.html$)'){
        return "기능/보조"
    }

    return "기타"
}

if(!(Test-Path -LiteralPath $root)){ throw "사이트 폴더 없음: $root" }

$files=@(
    Get-ChildItem -LiteralPath $root -Recurse -File -Filter "*.html" |
    Where-Object { $_.FullName -notmatch '\\\.git\\|\\node_modules\\|\\backup\\|\\_backup\\' }
)

# 1차: canonical 수집
$tempRows=New-Object System.Collections.Generic.List[object]
$hosts=New-Object System.Collections.Generic.List[string]

foreach($f in $files){
    $rel=$f.FullName.Substring($root.Length).TrimStart("\").Replace("\","/")
    $html=Get-Content -LiteralPath $f.FullName -Raw -Encoding UTF8
    $title=Get-Title $html
    $canonical=Get-Canonical $html
    $cc=Get-CanonicalCount $html
    $class=Classify $rel $title $html

    $canonHost=""
    $scheme=""
    $canonPath=""
    $absolute=$false

    if($canonical){
        try{
            $u=[uri]$canonical
            if($u.IsAbsoluteUri){
                $absolute=$true
                $canonHost=$u.Host.ToLowerInvariant()
                $scheme=$u.Scheme.ToLowerInvariant()
                $canonPath=$u.AbsolutePath
                if($canonHost){ $hosts.Add($canonHost) }
            }
        }catch{}
    }

    $tempRows.Add([pscustomobject]@{
        File=$rel
        Class=$class
        Title=$title
        Canonical=$canonical
        CanonicalTagCount=$cc
        Absolute=$absolute
        Scheme=$scheme
        Host=$canonHost
        CanonicalPath=$canonPath
        ExpectedPath=Local-ExpectedPath $rel
    })
}

$dominantHost=""
if($hosts.Count -gt 0){
    $dominantHost=(
        $hosts |
        Group-Object |
        Sort-Object Count -Descending |
        Select-Object -First 1
    ).Name
}

# 2차: 비교
$rows=New-Object System.Collections.Generic.List[object]

foreach($r in $tempRows){
    $missing=[string]::IsNullOrWhiteSpace($r.Canonical)
    $relative=(!$missing -and !$r.Absolute)
    $wrongScheme=(!$missing -and $r.Absolute -and $r.Scheme -ne "https")
    $wrongHost=(!$missing -and $r.Absolute -and $dominantHost -and $r.Host -ne $dominantHost)

    $pathMatch=$false
    if(!$missing -and $r.Absolute){
        $pathMatch=(Norm-Path $r.CanonicalPath) -eq (Norm-Path $r.ExpectedPath)
    }

    $rows.Add([pscustomobject]@{
        File=$r.File
        Class=$r.Class
        Title=$r.Title
        Canonical=$r.Canonical
        CanonicalTagCount=$r.CanonicalTagCount
        Missing=$missing
        RelativeCanonical=$relative
        NonHttps=$wrongScheme
        Host=$r.Host
        WrongHost=$wrongHost
        ExpectedPath=(Norm-Path $r.ExpectedPath)
        CanonicalPath=(Norm-Path $r.CanonicalPath)
        SelfPathMatch=$pathMatch
    })
}

# canonical 값 자체 중복
$dupGroups=@(
    $rows |
    Where-Object { -not $_.Missing } |
    Group-Object Canonical |
    Where-Object { $_.Count -gt 1 } |
    Sort-Object Count -Descending
)

$dupFiles=New-Object System.Collections.Generic.HashSet[string]
foreach($g in $dupGroups){
    foreach($x in $g.Group){ [void]$dupFiles.Add($x.File) }
}

$final=foreach($r in $rows){
    [pscustomobject]@{
        File=$r.File
        Class=$r.Class
        Title=$r.Title
        Canonical=$r.Canonical
        CanonicalTagCount=$r.CanonicalTagCount
        Missing=$r.Missing
        DuplicateCanonical=$dupFiles.Contains($r.File)
        RelativeCanonical=$r.RelativeCanonical
        NonHttps=$r.NonHttps
        Host=$r.Host
        WrongHost=$r.WrongHost
        ExpectedPath=$r.ExpectedPath
        CanonicalPath=$r.CanonicalPath
        SelfPathMatch=$r.SelfPathMatch
    }
}

$final | Export-Csv -LiteralPath $detailReport -NoTypeInformation -Encoding UTF8

$summary=@(
    [pscustomobject]@{Item="전체 HTML";Count=$final.Count},
    [pscustomobject]@{Item="대표 canonical 호스트";Count=$dominantHost},
    [pscustomobject]@{Item="canonical 누락";Count=@($final|Where-Object{$_.Missing}).Count},
    [pscustomobject]@{Item="canonical 태그 2개 이상";Count=@($final|Where-Object{[int]$_.CanonicalTagCount -ge 2}).Count},
    [pscustomobject]@{Item="상대경로 canonical";Count=@($final|Where-Object{$_.RelativeCanonical}).Count},
    [pscustomobject]@{Item="HTTP canonical";Count=@($final|Where-Object{$_.NonHttps}).Count},
    [pscustomobject]@{Item="다른 호스트 canonical";Count=@($final|Where-Object{$_.WrongHost}).Count},
    [pscustomobject]@{Item="canonical 중복 그룹";Count=$dupGroups.Count},
    [pscustomobject]@{Item="canonical 중복 페이지";Count=$dupFiles.Count},
    [pscustomobject]@{Item="자기 경로 불일치";Count=@($final|Where-Object{-not $_.Missing -and $_.Absolute -ne $false -and -not $_.SelfPathMatch}).Count}
)
$summary | Export-Csv -LiteralPath $summaryReport -NoTypeInformation -Encoding UTF8

Write-Host ""
Write-Host "=== CANONICAL 읽기 전용 최종 점검 ==="
Write-Host "사이트 파일 수정: 0"
Write-Host "전체 HTML:" $final.Count
Write-Host "대표 canonical 호스트:" $dominantHost
Write-Host ""
Write-Host "canonical 누락:" @($final|Where-Object{$_.Missing}).Count
Write-Host "canonical 태그 2개 이상:" @($final|Where-Object{[int]$_.CanonicalTagCount -ge 2}).Count
Write-Host "상대경로 canonical:" @($final|Where-Object{$_.RelativeCanonical}).Count
Write-Host "HTTP canonical:" @($final|Where-Object{$_.NonHttps}).Count
Write-Host "다른 호스트 canonical:" @($final|Where-Object{$_.WrongHost}).Count
Write-Host "canonical 중복 그룹:" $dupGroups.Count
Write-Host "canonical 중복 페이지:" $dupFiles.Count
Write-Host "자기 경로 불일치:" @($final|Where-Object{-not $_.Missing -and -not $_.RelativeCanonical -and -not $_.SelfPathMatch}).Count

Write-Host ""
Write-Host "=== 실제 SEO 페이지별 문제 수 ==="
foreach($c in @("지역SEO","업체상세","이동/연결","기능/보조","기타")){
    $g=@($final|Where-Object{$_.Class -eq $c})
    if($g.Count -eq 0){continue}
    $problem=@($g|Where-Object{
        $_.Missing -or $_.DuplicateCanonical -or $_.RelativeCanonical -or
        $_.NonHttps -or $_.WrongHost -or ([int]$_.CanonicalTagCount -ge 2) -or
        (-not $_.Missing -and -not $_.RelativeCanonical -and -not $_.SelfPathMatch)
    }).Count
    Write-Host "$c : $problem / $($g.Count)"
}

if($dupGroups.Count -gt 0){
    Write-Host ""
    Write-Host "=== 중복 canonical 상위 10개 ==="
    $dupGroups|Select-Object -First 10|ForEach-Object{
        Write-Host "$($_.Count)개 | $($_.Name)"
    }
}

Write-Host ""
Write-Host "상세 보고서:" $detailReport
Write-Host "요약 보고서:" $summaryReport
