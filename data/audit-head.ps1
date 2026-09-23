$ErrorActionPreference = "Stop"

$root = "E:\mong-lounge-site"
$dataDir = Join-Path $root "data"
$stamp = Get-Date -Format "yyyyMMdd-HHmmss"

$detailReport  = Join-Path $dataDir "head-pattern-audit-detail-$stamp.csv"
$summaryReport = Join-Path $dataDir "head-pattern-audit-summary-$stamp.csv"

function Decode-Clean([string]$text) {
    if ([string]::IsNullOrWhiteSpace($text)) { return "" }
    $x = [System.Net.WebUtility]::HtmlDecode($text)
    $x = [regex]::Replace($x, '<[^>]+>', ' ')
    $x = [regex]::Replace($x, '\s+', ' ')
    return $x.Trim()
}

function Get-TagText([string]$html, [string]$tag) {
    $m = [regex]::Match($html, "(?is)<$tag\b[^>]*>(.*?)</$tag>")
    if ($m.Success) { return Decode-Clean $m.Groups[1].Value }
    return ""
}

function Get-MetaContent([string]$html, [string]$attr, [string]$value) {
    $p1 = "(?is)<meta\b(?=[^>]*\b$attr\s*=\s*[""']$([regex]::Escape($value))[""'])(?=[^>]*\bcontent\s*=\s*[""']([^""']*)[""'])[^>]*>"
    $p2 = "(?is)<meta\b(?=[^>]*\bcontent\s*=\s*[""']([^""']*)[""'])(?=[^>]*\b$attr\s*=\s*[""']$([regex]::Escape($value))[""'])[^>]*>"
    foreach($p in @($p1,$p2)) {
        $m = [regex]::Match($html,$p)
        if($m.Success) { return Decode-Clean $m.Groups[1].Value }
    }
    return ""
}

function Get-Canonical([string]$html) {
    $m = [regex]::Match($html,'(?is)<link\b(?=[^>]*\brel\s*=\s*["'']canonical["''])(?=[^>]*\bhref\s*=\s*["'']([^"'']+)["''])[^>]*>')
    if($m.Success){ return $m.Groups[1].Value.Trim() }
    $m = [regex]::Match($html,'(?is)<link\b(?=[^>]*\bhref\s*=\s*["'']([^"'']+)["''])(?=[^>]*\brel\s*=\s*["'']canonical["''])[^>]*>')
    if($m.Success){ return $m.Groups[1].Value.Trim() }
    return ""
}

function Count-KW([string]$text) {
    if ([string]::IsNullOrEmpty($text)) { return 0 }
    return ([regex]::Matches($text,'출장마사지',[Text.RegularExpressions.RegexOptions]::IgnoreCase)).Count
}

function Get-PageType([string]$relPath) {
    $p = $relPath.Replace("\","/").TrimStart("/")
    $lower = $p.ToLowerInvariant()

    if ($lower -eq "index.html") { return "메인" }

    if ($lower.StartsWith("shops/")) { return "업체상세" }

    $segments = @($lower -split "/" | Where-Object { $_ -and $_ -ne "index.html" })

    # 동/읍/면 계층
    foreach($s in $segments) {
        if ($s -match '(^|[-_])(dong|eup|myeon)([-_]|$)' -or
            $s -match '(동|읍|면)$') {
            return "동읍면"
        }
    }

    # 시/군/구 계층
    foreach($s in $segments) {
        if ($s -match '(^|[-_])(gu|si|gun)([-_]|$)' -or
            $s -match '(시|군|구)$') {
            return "시군구"
        }
    }

    # 시도/광역권
    $sidoTokens = @(
        "seoul","gyeonggi","incheon","daejeon","daegu","gwangju",
        "busan","ulsan","sejong","gangwon","chungbuk","chungnam",
        "jeonbuk","jeonnam","gyeongbuk","gyeongnam","jeju",
        "서울","경기","인천","대전","대구","광주","부산","울산","세종",
        "강원","충북","충남","전북","전남","경북","경남","제주"
    )
    foreach($s in $segments) {
        foreach($t in $sidoTokens) {
            if($s -eq $t -or $s.StartsWith("$t-") -or $s.EndsWith("-$t")) {
                return "시도"
            }
        }
    }

    return "기타"
}

if(!(Test-Path -LiteralPath $root)) {
    throw "사이트 폴더 없음: $root"
}

$files = @(
    Get-ChildItem -LiteralPath $root -Recurse -File -Filter "*.html" |
    Where-Object {
        $_.FullName -notmatch '\\\.git\\|\\node_modules\\|\\backup\\|\\_backup\\'
    }
)

$rows = New-Object System.Collections.Generic.List[object]

foreach($f in $files) {
    try {
        $html = Get-Content -LiteralPath $f.FullName -Raw -Encoding UTF8
        $rel = $f.FullName.Substring($root.Length).TrimStart("\")
        $title = Get-TagText $html "title"
        $desc = Get-MetaContent $html "name" "description"
        $ogTitle = Get-MetaContent $html "property" "og:title"
        $ogDesc = Get-MetaContent $html "property" "og:description"
        $h1 = Get-TagText $html "h1"
        $canonical = Get-Canonical $html
        $type = Get-PageType $rel

        $rows.Add([pscustomobject]@{
            Type = $type
            File = $rel.Replace("\","/")
            Title = $title
            Description = $desc
            OGTitle = $ogTitle
            OGDescription = $ogDesc
            H1 = $h1
            Canonical = $canonical
            TitleLength = $title.Length
            DescriptionLength = $desc.Length
            Title_KW = Count-KW $title
            Description_KW = Count-KW $desc
            OGTitle_KW = Count-KW $ogTitle
            OGDescription_KW = Count-KW $ogDesc
            H1_KW = Count-KW $h1
            TitleEqualsOG = ($title -and $ogTitle -and $title -eq $ogTitle)
            DescEqualsOG = ($desc -and $ogDesc -and $desc -eq $ogDesc)
            TitleMissing = [string]::IsNullOrWhiteSpace($title)
            DescMissing = [string]::IsNullOrWhiteSpace($desc)
        })
    }
    catch {
        Write-Warning "읽기 실패: $($f.FullName)"
    }
}

$rows | Export-Csv -LiteralPath $detailReport -NoTypeInformation -Encoding UTF8

$types = @("메인","시도","시군구","동읍면","업체상세","기타")
$summary = New-Object System.Collections.Generic.List[object]

foreach($type in $types) {
    $g = @($rows | Where-Object {$_.Type -eq $type})
    if($g.Count -eq 0){ continue }

    $summary.Add([pscustomobject]@{
        Type = $type
        Pages = $g.Count
        TitleMissing = @($g | Where-Object {$_.TitleMissing}).Count
        DescMissing = @($g | Where-Object {$_.DescMissing}).Count
        TitleKW0 = @($g | Where-Object {$_.Title_KW -eq 0}).Count
        TitleKW1 = @($g | Where-Object {$_.Title_KW -eq 1}).Count
        TitleKW2Plus = @($g | Where-Object {$_.Title_KW -ge 2}).Count
        DescKW0 = @($g | Where-Object {$_.Description_KW -eq 0}).Count
        DescKW1 = @($g | Where-Object {$_.Description_KW -eq 1}).Count
        DescKW2Plus = @($g | Where-Object {$_.Description_KW -ge 2}).Count
        TitleEqualsOG = @($g | Where-Object {$_.TitleEqualsOG}).Count
        DescEqualsOG = @($g | Where-Object {$_.DescEqualsOG}).Count
    })
}

$summary | Export-Csv -LiteralPath $summaryReport -NoTypeInformation -Encoding UTF8

Write-Host ""
Write-Host "=== 몽라운지 HEAD 패턴 읽기 전용 점검 ==="
Write-Host "전체 HTML:" $rows.Count
Write-Host "사이트 파일 수정: 0"
Write-Host ""

foreach($s in $summary) {
    Write-Host "[$($s.Type)]"
    Write-Host "페이지:" $s.Pages
    Write-Host "TITLE 출장마사지 0회 / 1회 / 2회이상:" "$($s.TitleKW0) / $($s.TitleKW1) / $($s.TitleKW2Plus)"
    Write-Host "META  출장마사지 0회 / 1회 / 2회이상:" "$($s.DescKW0) / $($s.DescKW1) / $($s.DescKW2Plus)"
    Write-Host "TITLE=OG TITLE:" "$($s.TitleEqualsOG) / $($s.Pages)"
    Write-Host "META=OG DESCRIPTION:" "$($s.DescEqualsOG) / $($s.Pages)"
    Write-Host "TITLE 누락 / META 누락:" "$($s.TitleMissing) / $($s.DescMissing)"
    Write-Host ""
}

foreach($type in @("메인","시도","시군구","동읍면","업체상세")) {
    $g = @($rows | Where-Object {$_.Type -eq $type})
    if($g.Count -eq 0){ continue }

    Write-Host "=== [$type] 샘플 최대 5개 ==="
    $g | Select-Object -First 5 | ForEach-Object {
        Write-Host "파일:" $_.File
        Write-Host "TITLE:" $_.Title
        Write-Host "META :" $_.Description
        Write-Host "OG   :" $_.OGTitle
        Write-Host "H1   :" $_.H1
        Write-Host "KW횟수 TITLE/META/OG/H1:" "$($_.Title_KW)/$($_.Description_KW)/$($_.OGTitle_KW)/$($_.H1_KW)"
        Write-Host "---"
    }
    Write-Host ""
}

Write-Host "상세 보고서:" $detailReport
Write-Host "요약 보고서:" $summaryReport
Write-Host ""
Write-Host "이 결과를 그대로 보내주시면 페이지 유형별로 어떤 패턴을 줄이고 어디에 출장마사지 키워드를 남길지 정리할 수 있습니다."
