$ErrorActionPreference = "Stop"

$root = "E:\mong-lounge-site"
$dataDir = Join-Path $root "data"
$stamp = Get-Date -Format "yyyyMMdd-HHmmss"

$detailReport  = Join-Path $dataDir "near-duplicate-head-detail-$stamp.csv"
$summaryReport = Join-Path $dataDir "near-duplicate-head-summary-$stamp.csv"

function Clean([string]$s){
    if([string]::IsNullOrWhiteSpace($s)){ return "" }
    $x=[System.Net.WebUtility]::HtmlDecode($s)
    $x=[regex]::Replace($x,'<[^>]+>',' ')
    $x=[regex]::Replace($x,'\s+',' ')
    return $x.Trim()
}

function Get-Title([string]$html){
    $m=[regex]::Match($html,'(?is)<title\b[^>]*>(.*?)</title>')
    if($m.Success){ return Clean $m.Groups[1].Value }
    return ""
}

function Get-Description([string]$html){
    foreach($p in @(
        '(?is)<meta\b(?=[^>]*\bname\s*=\s*["'']description["''])(?=[^>]*\bcontent\s*=\s*["'']([^"'']*)["''])[^>]*>',
        '(?is)<meta\b(?=[^>]*\bcontent\s*=\s*["'']([^"'']*)["''])(?=[^>]*\bname\s*=\s*["'']description["''])[^>]*>'
    )){
        $m=[regex]::Match($html,$p)
        if($m.Success){ return Clean $m.Groups[1].Value }
    }
    return ""
}

function Page-Class([string]$rel){
    $p=$rel.Replace("\","/")
    if($p -like "shops/*"){ return "업체상세" }
    if($p -like "massage/*"){ return "지역SEO" }
    return "기타"
}

$generic = New-Object System.Collections.Generic.HashSet[string]([StringComparer]::OrdinalIgnoreCase)
@(
    "massage","shops","index","html","출장마사지","마사지","몽라운지",
    "생활권","지역","지역정보","이용안내","안내","이용","정보",
    "방문","홈케어","업체","상세","상세정보","관리","추천","코스","가격",
    "seoul","gyeonggi","incheon","other","daejeon","daegu","gwangju",
    "gangnam","gangseo","gangdong","gangbuk","gu","si","gun","dong","eup","myeon"
) | ForEach-Object { [void]$generic.Add($_) }

function Get-PathTokens([string]$rel){
    $decoded=$rel
    try { $decoded=[uri]::UnescapeDataString($decoded) } catch {}
    $parts=@(
        $decoded -split '[\\/_.\-\s·|]+' |
        ForEach-Object { $_.Trim() } |
        Where-Object {
            $_.Length -ge 2 -and
            -not $generic.Contains($_) -and
            $_ -notmatch '^\d+$'
        }
    )
    return @($parts | Sort-Object Length -Descending -Unique)
}

function Normalize-Template([string]$text,[string[]]$tokens){
    if([string]::IsNullOrWhiteSpace($text)){ return "" }

    $x=Clean $text

    # 브랜드/핵심키워드는 구조를 보기 위해 고정 토큰으로 유지
    $x=$x.Replace("몽라운지","{BRAND}")
    $x=$x.Replace("출장마사지","{KW}")

    # 해당 파일 경로에서 파생된 지역명·업체명·상황어를 변수로 마스킹
    foreach($t in $tokens){
        if($t -eq "{KW}" -or $t -eq "{BRAND}") { continue }
        if($t.Length -lt 2){ continue }
        $x=[regex]::Replace($x,[regex]::Escape($t),'{X}')
    }

    # 남아있는 전형적인 행정구역 토큰도 마스킹
    $x=[regex]::Replace($x,'[가-힣A-Za-z0-9]+(?:특별시|광역시|특별자치시|특별자치도|도|시|군|구|읍|면|동|리)\b','{LOC}')
    $x=[regex]::Replace($x,'\d+(?:개|곳|분|시간|원|만원)?','{N}')

    # 문장 구조만 비교할 수 있게 공백/구두점 정규화
    $x=$x.ToLowerInvariant()
    $x=[regex]::Replace($x,'[“”"''`~!@#$%^&*()_=+\[\]{}<>?/\\:;,.·|–—-]+',' ')
    $x=[regex]::Replace($x,'(?:\{x\}\s*){2,}','{X} ')
    $x=[regex]::Replace($x,'(?:\{loc\}\s*){2,}','{LOC} ')
    $x=[regex]::Replace($x,'\s+',' ')
    return $x.Trim()
}

$files=@(
    Get-ChildItem -LiteralPath $root -Recurse -File -Filter "*.html" |
    Where-Object { $_.FullName -notmatch '\\\.git\\|\\node_modules\\|\\backup\\|\\_backup\\' }
)

$rows=New-Object System.Collections.Generic.List[object]

foreach($f in $files){
    try{
        $rel=$f.FullName.Substring($root.Length).TrimStart("\").Replace("\","/")
        $html=Get-Content -LiteralPath $f.FullName -Raw -Encoding UTF8
        $title=Get-Title $html
        $desc=Get-Description $html
        $tokens=Get-PathTokens $rel

        $rows.Add([pscustomobject]@{
            Class=Page-Class $rel
            File=$rel
            Title=$title
            Description=$desc
            TitleTemplate=Normalize-Template $title $tokens
            DescriptionTemplate=Normalize-Template $desc $tokens
        })
    }
    catch{}
}

$titleGroups=@(
    $rows |
    Where-Object {$_.TitleTemplate} |
    Group-Object Class,TitleTemplate |
    Where-Object {$_.Count -gt 1} |
    Sort-Object Count -Descending
)

$descGroups=@(
    $rows |
    Where-Object {$_.DescriptionTemplate} |
    Group-Object Class,DescriptionTemplate |
    Where-Object {$_.Count -gt 1} |
    Sort-Object Count -Descending
)

$detail=New-Object System.Collections.Generic.List[object]
$gid=0

foreach($g in $titleGroups){
    $gid++
    $sample=@($g.Group|Select-Object -First 3)
    $detail.Add([pscustomobject]@{
        Type="TITLE_TEMPLATE"
        Group=$gid
        Class=$sample[0].Class
        Count=$g.Count
        Template=$sample[0].TitleTemplate
        Sample1=if($sample.Count -ge 1){"$($sample[0].File) || $($sample[0].Title)"}else{""}
        Sample2=if($sample.Count -ge 2){"$($sample[1].File) || $($sample[1].Title)"}else{""}
        Sample3=if($sample.Count -ge 3){"$($sample[2].File) || $($sample[2].Title)"}else{""}
    })
}

$gid=0
foreach($g in $descGroups){
    $gid++
    $sample=@($g.Group|Select-Object -First 3)
    $detail.Add([pscustomobject]@{
        Type="DESCRIPTION_TEMPLATE"
        Group=$gid
        Class=$sample[0].Class
        Count=$g.Count
        Template=$sample[0].DescriptionTemplate
        Sample1=if($sample.Count -ge 1){"$($sample[0].File) || $($sample[0].Description)"}else{""}
        Sample2=if($sample.Count -ge 2){"$($sample[1].File) || $($sample[1].Description)"}else{""}
        Sample3=if($sample.Count -ge 3){"$($sample[2].File) || $($sample[2].Description)"}else{""}
    })
}

$detail | Export-Csv -LiteralPath $detailReport -NoTypeInformation -Encoding UTF8

$summary=New-Object System.Collections.Generic.List[object]
foreach($class in @("지역SEO","업체상세","기타")){
    $g=@($rows|Where-Object {$_.Class -eq $class})
    if($g.Count -eq 0){continue}

    $tGroups=@($titleGroups|Where-Object {$_.Group[0].Class -eq $class})
    $dGroups=@($descGroups|Where-Object {$_.Group[0].Class -eq $class})

    $tAffected=New-Object System.Collections.Generic.HashSet[string]
    foreach($tg in $tGroups){ foreach($r in $tg.Group){ [void]$tAffected.Add($r.File) } }

    $dAffected=New-Object System.Collections.Generic.HashSet[string]
    foreach($dg in $dGroups){ foreach($r in $dg.Group){ [void]$dAffected.Add($r.File) } }

    $summary.Add([pscustomobject]@{
        Class=$class
        Pages=$g.Count
        TitleTemplateGroups=$tGroups.Count
        TitleTemplateAffectedPages=$tAffected.Count
        DescriptionTemplateGroups=$dGroups.Count
        DescriptionTemplateAffectedPages=$dAffected.Count
    })
}

$summary | Export-Csv -LiteralPath $summaryReport -NoTypeInformation -Encoding UTF8

Write-Host ""
Write-Host "=== TITLE / META 유사 템플릿 읽기 전용 점검 ==="
Write-Host "사이트 파일 수정: 0"
Write-Host "전체 HTML:" $rows.Count
Write-Host ""

foreach($s in $summary){
    Write-Host "[$($s.Class)]"
    Write-Host "페이지:" $s.Pages
    Write-Host "유사 TITLE 템플릿 그룹:" $s.TitleTemplateGroups
    Write-Host "유사 TITLE 영향 페이지:" $s.TitleTemplateAffectedPages
    Write-Host "유사 META 템플릿 그룹:" $s.DescriptionTemplateGroups
    Write-Host "유사 META 영향 페이지:" $s.DescriptionTemplateAffectedPages
    Write-Host ""
}

Write-Host "=== TITLE 유사 템플릿 상위 10개 ==="
$titleGroups | Select-Object -First 10 | ForEach-Object {
    $sample=@($_.Group|Select-Object -First 2)
    Write-Host "$($_.Count)개 | [$($sample[0].Class)]"
    Write-Host " 1) $($sample[0].Title)"
    if($sample.Count -ge 2){ Write-Host " 2) $($sample[1].Title)" }
}

Write-Host ""
Write-Host "=== META 유사 템플릿 상위 10개 ==="
$descGroups | Select-Object -First 10 | ForEach-Object {
    $sample=@($_.Group|Select-Object -First 2)
    Write-Host "$($_.Count)개 | [$($sample[0].Class)]"
    Write-Host " 1) $($sample[0].Description)"
    if($sample.Count -ge 2){ Write-Host " 2) $($sample[1].Description)" }
}

Write-Host ""
Write-Host "상세 보고서:" $detailReport
Write-Host "요약 보고서:" $summaryReport
