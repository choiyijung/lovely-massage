$ErrorActionPreference = "Stop"

$root = "E:\mong-lounge-site"
$dataDir = Join-Path $root "data"
$stamp = Get-Date -Format "yyyyMMdd-HHmmss"
$out = Join-Path $dataDir "title-meta-rewrite-preview-$stamp.csv"

function Clean([string]$s){
    if([string]::IsNullOrWhiteSpace($s)){ return "" }
    $s = [System.Net.WebUtility]::HtmlDecode($s)
    $s = [regex]::Replace($s,'\s+',' ')
    return $s.Trim()
}

function Get-Title([string]$html){
    $m=[regex]::Match($html,'(?is)<title\b[^>]*>(.*?)</title>')
    if($m.Success){ return Clean $m.Groups[1].Value }
    return ""
}

function Get-Desc([string]$html){
    foreach($p in @(
        '(?is)<meta\b(?=[^>]*\bname\s*=\s*["'']description["''])(?=[^>]*\bcontent\s*=\s*["'']([^"'']*)["''])[^>]*>',
        '(?is)<meta\b(?=[^>]*\bcontent\s*=\s*["'']([^"'']*)["''])(?=[^>]*\bname\s*=\s*["'']description["''])[^>]*>'
    )){
        $m=[regex]::Match($html,$p)
        if($m.Success){ return Clean $m.Groups[1].Value }
    }
    return ""
}

function Stable-Mod([string]$text,[int]$mod){
    $bytes=[Text.Encoding]::UTF8.GetBytes($text)
    $sum=0
    foreach($b in $bytes){ $sum=($sum+$b)%2147483647 }
    return $sum % $mod
}

$files = Get-ChildItem -LiteralPath (Join-Path $root "massage") -Recurse -File -Filter "index.html"

$rows = New-Object System.Collections.Generic.List[object]

foreach($f in $files){
    $rel=$f.FullName.Substring($root.Length).TrimStart("\").Replace("\","/")
    $html=Get-Content -LiteralPath $f.FullName -Raw -Encoding UTF8
    $title=Get-Title $html
    $desc=Get-Desc $html

    $type="기타지역페이지"
    $newTitle=$title
    $newDesc=$desc
    $action="유지"

    # massage/other/* 지역 허브는 현재처럼 키워드를 약하게 쓰는 구조라 우선 유지
    if($rel -like "massage/other/*"){
        $type="지역허브"
        $action="유지"
    }
    # 읍/면/동 대표 랜딩: 제목 앞단의 지역명을 추출
    elseif($title -match '^(?<loc>.+?(?:동|읍|면))\s+출장마사지\s*\|\s*(?<parent>.+?)\s+(?:동|읍|면)\s+생활권·지역\s+안내\s*\|\s*몽라운지$'){
        $type="동읍면랜딩"
        $loc=$Matches.loc.Trim()
        $parent=$Matches.parent.Trim()
        $v=Stable-Mod $rel 4

        switch($v){
            0 {
                $newTitle="$loc 생활권 안내 | $parent 인접 지역·이동 정보 | 몽라운지"
                $newDesc="$loc 주변 생활권과 인접 지역을 기준으로 현재 위치와 이동 범위를 확인할 수 있도록 정리했습니다. $parent 지역에서 이용 가능한 범위를 비교할 때 참고하세요."
            }
            1 {
                $newTitle="$parent $loc 지역 안내 | 출장마사지 이용 범위·생활권 정보 | 몽라운지"
                $newDesc="${loc}을 중심으로 주거·상업 생활권과 주변 이동 동선을 정리했습니다. 가까운 지역과 이용 범위를 비교하고 위치 선택에 참고할 수 있습니다."
            }
            2 {
                $newTitle="$loc 출장마사지 | $parent 생활권·인접 지역 안내 | 몽라운지"
                $newDesc="${loc}에서 오가는 주요 생활권과 주변 지역 정보를 한눈에 확인할 수 있도록 구성했습니다. $parent 내 현재 위치와 인접 지역의 이동 범위를 살펴보세요."
            }
            3 {
                $newTitle="$loc 지역 이용 정보 | $parent 생활권·위치 안내 | 몽라운지"
                $newDesc="$loc 생활권에서 출장마사지 이용 전 확인할 수 있는 위치 범위와 인접 지역 정보를 정리했습니다. $parent 주변 이동 동선과 가까운 생활권을 함께 확인해보세요."
            }
        }
        $action="변경후보"
    }
    # 일반 동 랜딩 (화곡동/염창동처럼 부모유형 문구가 다른 경우)
    elseif($title -match '^(?<loc>.+?동)\s+출장마사지\s*\|\s*(?<rest>.+?)\s*\|\s*몽라운지$'){
        $type="동랜딩"
        $loc=$Matches.loc.Trim()
        $rest=$Matches.rest.Trim()
        $v=Stable-Mod $rel 3

        switch($v){
            0 {
                $newTitle="$loc 생활권 안내 | $rest | 몽라운지"
                $newDesc=[regex]::Replace($desc,[regex]::Escape("$loc 출장마사지"),"$loc 지역 이용")
            }
            1 {
                $newTitle="$loc 출장마사지 | $rest | 몽라운지"
                $newDesc=[regex]::Replace($desc,[regex]::Escape("$loc 출장마사지"),"$loc 생활권 이용")
            }
            2 {
                $newTitle="$loc 지역 이용 안내 | $rest | 몽라운지"
                $newDesc=$desc
            }
        }
        $action="변경후보"
    }

    $rows.Add([pscustomobject]@{
        Type=$type
        File=$rel
        Action=$action
        CurrentTitle=$title
        ProposedTitle=$newTitle
        CurrentDescription=$desc
        ProposedDescription=$newDesc
        CurrentTitleKW=([regex]::Matches($title,'출장마사지')).Count
        ProposedTitleKW=([regex]::Matches($newTitle,'출장마사지')).Count
        CurrentDescKW=([regex]::Matches($desc,'출장마사지')).Count
        ProposedDescKW=([regex]::Matches($newDesc,'출장마사지')).Count
    })
}

$rows | Export-Csv -LiteralPath $out -NoTypeInformation -Encoding UTF8

$changed=@($rows|Where-Object{$_.Action -eq "변경후보"})
$hubs=@($rows|Where-Object{$_.Type -eq "지역허브"})

Write-Host ""
Write-Host "=== TITLE / META 분산안 미리보기 ==="
Write-Host "사이트 파일 수정: 0"
Write-Host "전체 massage 페이지:" $rows.Count
Write-Host "지역허브 유지:" $hubs.Count
Write-Host "동·읍·면 변경후보:" $changed.Count
Write-Host ""
Write-Host "변경후보 TITLE 출장마사지 0회:" @($changed|Where-Object{$_.ProposedTitleKW -eq 0}).Count
Write-Host "변경후보 TITLE 출장마사지 1회:" @($changed|Where-Object{$_.ProposedTitleKW -eq 1}).Count
Write-Host "변경후보 META  출장마사지 0회:" @($changed|Where-Object{$_.ProposedDescKW -eq 0}).Count
Write-Host "변경후보 META  출장마사지 1회:" @($changed|Where-Object{$_.ProposedDescKW -eq 1}).Count
Write-Host ""
Write-Host "=== 샘플 12개 ==="
$changed|Select-Object -First 12|ForEach-Object{
    Write-Host "파일:" $_.File
    Write-Host "현재 TITLE:" $_.CurrentTitle
    Write-Host "변경 TITLE:" $_.ProposedTitle
    Write-Host "현재 META :" $_.CurrentDescription
    Write-Host "변경 META :" $_.ProposedDescription
    Write-Host "---"
}
Write-Host ""
Write-Host "미리보기 CSV:" $out
