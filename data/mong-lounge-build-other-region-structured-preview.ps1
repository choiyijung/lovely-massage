$ErrorActionPreference = "Stop"

$root = "E:\mong-lounge-site"
$src  = Join-Path $root "data\other-region-link-audit.csv"
$out  = Join-Path $root "data\other-region-structured-preview.csv"

if (!(Test-Path -LiteralPath $src)) {
    throw "입력 파일을 찾을 수 없습니다: $src"
}

$provinceMap = @{
    "대전"       = "대전"
    "청주"       = "충북"
    "창원"       = "경남"
    "대구"       = "대구"
    "공주"       = "충남"
    "계룡"       = "충남"
    "전주"       = "전북"
    "광주광역시" = "광주광역시"
    "익산"       = "전북"
    "김해"       = "경남"
}

# 구/군 단위가 실제로 존재하는 지역
$hasDistrict = @("대전","청주","창원","대구","전주","광주광역시")

function Get-PlainText([string]$html) {
    if ([string]::IsNullOrWhiteSpace($html)) { return "" }
    $t = [regex]::Replace($html, '(?is)<[^>]+>', ' ')
    $t = [System.Net.WebUtility]::HtmlDecode($t)
    return ([regex]::Replace($t, '\s+', ' ').Trim())
}

function Get-PageParent([string]$path, [string]$region) {
    if (!(Test-Path -LiteralPath $path)) { return $region }

    $html = Get-Content -LiteralPath $path -Raw -Encoding UTF8
    $source = ""

    $h1 = [regex]::Match($html, '(?is)<h1\b[^>]*>(.*?)</h1>')
    if ($h1.Success) {
        $source = Get-PlainText $h1.Groups[1].Value
    }

    if ([string]::IsNullOrWhiteSpace($source)) {
        $title = [regex]::Match($html, '(?is)<title\b[^>]*>(.*?)</title>')
        if ($title.Success) {
            $source = Get-PlainText $title.Groups[1].Value
        }
    }

    # 대구의 달성군/군위군 같은 군 단위도 포함
    $m = [regex]::Match($source, '([가-힣]{1,10}(?:구|군))')
    if ($m.Success) {
        return $m.Groups[1].Value
    }

    return $region
}

$input = @(Import-Csv -LiteralPath $src)
$pageParentCache = @{}
$result = New-Object System.Collections.Generic.List[object]

foreach ($r in $input) {
    $region = [string]$r.지역
    $area   = ([string]$r.링크텍스트).Trim()
    $page   = [string]$r.원본페이지

    if ([string]::IsNullOrWhiteSpace($region) -or
        [string]::IsNullOrWhiteSpace($area)) {
        continue
    }

    if ($area -notmatch '(동|읍|면|구)$') {
        continue
    }

    # 구 링크 자체는 하위 동/읍/면 생성 대상에서 제외
    if ($area -match '구$') {
        continue
    }

    if (!$pageParentCache.ContainsKey($page)) {
        $pageParentCache[$page] = Get-PageParent $page $region
    }

    $parent = [string]$pageParentCache[$page]

    $status = "OK"
    if ($region -in $hasDistrict -and $parent -eq $region) {
        $status = "상위확인필요"
    }

    $result.Add([pscustomobject]@{
        광역지역   = $provinceMap[$region]
        적용지역   = $region
        상위지역   = $parent
        하위지역   = $area
        판정       = $status
        원본페이지 = $page
    })
}

# 동일 지역/상위/하위 중복 제거. OK 판정을 우선.
$dedup = @(
    $result |
    Sort-Object 광역지역, 적용지역, 상위지역, 하위지역, @{Expression={ if ($_.판정 -eq "OK") {0} else {1} }} |
    Group-Object 광역지역, 적용지역, 상위지역, 하위지역 |
    ForEach-Object { $_.Group | Select-Object -First 1 }
)

$dedup | Export-Csv -LiteralPath $out -NoTypeInformation -Encoding UTF8

Write-Host ""
Write-Host "=== 기타지역 구조화 미리보기 ==="
Write-Host "저장:" $out
Write-Host "하위지역 총계:" $dedup.Count
Write-Host "OK:" (@($dedup | Where-Object 판정 -eq "OK").Count)
Write-Host "상위확인필요:" (@($dedup | Where-Object 판정 -eq "상위확인필요").Count)
Write-Host ""

foreach ($region in @("대전","청주","창원","대구","공주","계룡","전주","광주광역시","익산","김해")) {
    $rr = @($dedup | Where-Object 적용지역 -eq $region)
    $ok = @($rr | Where-Object 판정 -eq "OK").Count
    $rv = @($rr | Where-Object 판정 -eq "상위확인필요").Count
    Write-Host "${region} : 총 $($rr.Count)개 / OK ${ok}개 / 확인필요 ${rv}개"
}

Write-Host ""
Write-Host "=== 상위지역별 개수 ==="
$dedup |
    Group-Object 적용지역, 상위지역 |
    Sort-Object Name |
    ForEach-Object {
        $first = $_.Group | Select-Object -First 1
        Write-Host "$($first.적용지역) > $($first.상위지역) : $($_.Count)개"
    }

Write-Host ""
Write-Host "사이트 HTML 수정: 0"
