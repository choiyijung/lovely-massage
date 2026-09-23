$ErrorActionPreference = "Stop"

$root  = "E:\mong-lounge-site"
$input = Join-Path $root "data\shop-region-expanded-preview.csv"
$cache = Join-Path $root "data\area-geocode-cache.csv"
$out   = Join-Path $root "data\shop-region-geocoded.csv"

if (!(Test-Path -LiteralPath $input)) {
    throw "입력 파일을 찾을 수 없습니다: $input"
}

function Get-ProvinceName([string]$g) {
    switch ($g) {
        "서울" { return "서울특별시" }
        "경기" { return "경기도" }
        "인천" { return "인천광역시" }
        "충남" { return "충청남도" }
        default { return $g }
    }
}

function Get-Query($row) {
    $province = Get-ProvinceName ([string]$row.광역지역)
    $parent   = ([string]$row.상위지역).Replace("-", " ").Trim()
    $child    = ([string]$row.하위지역).Trim()

    if ([string]::IsNullOrWhiteSpace($parent)) {
        return "대한민국 $province $child"
    }

    return "대한민국 $province $parent $child"
}

$rows = @(Import-Csv -LiteralPath $input)

# 기존 캐시 불러오기
$cacheRows = @()
$cacheMap = @{}

if (Test-Path -LiteralPath $cache) {
    $cacheRows = @(Import-Csv -LiteralPath $cache)

    foreach ($c in $cacheRows) {
        if (-not [string]::IsNullOrWhiteSpace($c.검색문구)) {
            $cacheMap[$c.검색문구] = $c
        }
    }
}

# 이번 입력에서 필요한 고유 검색어
$queries = @(
    $rows |
    ForEach-Object { Get-Query $_ } |
    Sort-Object -Unique
)

$need = @(
    $queries | Where-Object { -not $cacheMap.ContainsKey($_) }
)

Write-Host ""
Write-Host "=== 지역 좌표 자동수집 시작 ==="
Write-Host "입력 지역:" $rows.Count
Write-Host "고유 검색어:" $queries.Count
Write-Host "기존 캐시:" $cacheMap.Count
Write-Host "이번 신규 조회:" $need.Count
Write-Host ""
Write-Host "※ 1회 조회 후 1.1초 대기합니다."
Write-Host "※ 중간에 종료되어도 캐시 파일에 저장되어 다음 실행 때 이어집니다."
Write-Host ""

$headers = @{
    "User-Agent" = "MongLoungeRegionGeocoder/1.0"
    "Accept-Language" = "ko"
}

$idx = 0

foreach ($q in $need) {

    $idx++
    Write-Host "[$idx/$($need.Count)] $q"

    $encoded = [System.Uri]::EscapeDataString($q)
    $uri = "https://nominatim.openstreetmap.org/search?format=jsonv2&limit=1&countrycodes=kr&addressdetails=1&q=$encoded"

    $item = $null

    try {
        $res = @(Invoke-RestMethod -Uri $uri -Method Get -Headers $headers -TimeoutSec 30)

        if ($res.Count -gt 0) {
            $r = $res[0]

            $item = [pscustomobject]@{
                검색문구       = $q
                상태           = "OK"
                위도           = [string]$r.lat
                경도           = [string]$r.lon
                검색결과주소   = [string]$r.display_name
                OSM종류        = [string]$r.osm_type
                OSM_ID         = [string]$r.osm_id
                오류           = ""
            }
        }
        else {
            $item = [pscustomobject]@{
                검색문구       = $q
                상태           = "NOT_FOUND"
                위도           = ""
                경도           = ""
                검색결과주소   = ""
                OSM종류        = ""
                OSM_ID         = ""
                오류           = ""
            }
        }
    }
    catch {
        $item = [pscustomobject]@{
            검색문구       = $q
            상태           = "ERROR"
            위도           = ""
            경도           = ""
            검색결과주소   = ""
            OSM종류        = ""
            OSM_ID         = ""
            오류           = $_.Exception.Message
        }
    }

    $cacheRows += $item
    $cacheMap[$q] = $item

    # 매 요청마다 저장하여 중간 종료 대비
    $cacheRows |
        Sort-Object 검색문구 -Unique |
        Export-Csv -LiteralPath $cache -NoTypeInformation -Encoding UTF8

    Start-Sleep -Milliseconds 1100
}

# 원본 832행에 좌표 붙이기
$result = foreach ($row in $rows) {

    $q = Get-Query $row
    $c = $cacheMap[$q]

    [pscustomobject]@{
        업체명         = $row.업체명
        적용지역       = $row.적용지역
        광역지역       = $row.광역지역
        상위지역       = $row.상위지역
        하위지역       = $row.하위지역
        검색문구       = $q
        위도           = if ($c) { $c.위도 } else { "" }
        경도           = if ($c) { $c.경도 } else { "" }
        지오코딩상태   = if ($c) { $c.상태 } else { "MISSING" }
        검색결과주소   = if ($c) { $c.검색결과주소 } else { "" }
    }
}

$result |
    Export-Csv -LiteralPath $out -NoTypeInformation -Encoding UTF8

$ok   = @($result | Where-Object { $_.지오코딩상태 -eq "OK" }).Count
$nf   = @($result | Where-Object { $_.지오코딩상태 -eq "NOT_FOUND" }).Count
$err  = @($result | Where-Object { $_.지오코딩상태 -eq "ERROR" }).Count
$miss = @($result | Where-Object { $_.지오코딩상태 -eq "MISSING" }).Count

Write-Host ""
Write-Host "=== 지역 좌표 자동수집 결과 ==="
Write-Host "저장:" $out
Write-Host "전체:" $result.Count
Write-Host "좌표 성공:" $ok
Write-Host "못 찾음:" $nf
Write-Host "오류:" $err
Write-Host "캐시 누락:" $miss

Write-Host ""
Write-Host "=== 강남구 샘플 ==="
$result |
    Where-Object { $_.상위지역 -eq "강남구" } |
    Select-Object -First 10 |
    ForEach-Object {
        Write-Host "$($_.하위지역) | $($_.위도), $($_.경도) | $($_.검색결과주소)"
    }

Write-Host ""
Write-Host "사이트 HTML 수정: 0"
Write-Host "지역 좌표 수집 완료"
