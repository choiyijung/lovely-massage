$ErrorActionPreference = "Stop"

$root      = "E:\mong-lounge-site"
$geoFile   = Join-Path $root "data\shop-region-geocoded.csv"
$cacheFile = Join-Path $root "data\area-geocode-cache.csv"

if (!(Test-Path -LiteralPath $geoFile)) {
    throw "파일 없음: $geoFile"
}

$rows = @(Import-Csv -LiteralPath $geoFile)

$bad = @(
    $rows | Where-Object {
        $_.지오코딩상태 -eq "OK" -and (
            [string]::IsNullOrWhiteSpace($_.위도) -or
            [string]::IsNullOrWhiteSpace($_.경도)
        )
    }
)

Write-Host ""
Write-Host "=== 좌표 빈값 보정 시작 ==="
Write-Host "전체 지역:" $rows.Count
Write-Host "OK인데 좌표 빈값:" $bad.Count

if ($bad.Count -eq 0) {
    Write-Host "보정할 항목이 없습니다."
    Write-Host "사이트 HTML 수정: 0"
    exit
}

$headers = @{
    "User-Agent"      = "MongLoungeRegionGeocoder/1.1"
    "Accept-Language" = "ko"
}

$fixed = 0
$failed = 0
$i = 0

foreach ($row in $bad) {
    $i++

    $province = switch ([string]$row.광역지역) {
        "서울" { "서울특별시" }
        "경기" { "경기도" }
        "인천" { "인천광역시" }
        "충남" { "충청남도" }
        default { [string]$row.광역지역 }
    }

    $parent = ([string]$row.상위지역).Replace("-", " ").Trim()
    $child  = ([string]$row.하위지역).Trim()

    $queries = @(
        ([string]$row.검색문구).Trim(),
        ("대한민국 {0} {1} {2}" -f $province, $parent, $child).Trim(),
        ("{0} {1} {2}" -f $child, $parent, $province).Trim()
    ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique

    $found = $null
    $usedQuery = ""

    Write-Host "[$i/$($bad.Count)] $($row.광역지역) | $($row.상위지역) -> $child"

    foreach ($q in $queries) {
        try {
            $encoded = [Uri]::EscapeDataString($q)
            $uri = "https://nominatim.openstreetmap.org/search?format=jsonv2&limit=1&countrycodes=kr&addressdetails=1&q=$encoded"

            # Invoke-RestMethod 대신 원문 JSON을 받아 직접 변환
            $resp = Invoke-WebRequest -Uri $uri -UseBasicParsing -Headers $headers -TimeoutSec 30
            $json = $resp.Content | ConvertFrom-Json

            if ($json -and $json.Count -gt 0) {
                $cand = @($json)[0]

                if (-not [string]::IsNullOrWhiteSpace([string]$cand.lat) -and
                    -not [string]::IsNullOrWhiteSpace([string]$cand.lon)) {
                    $found = $cand
                    $usedQuery = $q
                    break
                }
            }
        }
        catch {
            # 다음 검색문구로 계속
        }

        Start-Sleep -Milliseconds 1100
    }

    if ($found) {
        $row.위도 = [string]$found.lat
        $row.경도 = [string]$found.lon
        $row.지오코딩상태 = "OK"
        $row.검색결과주소 = [string]$found.display_name
        $fixed++

        Write-Host "  보정: $($row.위도), $($row.경도)"
        Write-Host "  주소: $($row.검색결과주소)"
    }
    else {
        $row.지오코딩상태 = "NOT_FOUND"
        $failed++
        Write-Host "  결과 없음"
    }

    Start-Sleep -Milliseconds 1100
}

# 지역 좌표 파일 갱신
$rows | Export-Csv -LiteralPath $geoFile -NoTypeInformation -Encoding UTF8

# 기존 캐시도 같은 검색문구 기준으로 갱신
if (Test-Path -LiteralPath $cacheFile) {
    $cacheRows = @(Import-Csv -LiteralPath $cacheFile)

    foreach ($c in $cacheRows) {
        $match = $rows | Where-Object { $_.검색문구 -eq $c.검색문구 } | Select-Object -First 1

        if ($match) {
            $c.상태 = $match.지오코딩상태
            $c.위도 = $match.위도
            $c.경도 = $match.경도
            $c.검색결과주소 = $match.검색결과주소
        }
    }

    $cacheRows | Export-Csv -LiteralPath $cacheFile -NoTypeInformation -Encoding UTF8
}

$remaining = @(
    $rows | Where-Object {
        $_.지오코딩상태 -eq "OK" -and (
            [string]::IsNullOrWhiteSpace($_.위도) -or
            [string]::IsNullOrWhiteSpace($_.경도)
        )
    }
).Count

Write-Host ""
Write-Host "=== 좌표 빈값 보정 결과 ==="
Write-Host "보정 성공:" $fixed
Write-Host "못 찾음:" $failed
Write-Host "OK인데 좌표 빈값 남음:" $remaining

Write-Host ""
Write-Host "=== 확인 대상 ==="
$rows |
    Where-Object { $_.하위지역 -in @("탄현동","군포동","역곡동") } |
    ForEach-Object {
        Write-Host "$($_.광역지역) | $($_.상위지역) -> $($_.하위지역) | $($_.위도), $($_.경도) | $($_.지오코딩상태) | $($_.검색결과주소)"
    }

Write-Host ""
Write-Host "사이트 HTML 수정: 0"
Write-Host "좌표 빈값 보정 완료"
