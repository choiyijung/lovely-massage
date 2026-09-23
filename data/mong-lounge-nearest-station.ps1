$ErrorActionPreference = "Stop"

$root = "E:\mong-lounge-site"

$areasFile    = Join-Path $root "data\shop-region-geocoded.csv"
$stationsFile = Join-Path $root "data\station-master-national.csv"
$outFile      = Join-Path $root "data\shop-region-with-nearest-station.csv"

if (!(Test-Path -LiteralPath $areasFile)) {
    throw "지역 좌표 파일을 찾을 수 없습니다: $areasFile"
}

if (!(Test-Path -LiteralPath $stationsFile)) {
    throw "전국 역 파일을 찾을 수 없습니다: $stationsFile"
}

$culture = [System.Globalization.CultureInfo]::InvariantCulture

function To-Double([string]$v) {
    if ([string]::IsNullOrWhiteSpace($v)) { return $null }
    $d = 0.0
    if ([double]::TryParse($v, [System.Globalization.NumberStyles]::Float, $culture, [ref]$d)) {
        return $d
    }
    return $null
}

function Deg2Rad([double]$deg) {
    return $deg * [Math]::PI / 180.0
}

function Get-DistanceKm(
    [double]$lat1,
    [double]$lon1,
    [double]$lat2,
    [double]$lon2
) {
    $R = 6371.0088

    $dLat = Deg2Rad ($lat2 - $lat1)
    $dLon = Deg2Rad ($lon2 - $lon1)

    $a = [Math]::Sin($dLat / 2) * [Math]::Sin($dLat / 2) +
         [Math]::Cos((Deg2Rad $lat1)) * [Math]::Cos((Deg2Rad $lat2)) *
         [Math]::Sin($dLon / 2) * [Math]::Sin($dLon / 2)

    $c = 2 * [Math]::Atan2([Math]::Sqrt($a), [Math]::Sqrt(1 - $a))

    return $R * $c
}

$areas = @(Import-Csv -LiteralPath $areasFile)
$stationRaw = @(Import-Csv -LiteralPath $stationsFile)

$stations = @()

foreach ($s in $stationRaw) {
    $lat = To-Double ([string]$s.위도)
    $lon = To-Double ([string]$s.경도)

    if ($null -eq $lat -or $null -eq $lon) {
        continue
    }

    $stations += [pscustomobject]@{
        운영기관 = [string]$s.운영기관
        노선 = [string]$s.노선
        대표역 = [string]$s.대표역
        원본역사명 = [string]$s.원본역사명
        도로명주소 = [string]$s.도로명주소
        위도 = [double]$lat
        경도 = [double]$lon
    }
}

Write-Host ""
Write-Host "=== 주변 대표역 계산 시작 ==="
Write-Host "지역 행:" $areas.Count
Write-Host "사용 가능 역:" $stations.Count
Write-Host ""

$result = @()
$i = 0

foreach ($a in $areas) {

    $i++

    if (($i % 50) -eq 0 -or $i -eq 1 -or $i -eq $areas.Count) {
        Write-Host "진행: $i / $($areas.Count)"
    }

    $alat = To-Double ([string]$a.위도)
    $alon = To-Double ([string]$a.경도)

    $best = $null
    $bestKm = [double]::PositiveInfinity

    if ($a.지오코딩상태 -eq "OK" -and $null -ne $alat -and $null -ne $alon) {

        foreach ($s in $stations) {

            # 매우 먼 역 계산을 줄이기 위한 간단한 좌표 사전 필터
            if ([Math]::Abs($s.위도 - $alat) -gt 1.5) { continue }
            if ([Math]::Abs($s.경도 - $alon) -gt 1.8) { continue }

            $km = Get-DistanceKm $alat $alon $s.위도 $s.경도

            if ($km -lt $bestKm) {
                $bestKm = $km
                $best = $s
            }
        }
    }

    $flag = ""

    if ($null -eq $best) {
        $flag = "역미배정"
    }
    elseif ($bestKm -gt 15) {
        $flag = "15km초과"
    }
    elseif ($bestKm -gt 8) {
        $flag = "8km초과"
    }

    $result += [pscustomobject]@{
        업체명           = $a.업체명
        적용지역         = $a.적용지역
        광역지역         = $a.광역지역
        상위지역         = $a.상위지역
        하위지역         = $a.하위지역
        지역위도         = $a.위도
        지역경도         = $a.경도
        지오코딩상태     = $a.지오코딩상태
        대표역           = if ($best) { $best.대표역 } else { "" }
        역노선           = if ($best) { $best.노선 } else { "" }
        역운영기관       = if ($best) { $best.운영기관 } else { "" }
        역도로명주소     = if ($best) { $best.도로명주소 } else { "" }
        역거리km         = if ($best) { [Math]::Round($bestKm, 2).ToString("0.00", $culture) } else { "" }
        검토필요         = $flag
        검색결과주소     = $a.검색결과주소
    }
}

$result |
    Export-Csv -LiteralPath $outFile -NoTypeInformation -Encoding UTF8

$assigned = @($result | Where-Object { -not [string]::IsNullOrWhiteSpace($_.대표역) }).Count
$unassigned = @($result | Where-Object { $_.검토필요 -eq "역미배정" }).Count
$over8 = @($result | Where-Object { $_.검토필요 -eq "8km초과" }).Count
$over15 = @($result | Where-Object { $_.검토필요 -eq "15km초과" }).Count

Write-Host ""
Write-Host "=== 주변 대표역 계산 결과 ==="
Write-Host "저장:" $outFile
Write-Host "전체:" $result.Count
Write-Host "대표역 배정:" $assigned
Write-Host "역 미배정:" $unassigned
Write-Host "8km 초과:" $over8
Write-Host "15km 초과:" $over15

Write-Host ""
Write-Host "=== 강남구 샘플 ==="
$result |
    Where-Object { $_.상위지역 -eq "강남구" } |
    Select-Object -First 15 |
    ForEach-Object {
        Write-Host "$($_.하위지역) -> $($_.대표역) | $($_.역거리km)km | $($_.역도로명주소)"
    }

Write-Host ""
Write-Host "=== 천안 샘플 ==="
$result |
    Where-Object { $_.상위지역 -eq "천안" } |
    Select-Object -First 8 |
    ForEach-Object {
        Write-Host "$($_.하위지역) -> $($_.대표역) | $($_.역거리km)km | $($_.역도로명주소)"
    }

Write-Host ""
Write-Host "=== 아산 샘플 ==="
$result |
    Where-Object { $_.상위지역 -eq "아산" } |
    Select-Object -First 8 |
    ForEach-Object {
        Write-Host "$($_.하위지역) -> $($_.대표역) | $($_.역거리km)km | $($_.역도로명주소)"
    }

Write-Host ""
Write-Host "=== 거리 긴 상위 10개 ==="
$result |
    Where-Object { $_.역거리km } |
    Sort-Object { [double]::Parse($_.역거리km, $culture) } -Descending |
    Select-Object -First 10 |
    ForEach-Object {
        Write-Host "$($_.광역지역) | $($_.상위지역) -> $($_.하위지역) | $($_.대표역) | $($_.역거리km)km | $($_.검토필요)"
    }

Write-Host ""
Write-Host "사이트 HTML 수정: 0"
Write-Host "주변 대표역 계산 완료"
