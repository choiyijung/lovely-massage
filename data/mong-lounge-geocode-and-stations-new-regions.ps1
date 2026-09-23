$ErrorActionPreference = "Stop"

$root       = "E:\mong-lounge-site"
$masterIn   = Join-Path $root "data\region-location-master.csv"
$stationCsv = Join-Path $root "data\station-master-national.csv"
$cacheCsv   = Join-Path $root "data\region-location-geocode-cache-v2.csv"
$masterOut  = Join-Path $root "data\region-location-master-final.csv"

if (!(Test-Path -LiteralPath $masterIn))   { throw "공용 지역 마스터를 찾을 수 없습니다: $masterIn" }
if (!(Test-Path -LiteralPath $stationCsv)) { throw "전국 역 데이터를 찾을 수 없습니다: $stationCsv" }

$targetRegions = @(
    "대전","청주","창원","대구","공주",
    "계룡","전주","광주광역시","익산","김해"
)

function To-Double([string]$v) {
    if ([string]::IsNullOrWhiteSpace($v)) { return $null }
    try {
        return [double]::Parse($v, [Globalization.CultureInfo]::InvariantCulture)
    } catch {
        return $null
    }
}

function Get-DistanceKm([double]$lat1,[double]$lon1,[double]$lat2,[double]$lon2) {
    $R = 6371.0088
    $rad = [Math]::PI / 180.0
    $dLat = ($lat2 - $lat1) * $rad
    $dLon = ($lon2 - $lon1) * $rad
    $a = [Math]::Sin($dLat/2) * [Math]::Sin($dLat/2) +
         [Math]::Cos($lat1*$rad) * [Math]::Cos($lat2*$rad) *
         [Math]::Sin($dLon/2) * [Math]::Sin($dLon/2)
    $c = 2 * [Math]::Atan2([Math]::Sqrt($a), [Math]::Sqrt(1-$a))
    return $R * $c
}

function Get-Queries($r) {
    $region = [string]$r.적용지역
    $parent = [string]$r.상위지역
    $area   = [string]$r.하위지역

    switch ($region) {
        "대전"       { return @("대전광역시 $parent $area 대한민국") }
        "청주"       { return @("충청북도 청주시 $parent $area 대한민국") }
        "창원"       { return @("경상남도 창원시 $parent $area 대한민국") }
        "대구"       { return @("대구광역시 $parent $area 대한민국") }
        "공주"       { return @("충청남도 공주시 $area 대한민국") }
        "계룡"       { return @("충청남도 계룡시 $area 대한민국") }
        "전주"       { return @("전북특별자치도 전주시 $parent $area 대한민국","전라북도 전주시 $parent $area 대한민국") }
        "광주광역시" { return @("광주광역시 $parent $area 대한민국") }
        "익산"       { return @("전북특별자치도 익산시 $area 대한민국","전라북도 익산시 $area 대한민국") }
        "김해"       { return @("경상남도 김해시 $area 대한민국") }
        default      { return @("$region $parent $area 대한민국") }
    }
}

function Test-StationRegion($station, [string]$region) {
    $addr = [string]$station.도로명주소
    switch ($region) {
        "대전"       { return $addr -match "대전광역시" }
        "청주"       { return ($addr -match "충청북도" -and $addr -match "청주") }
        "창원"       { return ($addr -match "경상남도" -and $addr -match "창원") }
        "대구"       { return $addr -match "대구광역시" }
        "공주"       { return ($addr -match "충청남도" -and $addr -match "공주") }
        "계룡"       { return ($addr -match "충청남도" -and $addr -match "계룡") }
        "전주"       { return (($addr -match "전북특별자치도" -or $addr -match "전라북도") -and $addr -match "전주") }
        "광주광역시" { return $addr -match "광주광역시" }
        "익산"       { return (($addr -match "전북특별자치도" -or $addr -match "전라북도") -and $addr -match "익산") }
        "김해"       { return ($addr -match "경상남도" -and $addr -match "김해") }
        default      { return $true }
    }
}

$master = @(Import-Csv -LiteralPath $masterIn)
$stationsRaw = @(Import-Csv -LiteralPath $stationCsv)

$stations = @()
foreach ($s in $stationsRaw) {
    $lat = To-Double ([string]$s.위도)
    $lon = To-Double ([string]$s.경도)
    if ($null -eq $lat -or $null -eq $lon) { continue }

    $stations += [pscustomobject]@{
        대표역     = [string]$s.대표역
        노선       = [string]$s.노선
        도로명주소 = [string]$s.도로명주소
        위도       = $lat
        경도       = $lon
    }
}

# 기존 캐시 읽기
$cache = @{}
if (Test-Path -LiteralPath $cacheCsv) {
    foreach ($c in @(Import-Csv -LiteralPath $cacheCsv)) {
        $key = [string]$c.키
        if (-not [string]::IsNullOrWhiteSpace($key)) {
            $cache[$key] = $c
        }
    }
}

$targets = @(
    $master | Where-Object {
        $_.적용지역 -in $targetRegions
    }
)

Write-Host ""
Write-Host "=== 새 10개 지역 좌표/역 연결 시작 ==="
Write-Host "대상 지역 행:" $targets.Count
Write-Host "전국 역 데이터:" $stations.Count
Write-Host "기존 캐시:" $cache.Count
Write-Host ""
Write-Host "※ 처음 실행은 좌표 조회 때문에 몇 분 걸릴 수 있습니다."
Write-Host ""

$geocodeOk = 0
$geocodeFail = 0
$stationNear = 0
$stationSame = 0
$centerOnly = 0
$cacheHit = 0
$apiCalls = 0
$processed = 0

$headers = @{
    "User-Agent" = "MongLoungeRegionBuilder/1.0 (local static-site data preparation)"
    "Accept-Language" = "ko"
}

foreach ($r in $targets) {
    $processed++
    $key = "$($r.적용지역)|$($r.상위지역)|$($r.하위지역)"

    $geo = $null

    if ($cache.ContainsKey($key)) {
        $geo = $cache[$key]
        $cacheHit++
    }
    else {
        $queries = @(Get-Queries $r)
        $found = $null
        $usedQuery = ""

        foreach ($q in $queries) {
            $usedQuery = $q
            try {
                $url = "https://nominatim.openstreetmap.org/search?format=jsonv2&limit=1&countrycodes=kr&q=" + [uri]::EscapeDataString($q)
                $resp = @(Invoke-RestMethod -Uri $url -Headers $headers -Method Get -TimeoutSec 30)
                $apiCalls++

                if ($resp.Count -gt 0 -and
                    -not [string]::IsNullOrWhiteSpace([string]$resp[0].lat) -and
                    -not [string]::IsNullOrWhiteSpace([string]$resp[0].lon)) {
                    $found = $resp[0]
                    break
                }
            }
            catch {
                # 다음 쿼리 또는 중심 처리
            }

            Start-Sleep -Milliseconds 1100
        }

        if ($found) {
            $geo = [pscustomobject]@{
                키       = $key
                적용지역 = [string]$r.적용지역
                상위지역 = [string]$r.상위지역
                하위지역 = [string]$r.하위지역
                상태     = "OK"
                위도     = [string]$found.lat
                경도     = [string]$found.lon
                표시주소 = [string]$found.display_name
                조회문구 = $usedQuery
            }
        }
        else {
            $geo = [pscustomobject]@{
                키       = $key
                적용지역 = [string]$r.적용지역
                상위지역 = [string]$r.상위지역
                하위지역 = [string]$r.하위지역
                상태     = "NOT_FOUND"
                위도     = ""
                경도     = ""
                표시주소 = ""
                조회문구 = $usedQuery
            }
        }

        $cache[$key] = $geo

        # Nominatim 예의상 요청 사이 대기
        Start-Sleep -Milliseconds 1100
    }

    $lat = To-Double ([string]$geo.위도)
    $lon = To-Double ([string]$geo.경도)

    $nearest = $null
    $nearestKm = $null

    if ([string]$geo.상태 -eq "OK" -and $null -ne $lat -and $null -ne $lon) {
        $geocodeOk++

        foreach ($s in $stations) {
            $km = Get-DistanceKm $lat $lon $s.위도 $s.경도
            if ($null -eq $nearestKm -or $km -lt $nearestKm) {
                $nearestKm = $km
                $nearest = $s
            }
        }

        if ($nearest) {
            $r.계산상가까운역 = [string]$nearest.대표역
            $r.계산거리km = $nearestKm.ToString("0.00",[Globalization.CultureInfo]::InvariantCulture)
        }

        if ($nearest -and $nearestKm -le 8.0) {
            $r.대표역 = [string]$nearest.대표역
            $r.대표역주소 = [string]$nearest.도로명주소
            $r.역노선 = [string]$nearest.노선
            $r.위치표시 = "$($nearest.대표역) 인근"
            $r.주소표시 = [string]$nearest.도로명주소
            $r.역선정방식 = "좌표최단거리"
            $r.역거리km = $nearestKm.ToString("0.00",[Globalization.CultureInfo]::InvariantCulture)
            $stationNear++
        }
        else {
            # 8km 초과는 먼 역을 억지로 표시하지 않음
            $r.대표역 = ""
            $r.대표역주소 = ""
            $r.역노선 = ""
            $r.위치표시 = "$($r.하위지역) 중심"
            $r.주소표시 = ""
            $r.역선정방식 = "지역중심"
            $r.역거리km = ""
            $centerOnly++
        }
    }
    else {
        $geocodeFail++

        # 좌표를 못 찾았을 때만 동일 이름 역을 지역 검증 후 보정
        $base = ([string]$r.하위지역) -replace '(동|읍|면)$',''
        $stationName = "$base`역"
        $same = @(
            $stations | Where-Object {
                $_.대표역 -eq $stationName -and (Test-StationRegion $_ ([string]$r.적용지역))
            }
        )

        if ($same.Count -gt 0) {
            $pick = $same | Select-Object -First 1
            $r.대표역 = [string]$pick.대표역
            $r.대표역주소 = [string]$pick.도로명주소
            $r.역노선 = [string]$pick.노선
            $r.위치표시 = "$($pick.대표역) 인근"
            $r.주소표시 = [string]$pick.도로명주소
            $r.역선정방식 = "동일이름역보정"
            $r.역거리km = ""
            $r.계산상가까운역 = [string]$pick.대표역
            $r.계산거리km = ""
            $stationSame++
        }
        else {
            $r.대표역 = ""
            $r.대표역주소 = ""
            $r.역노선 = ""
            $r.위치표시 = "$($r.하위지역) 중심"
            $r.주소표시 = ""
            $r.역선정방식 = "지역중심"
            $r.역거리km = ""
            $r.계산상가까운역 = ""
            $r.계산거리km = ""
            $centerOnly++
        }
    }

    if (($processed % 25) -eq 0 -or $processed -eq $targets.Count) {
        Write-Host "진행: $processed / $($targets.Count)"
    }
}

# 캐시 저장
@($cache.Values) |
    Sort-Object 적용지역,상위지역,하위지역 |
    Export-Csv -LiteralPath $cacheCsv -NoTypeInformation -Encoding UTF8

# 완성 마스터 저장
$master |
    Export-Csv -LiteralPath $masterOut -NoTypeInformation -Encoding UTF8

$allStation = @($master | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.대표역) }).Count
$allCenter  = @($master | Where-Object { [string]$_.역선정방식 -eq "지역중심" }).Count

Write-Host ""
Write-Host "=== 새 지역 좌표/역 연결 완료 ==="
Write-Host "처리 대상:" $targets.Count
Write-Host "좌표 성공:" $geocodeOk
Write-Host "좌표 실패:" $geocodeFail
Write-Host "캐시 사용:" $cacheHit
Write-Host "새 API 조회:" $apiCalls
Write-Host ""
Write-Host "8km 이내 역 연결:" $stationNear
Write-Host "동일이름역 보정:" $stationSame
Write-Host "지역 중심 처리:" $centerOnly
Write-Host ""
Write-Host "최종 마스터 전체:" $master.Count
Write-Host "전체 대표역 있음:" $allStation
Write-Host "전체 지역중심:" $allCenter
Write-Host ""
Write-Host "저장:" $masterOut
Write-Host "캐시:" $cacheCsv
Write-Host ""
Write-Host "원본 CSV 수정: 0"
Write-Host "사이트 HTML 수정: 0"
