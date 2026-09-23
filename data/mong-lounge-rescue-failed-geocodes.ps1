$ErrorActionPreference = "Stop"

$root       = "E:\mong-lounge-site"
$masterIn   = Join-Path $root "data\region-location-master-final.csv"
$stationCsv = Join-Path $root "data\station-master-national.csv"
$cacheCsv   = Join-Path $root "data\region-location-geocode-cache-v2.csv"
$masterOut  = Join-Path $root "data\region-location-master-final-v2.csv"

if (!(Test-Path -LiteralPath $masterIn))   { throw "마스터를 찾을 수 없습니다: $masterIn" }
if (!(Test-Path -LiteralPath $stationCsv)) { throw "전국 역 데이터를 찾을 수 없습니다: $stationCsv" }
if (!(Test-Path -LiteralPath $cacheCsv))   { throw "좌표 캐시를 찾을 수 없습니다: $cacheCsv" }

$targetRegions = @("대전","청주","창원","대구","공주","계룡","전주","광주광역시","익산","김해")

function To-Double([string]$v) {
    if ([string]::IsNullOrWhiteSpace($v)) { return $null }
    try { return [double]::Parse($v,[Globalization.CultureInfo]::InvariantCulture) }
    catch { return $null }
}

function Get-DistanceKm([double]$lat1,[double]$lon1,[double]$lat2,[double]$lon2) {
    $R = 6371.0088
    $rad = [Math]::PI / 180.0
    $dLat = ($lat2-$lat1)*$rad
    $dLon = ($lon2-$lon1)*$rad
    $a = [Math]::Sin($dLat/2)*[Math]::Sin($dLat/2) +
         [Math]::Cos($lat1*$rad)*[Math]::Cos($lat2*$rad) *
         [Math]::Sin($dLon/2)*[Math]::Sin($dLon/2)
    return $R * (2*[Math]::Atan2([Math]::Sqrt($a),[Math]::Sqrt(1-$a)))
}

function Get-CityToken([string]$region) {
    switch ($region) {
        "대전"       { "대전광역시" }
        "청주"       { "청주시" }
        "창원"       { "창원시" }
        "대구"       { "대구광역시" }
        "공주"       { "공주시" }
        "계룡"       { "계룡시" }
        "전주"       { "전주시" }
        "광주광역시" { "광주광역시" }
        "익산"       { "익산시" }
        "김해"       { "김해시" }
        default      { $region }
    }
}

function Get-BaseLocation([string]$region,[string]$parent) {
    switch ($region) {
        "대전"       { "대전광역시 $parent" }
        "청주"       { "청주시 $parent" }
        "창원"       { "창원시 $parent" }
        "대구"       { "대구광역시 $parent" }
        "공주"       { "공주시" }
        "계룡"       { "계룡시" }
        "전주"       { "전주시 $parent" }
        "광주광역시" { "광주광역시 $parent" }
        "익산"       { "익산시" }
        "김해"       { "김해시" }
        default      { "$region $parent" }
    }
}

function Get-FallbackQueries([string]$region,[string]$parent,[string]$area) {
    $base = Get-BaseLocation $region $parent
    $q = New-Object System.Collections.Generic.List[string]

    $q.Add("$area $base")

    if ($area -match "동$") {
        $q.Add("$area 행정복지센터 $base")
        $q.Add("$area 주민센터 $base")
    }
    elseif ($area -match "읍$") {
        $q.Add("$area 행정복지센터 $base")
        $q.Add("$area 읍사무소 $base")
    }
    elseif ($area -match "면$") {
        $q.Add("$area 행정복지센터 $base")
        $q.Add("$area 면사무소 $base")
    }

    $q.Add("행정복지센터 $area $base")
    return @($q | Select-Object -Unique)
}

function Select-SafeResult($results,[string]$region,[string]$parent,[string]$area) {
    if ($null -eq $results) { return $null }

    $cityToken = Get-CityToken $region
    $areaStem = $area -replace "(동|읍|면)$",""
    $best = $null
    $bestScore = -1

    foreach ($x in @($results)) {
        $display = [string]$x.display_name
        if ([string]::IsNullOrWhiteSpace($display)) { continue }

        # 다른 도시 동명이 지명을 받지 않도록 해당 도시/광역시가 반드시 있어야 함
        if ($display -notlike "*$cityToken*") { continue }

        $score = 100
        if (-not [string]::IsNullOrWhiteSpace($parent) -and $display -like "*$parent*") { $score += 40 }
        if ($display -like "*$area*") { $score += 60 }
        elseif ($display -like "*$areaStem*") { $score += 30 }
        else { continue }

        if ($score -gt $bestScore) {
            $best = $x
            $bestScore = $score
        }
    }

    return $best
}

function Test-StationRegion($station,[string]$region) {
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
$cacheRows = @(Import-Csv -LiteralPath $cacheCsv)
$cache = @{}
foreach ($c in $cacheRows) {
    if (-not [string]::IsNullOrWhiteSpace([string]$c.키)) {
        $cache[[string]$c.키] = $c
    }
}

$stations = @()
foreach ($s in @(Import-Csv -LiteralPath $stationCsv)) {
    $lat = To-Double ([string]$s.위도)
    $lon = To-Double ([string]$s.경도)
    if ($null -eq $lat -or $null -eq $lon) { continue }
    $stations += [pscustomobject]@{
        대표역=[string]$s.대표역
        노선=[string]$s.노선
        도로명주소=[string]$s.도로명주소
        위도=$lat
        경도=$lon
    }
}

$failedKeys = @(
    $cache.Values |
    Where-Object { [string]$_.상태 -ne "OK" -and $_.적용지역 -in $targetRegions } |
    Select-Object -ExpandProperty 키
)

Write-Host ""
Write-Host "=== 좌표 실패 재보정 시작 ==="
Write-Host "재조회 대상:" $failedKeys.Count
Write-Host "기존 성공 유지:" (@($cache.Values | Where-Object 상태 -eq "OK").Count)
Write-Host ""

$headers = @{
    "User-Agent"="MongLoungeRegionBuilder/1.1"
    "Accept-Language"="ko"
}

$retryOk = 0
$retryFail = 0
$apiCalls = 0
$i = 0

foreach ($key in $failedKeys) {
    $i++
    $c = $cache[$key]
    $region = [string]$c.적용지역
    $parent = [string]$c.상위지역
    $area   = [string]$c.하위지역

    $picked = $null
    $usedQuery = ""

    foreach ($q in @(Get-FallbackQueries $region $parent $area)) {
        try {
            $url = "https://nominatim.openstreetmap.org/search?format=jsonv2&limit=5&countrycodes=kr&q=" + [uri]::EscapeDataString($q)
            $resp = @(Invoke-RestMethod -Uri $url -Headers $headers -Method Get -TimeoutSec 20)
            $apiCalls++
            $candidate = Select-SafeResult $resp $region $parent $area
            if ($candidate) {
                $picked = $candidate
                $usedQuery = $q
                break
            }
        } catch {
        }
        Start-Sleep -Milliseconds 1100
    }

    if ($picked) {
        $c.상태 = "OK"
        $c.위도 = [string]$picked.lat
        $c.경도 = [string]$picked.lon
        $c.표시주소 = [string]$picked.display_name
        $c.조회문구 = $usedQuery
        $retryOk++
    }
    else {
        $retryFail++
    }

    if (($i % 25) -eq 0 -or $i -eq $failedKeys.Count) {
        Write-Host "재보정 진행: $i / $($failedKeys.Count) | 성공 $retryOk"
    }
}

@($cache.Values) |
    Sort-Object 적용지역,상위지역,하위지역 |
    Export-Csv -LiteralPath $cacheCsv -NoTypeInformation -Encoding UTF8

# 새 캐시를 기준으로 새 10개 지역 역 정보를 다시 계산
$stationNear = 0
$sameName = 0
$centerOnly = 0

foreach ($r in $master) {
    if ($r.적용지역 -notin $targetRegions) { continue }

    $key = "$($r.적용지역)|$($r.상위지역)|$($r.하위지역)"
    $c = $cache[$key]
    $lat = if ($c) { To-Double ([string]$c.위도) } else { $null }
    $lon = if ($c) { To-Double ([string]$c.경도) } else { $null }

    $r.대표역 = ""
    $r.대표역주소 = ""
    $r.역노선 = ""
    $r.위치표시 = "$($r.하위지역) 중심"
    $r.주소표시 = ""
    $r.역선정방식 = "지역중심"
    $r.역거리km = ""
    $r.계산상가까운역 = ""
    $r.계산거리km = ""

    if ($c -and [string]$c.상태 -eq "OK" -and $null -ne $lat -and $null -ne $lon) {
        $nearest = $null
        $nearestKm = $null

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
            continue
        }
    }

    # 좌표가 없거나 8km 밖이면 같은 지역의 동일 이름 역만 안전하게 보정
    $base = ([string]$r.하위지역) -replace "(동|읍|면)$",""
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
        $r.계산상가까운역 = [string]$pick.대표역
        $sameName++
    }
    else {
        $centerOnly++
    }
}

$master | Export-Csv -LiteralPath $masterOut -NoTypeInformation -Encoding UTF8

$cacheOk = @($cache.Values | Where-Object 상태 -eq "OK").Count
$cacheFail = @($cache.Values | Where-Object 상태 -ne "OK").Count

Write-Host ""
Write-Host "=== 좌표 실패 재보정 완료 ==="
Write-Host "재조회 대상:" $failedKeys.Count
Write-Host "추가 성공:" $retryOk
Write-Host "여전히 실패:" $retryFail
Write-Host "API 조회:" $apiCalls
Write-Host ""
Write-Host "캐시 최종 성공:" $cacheOk
Write-Host "캐시 최종 실패:" $cacheFail
Write-Host ""
Write-Host "새 10개 지역 역 연결:" $stationNear
Write-Host "동일이름역 보정:" $sameName
Write-Host "지역 중심:" $centerOnly
Write-Host ""
Write-Host "최종 저장:" $masterOut
Write-Host "사이트 HTML 수정: 0"
