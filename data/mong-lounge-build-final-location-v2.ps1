$ErrorActionPreference = "Stop"

$root = "E:\mong-lounge-site"

$nearestFile = Join-Path $root "data\shop-region-with-nearest-station.csv"
$stationFile = Join-Path $root "data\station-master-national.csv"
$outFile     = Join-Path $root "data\shop-region-location-final-v2.csv"

if (!(Test-Path -LiteralPath $nearestFile)) {
    throw "파일 없음: $nearestFile"
}
if (!(Test-Path -LiteralPath $stationFile)) {
    throw "파일 없음: $stationFile"
}

$ci = [System.Globalization.CultureInfo]::InvariantCulture

function To-Double([string]$v) {
    $d = 0.0
    if ([string]::IsNullOrWhiteSpace($v)) { return $null }
    if ([double]::TryParse($v,[System.Globalization.NumberStyles]::Float,$ci,[ref]$d)) {
        return $d
    }
    return $null
}

function Get-ProvinceText([string]$g) {
    switch ($g) {
        "서울" { return "서울" }
        "경기" { return "경기" }
        "인천" { return "인천" }
        "충남" { return "충청남도" }
        default { return $g }
    }
}

function Get-ParentTokens([string]$parent) {
    if ([string]::IsNullOrWhiteSpace($parent)) { return @() }

    $tokens = @()

    foreach ($p in @($parent -split '-')) {
        if ($p -match '(시|군|구)$') {
            $tokens += $p
        }
        elseif ($p -eq "천안") {
            $tokens += "천안시"
        }
        elseif ($p -eq "아산") {
            $tokens += "아산시"
        }
    }

    return @($tokens | Select-Object -Unique)
}

function Get-SameNameFallback($row,$stations) {

    $stem = ([string]$row.하위지역) -replace '(동|읍|면)$',''
    if ([string]::IsNullOrWhiteSpace($stem)) {
        return $null
    }

    $target = $stem + "역"
    $province = Get-ProvinceText ([string]$row.광역지역)
    $tokens = @(Get-ParentTokens ([string]$row.상위지역))

    $cand = @(
        $stations | Where-Object {
            $_.대표역 -eq $target -and
            ([string]$_.도로명주소) -like "*$province*"
        }
    )

    if ($tokens.Count -gt 0) {
        $cand = @(
            $cand | Where-Object {
                $addr = [string]$_.도로명주소
                $ok = $true

                foreach ($token in $tokens) {
                    if ($addr -notlike "*$token*") {
                        $ok = $false
                        break
                    }
                }

                $ok
            }
        )
    }

    if ($cand.Count -eq 0) {
        return $null
    }

    # 같은 역이 노선별로 중복돼도 지역이 모두 맞으면 한 건 사용
    return ($cand | Sort-Object 대표역,도로명주소,노선 | Select-Object -First 1)
}

$rows = @(Import-Csv -LiteralPath $nearestFile)
$stations = @(Import-Csv -LiteralPath $stationFile)

$result = @()

foreach ($x in $rows) {

    $km = To-Double ([string]$x.역거리km)

    $useStation = $false
    $stationName = ""
    $stationAddress = ""
    $stationLine = ""
    $method = "지역중심"

    # 1순위: 좌표 기반 8km 이내 실제 최근접역
    if (
        -not [string]::IsNullOrWhiteSpace($x.대표역) -and
        $null -ne $km -and
        $km -le 8.0
    ) {
        $useStation = $true
        $stationName = [string]$x.대표역
        $stationAddress = [string]$x.역도로명주소
        $stationLine = [string]$x.역노선
        $method = "좌표최단거리"
    }
    else {
        # 2순위: 좌표가 없거나 역이 비어 있을 때만 동일이름 역 + 지역 검증
        if (
            [string]::IsNullOrWhiteSpace($x.대표역) -or
            $null -eq $km
        ) {
            $fallback = Get-SameNameFallback $x $stations

            if ($null -ne $fallback) {
                $useStation = $true
                $stationName = [string]$fallback.대표역
                $stationAddress = [string]$fallback.도로명주소
                $stationLine = [string]$fallback.노선
                $method = "동일이름지역검증"
            }
        }
    }

    $parentText = ([string]$x.상위지역).Replace("-"," ")

    $result += [pscustomobject]@{
        업체명          = $x.업체명
        적용지역        = $x.적용지역
        광역지역        = $x.광역지역
        상위지역        = $x.상위지역
        하위지역        = $x.하위지역

        대표역          = if ($useStation) { $stationName } else { "" }
        대표역주소      = if ($useStation) { $stationAddress } else { "" }
        역노선          = if ($useStation) { $stationLine } else { "" }

        위치표시        = if ($useStation) { "$stationName 인근" } else { "$($x.하위지역) 중심" }
        주소표시        = if ($useStation) { $stationAddress } else { "$($x.광역지역) $parentText $($x.하위지역)" }

        역선정방식      = $method
        역거리km        = if ($method -eq "좌표최단거리") { $x.역거리km } else { "" }

        계산상가까운역  = $x.대표역
        계산거리km      = $x.역거리km
    }
}

$result | Export-Csv -LiteralPath $outFile -NoTypeInformation -Encoding UTF8

$near = @($result | Where-Object 역선정방식 -eq "좌표최단거리").Count
$fallback = @($result | Where-Object 역선정방식 -eq "동일이름지역검증").Count
$center = @($result | Where-Object 역선정방식 -eq "지역중심").Count

Write-Host ""
Write-Host "=== 최종 지역/대표역 데이터 V2 ==="
Write-Host "저장:" $outFile
Write-Host "전체:" $result.Count
Write-Host "좌표 최근접역 사용:" $near
Write-Host "동일이름 지역검증 역 사용:" $fallback
Write-Host "지역 중심 사용:" $center

Write-Host ""
Write-Host "=== 문제 사례 확인 ==="
$result |
    Where-Object {
        $_.하위지역 -in @("탄현동","군포동","역곡동","신평동","왕십리동","계양동")
    } |
    ForEach-Object {
        Write-Host "$($_.광역지역) | $($_.상위지역) -> $($_.하위지역) | $($_.역선정방식) | $($_.대표역) | $($_.대표역주소)"
    }

Write-Host ""
Write-Host "=== 지역 중심 샘플 ==="
$result |
    Where-Object 역선정방식 -eq "지역중심" |
    Select-Object -First 12 |
    ForEach-Object {
        Write-Host "$($_.광역지역) | $($_.상위지역) -> $($_.하위지역) | $($_.위치표시)"
    }

Write-Host ""
Write-Host "사이트 HTML 수정: 0"
Write-Host "최종 지역/대표역 데이터 생성 완료"
