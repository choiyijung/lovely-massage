$ErrorActionPreference = "Stop"

$root = "E:\mong-lounge-site"
$geoFile = Join-Path $root "data\shop-region-geocoded.csv"
$stationFile = Join-Path $root "data\station-master-national.csv"
$outFile = Join-Path $root "data\unresolved-same-name-station-audit.csv"

if (!(Test-Path -LiteralPath $geoFile)) {
    throw "파일 없음: $geoFile"
}
if (!(Test-Path -LiteralPath $stationFile)) {
    throw "파일 없음: $stationFile"
}

$areas = @(Import-Csv -LiteralPath $geoFile)
$stations = @(Import-Csv -LiteralPath $stationFile)

$bad = @(
    $areas | Where-Object {
        $_.지오코딩상태 -ne "OK" -or
        [string]::IsNullOrWhiteSpace($_.위도) -or
        [string]::IsNullOrWhiteSpace($_.경도)
    }
)

function Get-ProvinceText([string]$g) {
    switch ($g) {
        "서울" { "서울" }
        "경기" { "경기" }
        "인천" { "인천" }
        "충남" { "충청남도" }
        default { $g }
    }
}

function Get-ParentTokens([string]$parent) {
    if ([string]::IsNullOrWhiteSpace($parent)) {
        return @()
    }

    $parts = @($parent -split '-')
    return @(
        $parts |
        Where-Object {
            $_ -match '(시|군|구)$'
        }
    )
}

$result = @()

foreach ($x in $bad) {

    $stem = ([string]$x.하위지역) -replace '(동|읍|면)$',''
    $target = $stem + "역"

    $province = Get-ProvinceText ([string]$x.광역지역)
    $parentTokens = @(Get-ParentTokens ([string]$x.상위지역))

    $sameName = @(
        $stations | Where-Object {
            $_.대표역 -eq $target
        }
    )

    $provinceMatched = @(
        $sameName | Where-Object {
            ([string]$_.도로명주소) -like "*$province*"
        }
    )

    $strictMatched = @(
        $provinceMatched | Where-Object {
            $addr = [string]$_.도로명주소

            if ($parentTokens.Count -eq 0) {
                return $true
            }

            foreach ($token in $parentTokens) {
                if ($addr -notlike "*$token*") {
                    return $false
                }
            }

            return $true
        }
    )

    $status = ""
    $pick = $null

    if ($strictMatched.Count -eq 1) {
        $status = "사용가능"
        $pick = $strictMatched[0]
    }
    elseif ($strictMatched.Count -gt 1) {
        $status = "동일지역중복"
    }
    elseif ($provinceMatched.Count -gt 0) {
        $status = "광역일치_상위지역불일치"
    }
    elseif ($sameName.Count -gt 0) {
        $status = "타지역동명이역"
    }
    else {
        $status = "동일이름역없음"
    }

    $result += [pscustomobject]@{
        광역지역       = $x.광역지역
        상위지역       = $x.상위지역
        하위지역       = $x.하위지역
        예상역명       = $target
        판정           = $status
        사용역         = if ($pick) { $pick.대표역 } else { "" }
        사용역주소     = if ($pick) { $pick.도로명주소 } else { "" }
        같은이름역개수 = $sameName.Count
    }
}

$result |
    Export-Csv -LiteralPath $outFile -NoTypeInformation -Encoding UTF8

$usable = @($result | Where-Object 판정 -eq "사용가능").Count
$wrongArea = @($result | Where-Object 판정 -eq "타지역동명이역").Count
$provinceOnly = @($result | Where-Object 판정 -eq "광역일치_상위지역불일치").Count
$duplicate = @($result | Where-Object 판정 -eq "동일지역중복").Count
$none = @($result | Where-Object 판정 -eq "동일이름역없음").Count

Write-Host ""
Write-Host "=== 동일이름 역 지역검증 결과 ==="
Write-Host "검사 지역:" $result.Count
Write-Host "안전하게 사용 가능:" $usable
Write-Host "타지역 동명이역:" $wrongArea
Write-Host "광역만 일치:" $provinceOnly
Write-Host "동일지역 중복:" $duplicate
Write-Host "같은 이름 역 없음:" $none

Write-Host ""
Write-Host "=== 사용 가능 ==="
$result |
    Where-Object 판정 -eq "사용가능" |
    ForEach-Object {
        Write-Host "$($_.상위지역) -> $($_.하위지역) => $($_.사용역) | $($_.사용역주소)"
    }

Write-Host ""
Write-Host "=== 잘못 연결될 수 있어 제외 ==="
$result |
    Where-Object { $_.판정 -ne "사용가능" } |
    Select-Object -First 30 |
    ForEach-Object {
        Write-Host "$($_.상위지역) -> $($_.하위지역) | 예상 $($_.예상역명) | $($_.판정)"
    }

Write-Host ""
Write-Host "저장:" $outFile
Write-Host "사이트 HTML 수정: 0"

