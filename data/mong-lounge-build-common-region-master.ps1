$ErrorActionPreference = "Stop"

$root = "E:\mong-lounge-site"
$old  = Join-Path $root "data\shop-region-location-final-v2.csv"
$new  = Join-Path $root "data\other-region-structured-final.csv"
$out  = Join-Path $root "data\region-location-master.csv"

if (!(Test-Path -LiteralPath $old)) { throw "기존 832개 파일을 찾을 수 없습니다: $old" }
if (!(Test-Path -LiteralPath $new)) { throw "새 401개 파일을 찾을 수 없습니다: $new" }

$oldRows = @(Import-Csv -LiteralPath $old)
$newRows = @(Import-Csv -LiteralPath $new)

# 공용 지역 마스터의 고정 열 순서
$columns = @(
    "업체명",
    "적용지역",
    "광역지역",
    "상위지역",
    "하위지역",
    "대표역",
    "대표역주소",
    "역노선",
    "위치표시",
    "주소표시",
    "역선정방식",
    "역거리km",
    "계산상가까운역",
    "계산거리km"
)

$combined = New-Object System.Collections.Generic.List[object]

# 1) 기존 832개: 업체명만 비우고 역/주소 정보는 그대로 보존
foreach ($r in $oldRows) {
    $combined.Add([pscustomobject][ordered]@{
        업체명         = ""
        적용지역       = [string]$r.적용지역
        광역지역       = [string]$r.광역지역
        상위지역       = [string]$r.상위지역
        하위지역       = [string]$r.하위지역
        대표역         = [string]$r.대표역
        대표역주소     = [string]$r.대표역주소
        역노선         = [string]$r.역노선
        위치표시       = [string]$r.위치표시
        주소표시       = [string]$r.주소표시
        역선정방식     = [string]$r.역선정방식
        역거리km       = [string]$r.역거리km
        계산상가까운역 = [string]$r.계산상가까운역
        계산거리km     = [string]$r.계산거리km
    })
}

# 2) 새 401개: 지역 구조만 추가, 역/주소는 다음 단계에서 채움
foreach ($r in $newRows) {
    if ([string]$r.판정 -ne "OK") { continue }

    $combined.Add([pscustomobject][ordered]@{
        업체명         = ""
        적용지역       = [string]$r.적용지역
        광역지역       = [string]$r.광역지역
        상위지역       = [string]$r.상위지역
        하위지역       = [string]$r.하위지역
        대표역         = ""
        대표역주소     = ""
        역노선         = ""
        위치표시       = ""
        주소표시       = ""
        역선정방식     = ""
        역거리km       = ""
        계산상가까운역 = ""
        계산거리km     = ""
    })
}

# 3) 지역 중복 제거
# 같은 광역/적용/상위/하위 지역이 중복될 경우 기존 역정보가 있는 행을 우선
$dedup = @(
    $combined |
    Sort-Object `
        광역지역, 적용지역, 상위지역, 하위지역, `
        @{Expression={ if (-not [string]::IsNullOrWhiteSpace([string]$_.대표역)) {0} else {1} }} |
    Group-Object 광역지역, 적용지역, 상위지역, 하위지역 |
    ForEach-Object { $_.Group | Select-Object -First 1 }
)

$dedup |
    Select-Object $columns |
    Export-Csv -LiteralPath $out -NoTypeInformation -Encoding UTF8

$blankShop = @($dedup | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.업체명) }).Count
$withStation = @($dedup | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.대표역) }).Count
$withoutStation = $dedup.Count - $withStation
$expectedRaw = $oldRows.Count + @($newRows | Where-Object 판정 -eq "OK").Count
$removed = $expectedRaw - $dedup.Count

Write-Host ""
Write-Host "=== 공용 지역 마스터 생성 결과 ==="
Write-Host "기존 지역:" $oldRows.Count
Write-Host "새 지역:" (@($newRows | Where-Object 판정 -eq "OK").Count)
Write-Host "합치기 전:" $expectedRaw
Write-Host "중복 제거:" $removed
Write-Host "최종 지역:" $dedup.Count
Write-Host ""
Write-Host "업체명 남은 행:" $blankShop
Write-Host "대표역 있음:" $withStation
Write-Host "대표역 미입력:" $withoutStation
Write-Host ""
Write-Host "저장:" $out
Write-Host ""
Write-Host "=== 새 지역별 개수 ==="
$dedup |
    Where-Object { $_.적용지역 -in @("대전","청주","창원","대구","공주","계룡","전주","광주광역시","익산","김해") } |
    Group-Object 적용지역 |
    Sort-Object Name |
    ForEach-Object { Write-Host "$($_.Name) : $($_.Count)개" }

Write-Host ""
Write-Host "원본 CSV 수정: 0"
Write-Host "사이트 HTML 수정: 0"
