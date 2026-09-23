$ErrorActionPreference = "Stop"

$root = "E:\mong-lounge-site"
$src  = Join-Path $root "data\other-region-structured-preview.csv"
$out  = Join-Path $root "data\other-region-structured-final.csv"

if (!(Test-Path -LiteralPath $src)) {
    throw "입력 파일을 찾을 수 없습니다: $src"
}

$map = [ordered]@{
    "\nam-gu\"       = "남구"
    "\dalseo-gu\"    = "달서구"
    "\dong-gu\"      = "동구"
    "\buk-gu\"       = "북구"
    "\seo-gu\"       = "서구"
    "\suseong-gu\"   = "수성구"
    "\jung-gu\"      = "중구"
    "\gunwi-gun\"    = "군위군"
    "\dalseong-gun\" = "달성군"
}

$rows = @(Import-Csv -LiteralPath $src)

foreach ($r in $rows) {
    if ($r.적용지역 -ne "대구") { continue }

    $p = [string]$r.원본페이지
    $parent = $null

    foreach ($key in $map.Keys) {
        if ($p -like "*$key*") {
            $parent = $map[$key]
            break
        }
    }

    if ($parent) {
        $r.상위지역 = $parent
        $r.판정 = "OK"
    }
    else {
        $r.판정 = "상위확인필요"
    }
}

$rows | Export-Csv -LiteralPath $out -NoTypeInformation -Encoding UTF8

$daegu = @($rows | Where-Object 적용지역 -eq "대구")
$unresolved = @($daegu | Where-Object 판정 -ne "OK")

Write-Host ""
Write-Host "=== 대구 상위지역 보정 결과 ==="
Write-Host "저장:" $out
Write-Host "대구 총:" $daegu.Count
Write-Host "OK:" (@($daegu | Where-Object 판정 -eq "OK").Count)
Write-Host "확인필요:" $unresolved.Count
Write-Host ""

$daegu |
    Group-Object 상위지역 |
    Sort-Object Name |
    ForEach-Object {
        Write-Host "$($_.Name) : $($_.Count)개"
    }

if ($unresolved.Count -gt 0) {
    Write-Host ""
    Write-Host "=== 확인필요 목록 ==="
    $unresolved | Select-Object 하위지역,원본페이지 | Format-Table -AutoSize
}

Write-Host ""
Write-Host "전체 새 지역:" $rows.Count
Write-Host "전체 확인필요:" (@($rows | Where-Object 판정 -ne "OK").Count)
Write-Host "사이트 HTML 수정: 0"
