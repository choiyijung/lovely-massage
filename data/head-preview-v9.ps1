$ErrorActionPreference = "Stop"

$dataDir = "E:\mong-lounge-site\data"
$stamp = Get-Date -Format "yyyyMMdd-HHmmss"

$src = Get-ChildItem -LiteralPath $dataDir -File -Filter "head-rewrite-preview-v8-*.csv" |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1

if(!$src){
    throw "v8 미리보기 CSV를 찾지 못했습니다."
}

$out = Join-Path $dataDir "head-rewrite-preview-v9-$stamp.csv"
$rows = Import-Csv -LiteralPath $src.FullName

$fixed = 0
$noFactFixed = 0

foreach($r in $rows){
    $before = [string]$r.NewDescription
    $x = $before

    if($r.Class -eq "업체상세"){
        # 코스/가격 수치가 없는 페이지에서 남은 불완전 문구 제거
        if($x -match '\.\s{2,}확인되는 가격 정보를 포함해\s*'){
            $x = [regex]::Replace(
                $x,
                '\.\s{2,}확인되는 가격 정보를 포함해\s*',
                '. '
            )
            $noFactFixed++
        }

        # 혹시 공백 수가 달라도 같은 불완전 문구 제거
        $x = [regex]::Replace(
            $x,
            '(?<!\d원)\s+확인되는 가격 정보를 포함해\s+',
            ' '
        )

        # 어색한 "부터의 가격" 표현 자연스럽게 보정
        $x = [regex]::Replace(
            $x,
            '(\d{1,3}(?:,\d{3})*원)부터의 가격대',
            '$1부터 확인되는 가격대'
        )
        $x = [regex]::Replace(
            $x,
            '(\d{1,3}(?:,\d{3})*원)부터의 가격 정보',
            '$1부터 확인되는 가격 정보'
        )
    }

    # 전 페이지 공통 공백 정리
    $x = [regex]::Replace($x, '\s{2,}', ' ').Trim()

    if($x -ne $before){
        $r.NewDescription = $x
        $fixed++
    }
}

$rows | Export-Csv -LiteralPath $out -NoTypeInformation -Encoding UTF8

# 최종 검증
$badDouble = @($rows | Where-Object { $_.NewDescription -match '\s{2,}' })
$badFromFrom = @($rows | Where-Object { $_.NewDescription -match '부터부터|원부터부터' })
$badEmptyPrice = @($rows | Where-Object {
    $_.NewDescription -match '(^|[.!?]\s+)확인되는 가격 정보를 포함해'
})
$badTitleLoc = @()

foreach($r in ($rows | Where-Object Class -eq "업체상세")){
    $parts = @($r.OldTitle -split '\s*\|\s*')
    if($parts.Count -ge 2){
        $loc = $parts[1]
        $loc = $loc -replace '\s+(업체\s*)?상세정보\s*$',''
        $loc = $loc -replace '\s+이용\s*안내\s*$',''
        $loc = $loc.Trim()
        if($loc -and $r.NewTitle -notmatch [regex]::Escape($loc)){
            $badTitleLoc += $r
        }
    }
}

Write-Host ""
Write-Host "=== TITLE / META v9 보정 미리보기 ==="
Write-Host "사이트 파일 수정: 0"
Write-Host "사용 원본:" $src.FullName
Write-Host "DESCRIPTION 수정 후보:" $fixed
Write-Host "수치 없는 불완전 가격문구 제거:" $noFactFixed
Write-Host "업체 TITLE 지역명 누락:" $badTitleLoc.Count
Write-Host "부터부터 남음:" $badFromFrom.Count
Write-Host "이중공백 남음:" $badDouble.Count
Write-Host "빈 가격문구 남음:" $badEmptyPrice.Count
Write-Host ""

Write-Host "=== 보정된 샘플 최대 15개 ==="
$rows | Where-Object {
    $_.Class -eq "업체상세" -and
    $_.NewDescription -notmatch '\s{2,}'
} | Where-Object {
    $_.File -match '나비테라피'
} | Select-Object -First 15 | ForEach-Object {
    Write-Host "파일 :" $_.File
    Write-Host "제목 :" $_.NewTitle
    Write-Host "설명 :" $_.NewDescription
    Write-Host ""
}

Write-Host "미리보기 CSV:" $out
