$ErrorActionPreference = "Stop"

$dataDir = "E:\mong-lounge-site\data"

$csv = Get-ChildItem -LiteralPath $dataDir -File -Filter "head-rewrite-preview-v8-*.csv" |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1

if(!$csv){
    throw "v8 미리보기 CSV를 찾지 못했습니다."
}

$rows = Import-Csv -LiteralPath $csv.FullName

$bad = foreach($r in $rows){
    $reasons = New-Object System.Collections.Generic.List[string]

    if($r.NewTitle -match '부터부터'){ $reasons.Add("TITLE_부터부터") }
    if($r.NewTitle -match '원부터부터'){ $reasons.Add("TITLE_원부터부터") }
    if($r.NewDescription -match '부터부터'){ $reasons.Add("META_부터부터") }
    if($r.NewDescription -match '원부터부터'){ $reasons.Add("META_원부터부터") }
    if($r.NewDescription -match '\s{2,}'){ $reasons.Add("META_이중공백") }

    if($reasons.Count -gt 0){
        [pscustomobject]@{
            Reason = ($reasons -join ",")
            Class = $r.Class
            File = $r.File
            NewTitle = $r.NewTitle
            NewDescription = $r.NewDescription
        }
    }
}

Write-Host ""
Write-Host "=== v8 의심 69건 원인 확인 ==="
Write-Host "사용 CSV:" $csv.FullName
Write-Host "의심 건수:" @($bad).Count
Write-Host ""

$bad | Group-Object Reason | Sort-Object Count -Descending | ForEach-Object {
    Write-Host "$($_.Name): $($_.Count)"
}

Write-Host ""
Write-Host "=== 의심 샘플 최대 20개 ==="
$bad | Select-Object -First 20 | ForEach-Object {
    Write-Host "원인 :" $_.Reason
    Write-Host "파일 :" $_.File
    Write-Host "제목 :" $_.NewTitle
    Write-Host "설명 :" $_.NewDescription
    Write-Host ""
}

Write-Host "사이트 파일 수정: 0"
