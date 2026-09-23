$ErrorActionPreference = "Stop"

$root = "E:\mong-lounge-site"
$stamp = Get-Date -Format "yyyyMMdd-HHmmss"
$backup = "E:\mong-lounge-footer-merge-v2-backup-$stamp"

$block = @'
<!-- mong-company-info:start -->
<div class="mong-company-info-inline"
     style="display:block!important;margin:22px 0 12px!important;padding:18px 0 0!important;border-top:1px solid rgba(255,255,255,.12)!important;background:transparent!important;color:#cfcfcf!important;font-family:inherit!important;">
  <div style="font-size:12px!important;line-height:1.9!important;word-break:keep-all!important;">
    <div style="margin:0 0 8px!important;font-size:13px!important;font-weight:800!important;color:#f4c94f!important;">사업자 정보</div>
    <div>상호명: 주)에이치아이행복지킴이</div>
    <div>대표자: 최이정</div>
    <div>사업장주소: 서울 양천구 신월로 170 (신월동) 3층 304-24호</div>
    <div>연락처: <a href="tel:01083799746" style="color:#cfcfcf!important;text-decoration:none!important;">01083799746</a></div>
    <div>사업자등록번호: 803-81-00107</div>
    <div style="font-weight:700!important;color:#f1f1f1!important;">통신판매업신고번호: 2021-서울영등포-0841</div>
    <div>대표자 이메일: <a href="mailto:happyloanstar@gmail.com" style="color:#cfcfcf!important;text-decoration:none!important;">happyloanstar@gmail.com</a></div>
  </div>
</div>
<!-- mong-company-info:end -->
'@

if(!(Test-Path -LiteralPath $root)){
    throw "사이트 폴더가 없습니다: $root"
}

New-Item -ItemType Directory -Force -Path $backup | Out-Null

$files = @(
    Get-ChildItem -LiteralPath $root -Recurse -File -Filter "*.html" |
    Where-Object { $_.FullName -notmatch '\\\.git\\|\\node_modules\\|\\backup\\|\\_backup\\' }
)

$modified = 0
$failed = New-Object System.Collections.Generic.List[string]

foreach($f in $files){
    try{
        $html = Get-Content -LiteralPath $f.FullName -Raw -Encoding UTF8
        $before = $html

        # 기존 사업자 정보 블록 제거
        $html = [regex]::Replace(
            $html,
            '(?is)\s*<!-- mong-company-info:start -->.*?<!-- mong-company-info:end -->\s*',
            "`r`n",
            1
        )

        $inserted = $false

        # 기존 MONG LOUNGE copyright 바로 앞에 삽입
        $copyrightPattern = '(?is)(<(?<tag>[a-z0-9]+)\b[^>]*>\s*(?:©|&copy;)\s*2026\s+MONG\s+LOUNGE\s*</\k<tag>>)'
        if([regex]::IsMatch($html,$copyrightPattern)){
            $html = [regex]::Replace(
                $html,
                $copyrightPattern,
                [System.Text.RegularExpressions.MatchEvaluator]{
                    param($m)
                    return $block + "`r`n" + $m.Groups[1].Value
                },
                1
            )
            $inserted = $true
        }

        # copyright 구조가 다르면 footer 닫기 직전에 삽입
        if(!$inserted -and [regex]::IsMatch($html,'(?is)</footer>')){
            $html = [regex]::Replace(
                $html,
                '(?is)</footer>',
                $block + "`r`n</footer>",
                1
            )
            $inserted = $true
        }

        # footer 자체가 없는 경우 body 끝에 검정 영역으로 삽입
        if(!$inserted -and [regex]::IsMatch($html,'(?is)</body>')){
            $fallback = '<div style="background:#111!important;padding:0 20px 120px!important;color:#cfcfcf!important;"><div style="width:min(1180px,100%)!important;margin:0 auto!important;">' + $block + '</div></div>'
            $html = [regex]::Replace(
                $html,
                '(?is)</body>',
                $fallback + "`r`n</body>",
                1
            )
            $inserted = $true
        }

        if($html -ne $before){
            $rel = $f.FullName.Substring($root.Length).TrimStart("\")
            $dest = Join-Path $backup $rel
            New-Item -ItemType Directory -Force -Path (Split-Path $dest -Parent) | Out-Null
            Copy-Item -LiteralPath $f.FullName -Destination $dest -Force

            Set-Content -LiteralPath $f.FullName -Value $html -Encoding UTF8
            $modified++
        }
    }
    catch{
        $failed.Add("$($f.FullName) | $($_.Exception.Message)")
    }
}

# 검증
$found = 0
$duplicate = 0
$telecom = 0
$oldGreen = 0

foreach($f in $files){
    $html = Get-Content -LiteralPath $f.FullName -Raw -Encoding UTF8
    $cnt = ([regex]::Matches($html,'<!-- mong-company-info:start -->')).Count
    if($cnt -eq 1){ $found++ }
    if($cnt -gt 1){ $duplicate++ }
    if($html -match '통신판매업신고번호:\s*2021-서울영등포-0841'){ $telecom++ }
    if($html -match 'id="mong-company-info"'){ $oldGreen++ }
}

Write-Host ""
Write-Host "=== 사업자 정보 푸터 통합 v2 완료 ==="
Write-Host "전체 HTML:" $files.Count
Write-Host "수정 페이지:" $modified
Write-Host "실패:" $failed.Count
Write-Host "사업자 정보 1개 정상:" $found
Write-Host "통신판매업신고번호 표시 페이지:" $telecom
Write-Host "중복 삽입 페이지:" $duplicate
Write-Host "기존 초록 독립블록 남음:" $oldGreen
Write-Host "백업 폴더:" $backup

if($failed.Count -gt 0){
    $report = Join-Path $root "data\footer-merge-v2-fail-$stamp.txt"
    $failed | Set-Content -LiteralPath $report -Encoding UTF8
    Write-Host "실패 보고서:" $report
}
