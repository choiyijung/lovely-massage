$ErrorActionPreference = "Stop"

$root = "E:\mong-lounge-site"
$stamp = Get-Date -Format "yyyyMMdd-HHmmss"
$backup = "E:\mong-lounge-business-footer-visible-backup-$stamp"

$block = @'
<!-- mong-company-info:start -->
<div id="mong-company-info"
     style="display:block!important;visibility:visible!important;opacity:1!important;position:relative!important;z-index:20!important;clear:both!important;width:100%!important;box-sizing:border-box!important;margin:48px 0 0!important;padding:28px 20px 30px!important;background:#082f2a!important;color:#f6f2e8!important;border-top:1px solid rgba(255,255,255,.16)!important;font-family:inherit!important;">
  <div style="display:block!important;width:min(1180px,100%)!important;margin:0 auto!important;">
    <p style="display:block!important;margin:0 0 12px!important;font-size:15px!important;font-weight:800!important;line-height:1.5!important;color:#f4d889!important;">
      사업자 정보
    </p>
    <div style="display:block!important;font-size:12px!important;line-height:1.9!important;color:rgba(255,255,255,.86)!important;word-break:keep-all!important;">
      <div style="display:block!important;">상호명: 주)에이치아이행복지킴이</div>
      <div style="display:block!important;">대표자: 최이정</div>
      <div style="display:block!important;">사업장주소: 서울 양천구 신월로 170 (신월동) 3층 304-24호</div>
      <div style="display:block!important;">연락처: <a href="tel:01083799746" style="color:#f6f2e8!important;text-decoration:none!important;">01083799746</a></div>
      <div style="display:block!important;">사업자등록번호: 803-81-00107</div>
      <div style="display:block!important;">통신판매업신고번호: 2021-서울영등포-0841</div>
      <div style="display:block!important;">대표자 이메일: <a href="mailto:happyloanstar@gmail.com" style="color:#f6f2e8!important;text-decoration:none!important;">happyloanstar@gmail.com</a></div>
    </div>
  </div>
</div>
<!-- mong-company-info:end -->
'@

New-Item -ItemType Directory -Force -Path $backup | Out-Null

$files = @(
    Get-ChildItem -LiteralPath $root -Recurse -File -Filter "*.html" |
    Where-Object { $_.FullName -notmatch '\\\.git\\|\\node_modules\\|\\backup\\|\\_backup\\' }
)

$modified=0
$missingMarker=0
$failed=New-Object System.Collections.Generic.List[string]

foreach($f in $files){
    try{
        $html=Get-Content -LiteralPath $f.FullName -Raw -Encoding UTF8
        $before=$html

        if($html -match '(?is)<!-- mong-company-info:start -->.*?<!-- mong-company-info:end -->'){
            $rel=$f.FullName.Substring($root.Length).TrimStart("\")
            $dest=Join-Path $backup $rel
            New-Item -ItemType Directory -Force -Path (Split-Path $dest -Parent) | Out-Null
            Copy-Item -LiteralPath $f.FullName -Destination $dest -Force

            $html=[regex]::Replace(
                $html,
                '(?is)<!-- mong-company-info:start -->.*?<!-- mong-company-info:end -->',
                [System.Text.RegularExpressions.MatchEvaluator]{ param($m) $block },
                1
            )

            Set-Content -LiteralPath $f.FullName -Value $html -Encoding UTF8
            $modified++
        }
        else{
            $missingMarker++
        }
    }
    catch{
        $failed.Add("$($f.FullName) | $($_.Exception.Message)")
    }
}

# 검증
$visible=0
$duplicates=0
foreach($f in $files){
    $html=Get-Content -LiteralPath $f.FullName -Raw -Encoding UTF8
    $cnt=([regex]::Matches($html,'id="mong-company-info"')).Count
    if($cnt -eq 1){$visible++}
    if($cnt -gt 1){$duplicates++}
}

Write-Host ""
Write-Host "=== 사업자 정보 표시 강화 완료 ==="
Write-Host "전체 HTML:" $files.Count
Write-Host "수정 페이지:" $modified
Write-Host "기존 블록 없음:" $missingMarker
Write-Host "실패:" $failed.Count
Write-Host "표시 블록 정상 페이지:" $visible
Write-Host "중복 블록 페이지:" $duplicates
Write-Host "백업 폴더:" $backup

if($failed.Count -gt 0){
    $report=Join-Path $root "data\business-footer-visible-fail-$stamp.txt"
    $failed | Set-Content -LiteralPath $report -Encoding UTF8
    Write-Host "실패 보고서:" $report
}
