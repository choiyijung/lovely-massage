$ErrorActionPreference = "Stop"

$root = "E:\mong-lounge-site"
$stamp = Get-Date -Format "yyyyMMdd-HHmmss"
$backup = "E:\mong-lounge-business-footer-backup-$stamp"

$startMarker = "<!-- mong-company-info:start -->"
$endMarker   = "<!-- mong-company-info:end -->"

$block = @'
<!-- mong-company-info:start -->
<footer class="mong-company-info" aria-label="사업자 정보">
  <style>
    .mong-company-info{
      margin-top:48px;
      padding:28px 20px 30px;
      background:#082f2a;
      color:#f6f2e8;
      border-top:1px solid rgba(255,255,255,.12);
      font-family:inherit;
    }
    .mong-company-info__inner{
      width:min(1180px,100%);
      margin:0 auto;
    }
    .mong-company-info__title{
      margin:0 0 12px;
      font-size:15px;
      font-weight:800;
      color:#f4d889;
    }
    .mong-company-info__text{
      margin:0;
      font-size:12px;
      line-height:1.85;
      color:rgba(255,255,255,.82);
      word-break:keep-all;
    }
    .mong-company-info__text span{
      display:inline-block;
      margin-right:14px;
    }
    .mong-company-info a{
      color:inherit;
      text-decoration:none;
    }
    @media (max-width:640px){
      .mong-company-info{
        padding:24px 18px 26px;
      }
      .mong-company-info__text span{
        display:block;
        margin-right:0;
      }
    }
  </style>
  <div class="mong-company-info__inner">
    <p class="mong-company-info__title">사업자 정보</p>
    <p class="mong-company-info__text">
      <span>상호명: 주)에이치아이행복지킴이</span>
      <span>대표자: 최이정</span>
      <span>사업장주소: 서울 양천구 신월로 170 (신월동) 3층 304-24호</span>
      <span>연락처: <a href="tel:01083799746">01083799746</a></span>
      <span>사업자등록번호: 803-81-00107</span>
      <span>통신판매업신고번호: 2021-서울영등포-0841</span>
      <span>대표자 이메일: <a href="mailto:happyloanstar@gmail.com">happyloanstar@gmail.com</a></span>
    </p>
  </div>
</footer>
<!-- mong-company-info:end -->
'@

if(!(Test-Path -LiteralPath $root)){
    throw "사이트 폴더가 없습니다: $root"
}

New-Item -ItemType Directory -Force -Path $backup | Out-Null

$files = @(
    Get-ChildItem -LiteralPath $root -Recurse -File -Filter "*.html" |
    Where-Object {
        $_.FullName -notmatch '\\\.git\\|\\node_modules\\|\\backup\\|\\_backup\\'
    }
)

$modified = 0
$already = 0
$failed = New-Object System.Collections.Generic.List[string]

foreach($f in $files){
    try{
        $html = Get-Content -LiteralPath $f.FullName -Raw -Encoding UTF8
        $before = $html

        if($html.Contains($startMarker) -and $html.Contains($endMarker)){
            $pattern = '(?is)<!-- mong-company-info:start -->.*?<!-- mong-company-info:end -->'
            $html = [regex]::Replace($html,$pattern,[System.Text.RegularExpressions.MatchEvaluator]{ param($m) $block },1)
        }
        elseif($html -match '(?is)</body>'){
            $html = [regex]::Replace($html,'(?is)</body>',"`r`n$block`r`n</body>",1)
        }
        else{
            $html = $html.TrimEnd() + "`r`n`r`n" + $block + "`r`n"
        }

        if($html -eq $before){
            $already++
            continue
        }

        $rel = $f.FullName.Substring($root.Length).TrimStart("\")
        $dest = Join-Path $backup $rel
        New-Item -ItemType Directory -Force -Path (Split-Path $dest -Parent) | Out-Null
        Copy-Item -LiteralPath $f.FullName -Destination $dest -Force

        Set-Content -LiteralPath $f.FullName -Value $html -Encoding UTF8
        $modified++
    }
    catch{
        $failed.Add("$($f.FullName) | $($_.Exception.Message)")
    }
}

# 검증
$found = 0
$duplicates = 0
foreach($f in $files){
    $html = Get-Content -LiteralPath $f.FullName -Raw -Encoding UTF8
    $count = ([regex]::Matches($html,'<!-- mong-company-info:start -->')).Count
    if($count -ge 1){ $found++ }
    if($count -gt 1){ $duplicates++ }
}

Write-Host ""
Write-Host "=== 사업자 정보 하단 적용 완료 ==="
Write-Host "전체 HTML:" $files.Count
Write-Host "수정 페이지:" $modified
Write-Host "이미 동일:" $already
Write-Host "실패:" $failed.Count
Write-Host "사업자 정보 포함 페이지:" $found
Write-Host "중복 삽입 페이지:" $duplicates
Write-Host "백업 폴더:" $backup

if($failed.Count -gt 0){
    $report = Join-Path $root "data\business-footer-fail-$stamp.txt"
    $failed | Set-Content -LiteralPath $report -Encoding UTF8
    Write-Host "실패 보고서:" $report
}
