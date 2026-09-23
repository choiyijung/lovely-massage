$ErrorActionPreference = "Stop"

$root = "E:\mong-lounge-site"

$preview = Join-Path $root "data\shop-generation-preview.csv"

$regionPage = Join-Path $root "massage\seoul\gangnam-gu\강남구-출장마사지-압구정-청담-역삼-외출-미팅-일정이-길어진-날-목-어깨의-뻐근함-방문-홈케어\index.html"

$detailDir = Join-Path $root "shops\seoul\gangnam-gu\역삼동\한국미인테라피"
$detailPage = Join-Path $detailDir "index.html"

if (!(Test-Path -LiteralPath $preview)) {
    throw "미리보기 CSV를 찾을 수 없습니다: $preview"
}

if (!(Test-Path -LiteralPath $regionPage)) {
    throw "강남구 지역 페이지를 찾을 수 없습니다: $regionPage"
}

$row = Import-Csv -LiteralPath $preview |
    Where-Object {
        $_.업체명 -eq "한국미인테라피" -and
        $_.하위지역 -eq "역삼동"
    } |
    Select-Object -First 1

if ($null -eq $row) {
    throw "shop-generation-preview.csv에서 한국미인테라피 / 역삼동 행을 찾지 못했습니다."
}

function HtmlEncode([string]$text) {
    if ($null -eq $text) { return "" }
    return [System.Net.WebUtility]::HtmlEncode($text)
}

function Digits([string]$text) {
    return ($text -replace '[^\d]','')
}

$shopName = [string]$row.업체명
$hours = [string]$row.영업시간
$phone = [string]$row.전화번호
$phoneDigits = Digits $phone
$station = [string]$row.대표역
$stationAddress = [string]$row.대표역주소
$intro = [string]$row.한줄소개
$notice = [string]$row.공지사항
$about = [string]$row.업체소개
$courseSummary = [string]$row.코스요약

# 중복 표현 정리
$about = $about -replace '인근 주변', '인근 생활권'
$about = $about -replace '인근 생활권 생활권', '인근 생활권'

$smsMessage = "건마천사보고 연락드립니다 예약 가능 한가요?"
$smsBody = [Uri]::EscapeDataString($smsMessage)

# 코스요약을 상세 행으로 변환
$courseItems = @()

foreach ($part in @($courseSummary -split '\s*/\s*')) {
    $p = $part.Trim()
    if (!$p) { continue }

    $m = [regex]::Match(
        $p,
        '^(?<name>.+?)\s+(?<time>\d+\s*분)?\s*(?<price>[\d,]+\s*원)$'
    )

    if ($m.Success) {
        $courseItems += [pscustomobject]@{
            코스명 = $m.Groups["name"].Value.Trim()
            시간 = $m.Groups["time"].Value.Trim()
            가격 = $m.Groups["price"].Value.Trim()
        }
    }
    else {
        $courseItems += [pscustomobject]@{
            코스명 = $p
            시간 = ""
            가격 = ""
        }
    }
}

$courseHtml = ""

foreach ($c in $courseItems) {
    $courseHtml += @"
        <div class="ml-course-row">
          <div class="ml-course-name">$(HtmlEncode $c.코스명)</div>
          <div class="ml-course-time">$(HtmlEncode $c.시간)</div>
          <div class="ml-course-price">$(HtmlEncode $c.가격)</div>
        </div>
"@
}

if (!$courseHtml) {
    $courseHtml = '<div class="ml-empty">등록된 코스정보가 없습니다.</div>'
}

$stationDisplay = if ($station) { $station } else { "역삼동 중심" }
$addressDisplay = if ($stationAddress) { $stationAddress } else { "서울 강남구 역삼동" }

$utf8 = New-Object System.Text.UTF8Encoding($false)

# -----------------------------
# 백업
# -----------------------------
$stamp = Get-Date -Format "yyyyMMdd-HHmmss"
$backup = "E:\mong-lounge-v68-test-yeoksam-backup-$stamp"

New-Item -ItemType Directory -Path $backup -Force | Out-Null

Copy-Item -LiteralPath $regionPage -Destination (Join-Path $backup "gangnam-region-index.html") -Force

if (Test-Path -LiteralPath $detailPage) {
    Copy-Item -LiteralPath $detailPage -Destination (Join-Path $backup "yeoksam-shop-index.html") -Force
}

# -----------------------------
# 1. 역삼동 상세페이지 1개 생성
# -----------------------------
New-Item -ItemType Directory -Path $detailDir -Force | Out-Null

$detailHtml = @"
<!doctype html>
<html lang="ko">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>$(HtmlEncode $shopName) | 강남구 역삼동 출장마사지 업체 안내</title>
  <meta name="description" content="$(HtmlEncode $shopName) 강남구 역삼동 업체 안내입니다. 코스정보, 영업시간, 대표역과 위치정보, 전화 및 문자 예약 정보를 확인할 수 있습니다.">
  <link rel="stylesheet" href="../../../../../assets/css/style.css?v=68test1">

  <style>
    .ml-test-detail{max-width:640px;margin:0 auto;padding:20px 14px 60px;color:#222}
    .ml-test-back{display:inline-block;margin:0 0 12px;text-decoration:none;color:#555;font-size:13px}
    .ml-test-hero{overflow:hidden;border-radius:22px;background:#103c31;box-shadow:0 12px 32px rgba(0,0,0,.10)}
    .ml-test-hero-top{padding:34px 28px 28px;text-align:center}
    .ml-test-loc{font-size:12px;font-weight:800;color:#f0cf83;letter-spacing:.08em}
    .ml-test-title{margin:8px 0 8px;color:#f0cf83;font-size:34px;line-height:1.2}
    .ml-test-intro{margin:0 auto;color:#fff;font-size:15px;line-height:1.7;word-break:keep-all}
    .ml-test-feature{background:#ff82bd;padding:24px 28px;text-align:center;color:#fff;font-weight:900;line-height:1.6;text-shadow:0 1px 2px rgba(0,0,0,.28)}
    .ml-test-hero-bottom{padding:23px 28px 26px;text-align:center;color:#fff}
    .ml-test-hero-bottom strong{display:block;font-size:22px}
    .ml-test-hero-bottom span{display:block;margin-top:6px;font-size:13px;opacity:.84}
    .ml-test-contact{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin:14px 0}
    .ml-test-contact a{display:flex;align-items:center;justify-content:center;min-height:52px;border-radius:13px;text-decoration:none;font-weight:900}
    .ml-test-phone{background:#111;color:#fff}
    .ml-test-sms{background:#ececec;color:#111}
    .ml-test-tabs{display:grid;grid-template-columns:repeat(3,1fr);gap:8px;margin:20px 0 12px}
    .ml-test-tabs a{padding:12px 6px;text-align:center;border-radius:12px;background:#f2f0f7;text-decoration:none;color:#333;font-size:13px;font-weight:800}
    .ml-test-section{margin:12px 0;padding:22px;border:1px solid #e8e8e8;border-radius:18px;background:#fff}
    .ml-test-section h2{margin:0 0 17px;font-size:20px}
    .ml-course-row{display:grid;grid-template-columns:minmax(0,1.5fr) .65fr .8fr;gap:10px;padding:14px 0;border-bottom:1px solid #eee;align-items:center}
    .ml-course-row:last-child{border-bottom:0}
    .ml-course-name{font-weight:900}
    .ml-course-time{text-align:center;color:#666}
    .ml-course-price{text-align:right;font-weight:900;color:#c64d83}
    .ml-info-row{display:grid;grid-template-columns:100px 1fr;gap:12px;padding:11px 0;border-bottom:1px solid #eee}
    .ml-info-row:last-child{border-bottom:0}
    .ml-info-row b{color:#666;font-size:13px}
    .ml-info-row span{font-weight:700;word-break:keep-all}
    .ml-copy{margin:0;color:#444;line-height:1.85;word-break:keep-all}
    .ml-sms-note{margin-top:12px;padding:12px 14px;border-radius:12px;background:#faf4f7;color:#555;font-size:13px;line-height:1.6}
    @media(max-width:520px){
      .ml-test-title{font-size:29px}
      .ml-test-section{padding:18px}
      .ml-course-row{grid-template-columns:1fr auto}
      .ml-course-time{grid-row:2;grid-column:1;text-align:left}
      .ml-course-price{grid-row:1/3;grid-column:2}
      .ml-info-row{grid-template-columns:82px 1fr}
    }
  </style>
</head>

<body>
  <main class="ml-test-detail">
    <a class="ml-test-back" href="../../../../../massage/seoul/gangnam-gu/강남구-출장마사지-압구정-청담-역삼-외출-미팅-일정이-길어진-날-목-어깨의-뻐근함-방문-홈케어/index.html">← 강남구 업체 목록</a>

    <section class="ml-test-hero">
      <div class="ml-test-hero-top">
        <div class="ml-test-loc">서울 · 강남구 · 역삼동</div>
        <h1 class="ml-test-title">$(HtmlEncode $shopName)</h1>
        <p class="ml-test-intro">$(HtmlEncode $intro)</p>
      </div>

      <div class="ml-test-feature">$(HtmlEncode $stationDisplay) · $(HtmlEncode $hours) 운영</div>

      <div class="ml-test-hero-bottom">
        <strong>$(HtmlEncode $hours) 전화·문자 예약</strong>
        <span>예약 가능 여부는 전화 또는 문자로 확인해 주세요.</span>
      </div>
    </section>

    <div class="ml-test-contact">
      <a class="ml-test-phone" href="tel:$phoneDigits">전화 문의</a>
      <a class="ml-test-sms" href="sms:$phoneDigits?body=$smsBody">문자 문의</a>
    </div>

    <div class="ml-test-tabs">
      <a href="#course">코스정보</a>
      <a href="#info">업체정보</a>
      <a href="#notice">안내사항</a>
    </div>

    <section id="course" class="ml-test-section">
      <h2>코스정보</h2>
$courseHtml
    </section>

    <section id="info" class="ml-test-section">
      <h2>업체정보</h2>

      <div class="ml-info-row">
        <b>지역</b>
        <span>서울 강남구 역삼동</span>
      </div>

      <div class="ml-info-row">
        <b>대표역</b>
        <span>$(HtmlEncode $stationDisplay)</span>
      </div>

      <div class="ml-info-row">
        <b>주소</b>
        <span>$(HtmlEncode $addressDisplay)</span>
      </div>

      <div class="ml-info-row">
        <b>영업시간</b>
        <span>$(HtmlEncode $hours)</span>
      </div>

      <div class="ml-info-row">
        <b>전화번호</b>
        <span>$(HtmlEncode $phone)</span>
      </div>
    </section>

    <section id="notice" class="ml-test-section">
      <h2>공지사항</h2>
      <p class="ml-copy">$(HtmlEncode $notice)</p>

      <div class="ml-sms-note">
        문자 문의 버튼을 누르면<br>
        <strong>$(HtmlEncode $smsMessage)</strong><br>
        문구가 휴대폰 문자 작성창에 자동으로 입력됩니다.
      </div>
    </section>

    <section class="ml-test-section">
      <h2>업체소개</h2>
      <p class="ml-copy">$(HtmlEncode $about)</p>
    </section>
  </main>
</body>
</html>
"@

[System.IO.File]::WriteAllText($detailPage,$detailHtml,$utf8)

# -----------------------------
# 2. 강남구 페이지 업체 카드 1개 연결
# -----------------------------
$regionHtml = [System.IO.File]::ReadAllText($regionPage,$utf8)

$cardBlock = @"
<!-- SHOP-TEST-V68-START -->
<style>
  .ml-shop-test-wrap{margin:18px 0 24px}
  .ml-shop-test-label{margin:0 0 10px;font-size:13px;font-weight:900;color:#555}
  .ml-shop-test-card{display:block;padding:20px 21px;border-radius:17px;background:#103c31;color:#fff;text-decoration:none;box-shadow:0 9px 24px rgba(0,0,0,.08)}
  .ml-shop-test-card:hover{transform:translateY(-1px)}
  .ml-shop-test-top{display:flex;justify-content:space-between;gap:12px;align-items:center;margin-bottom:11px}
  .ml-shop-test-area{font-size:14px;font-weight:900;color:#f0cf83}
  .ml-shop-test-cat{font-size:11px;font-weight:800;opacity:.82}
  .ml-shop-test-intro{margin:0 0 14px;font-size:14px;line-height:1.65;word-break:keep-all}
  .ml-shop-test-name{font-size:20px;font-weight:900;color:#f0cf83}
</style>
<section class="ml-shop-test-wrap">
  <div class="ml-shop-test-label">역삼동 등록 업체</div>
  <a class="ml-shop-test-card" href="../../../../shops/seoul/gangnam-gu/역삼동/한국미인테라피/index.html">
    <div class="ml-shop-test-top">
      <span class="ml-shop-test-area">역삼동</span>
      <span class="ml-shop-test-cat">출장마사지</span>
    </div>
    <p class="ml-shop-test-intro">$(HtmlEncode $intro)</p>
    <div class="ml-shop-test-name">$(HtmlEncode $shopName)</div>
  </a>
</section>
<!-- SHOP-TEST-V68-END -->
"@

if ($regionHtml -match '(?s)<!-- SHOP-TEST-V68-START -->.*?<!-- SHOP-TEST-V68-END -->') {
    $regionHtml = [regex]::Replace(
        $regionHtml,
        '(?s)<!-- SHOP-TEST-V68-START -->.*?<!-- SHOP-TEST-V68-END -->',
        [System.Text.RegularExpressions.MatchEvaluator]{ param($m) $cardBlock }
    )
}
elseif ($regionHtml -match '(?s)<!-- SHOP-LIST-START -->.*?<!-- SHOP-LIST-END -->') {
    $regionHtml = [regex]::Replace(
        $regionHtml,
        '(?s)<!-- SHOP-LIST-START -->.*?<!-- SHOP-LIST-END -->',
        [System.Text.RegularExpressions.MatchEvaluator]{
            param($m)
            "<!-- SHOP-LIST-START -->`r`n$cardBlock`r`n<!-- SHOP-LIST-END -->"
        }
    )
}
elseif ($regionHtml -match '<main\b[^>]*>') {
    $regionHtml = [regex]::Replace(
        $regionHtml,
        '<main\b[^>]*>',
        [System.Text.RegularExpressions.MatchEvaluator]{
            param($m)
            $m.Value + "`r`n" + $cardBlock
        },
        1
    )
}
else {
    throw "강남구 페이지에서 업체 카드를 넣을 위치를 찾지 못했습니다."
}

[System.IO.File]::WriteAllText($regionPage,$regionHtml,$utf8)

# -----------------------------
# 검증
# -----------------------------
$checkDetail = [System.IO.File]::ReadAllText($detailPage,$utf8)
$checkRegion = [System.IO.File]::ReadAllText($regionPage,$utf8)

Write-Host ""
Write-Host "=== 역삼동 1개 업체 테스트 생성 결과 ==="
Write-Host "상세페이지 생성:" (Test-Path -LiteralPath $detailPage)
Write-Host "코스 개수:" $courseItems.Count
Write-Host "전화번호 포함:" $checkDetail.Contains($phone)
Write-Host "문자 자동문구 포함:" $checkDetail.Contains($smsBody)
Write-Host "역삼역 포함:" $checkDetail.Contains("역삼역")
Write-Host "강남구 카드 포함:" $checkRegion.Contains("SHOP-TEST-V68-START")
Write-Host "상세페이지 링크 포함:" $checkRegion.Contains("../../../../shops/seoul/gangnam-gu/역삼동/한국미인테라피/index.html")
Write-Host ""
Write-Host "상세페이지:"
Write-Host $detailPage
Write-Host ""
Write-Host "백업:"
Write-Host $backup
Write-Host ""
Write-Host "역삼동 1개 테스트 생성 완료"

