param(
    [switch]$FromGenerator
)

$ErrorActionPreference = "Stop"

$root       = "E:\mong-lounge-site"
$dataDir    = Join-Path $root "data"
$template   = Join-Path $dataDir "shop-detail-template.html"
$previewCsv = Join-Path $dataDir "shop-generation-preview.csv"
$reportCsv  = Join-Path $dataDir "shop-generation-run-report.csv"
$xlsx       = Join-Path $dataDir "몽라운지_업체자동등록_읽기용.xlsx"
$generator  = Join-Path $dataDir "mong-lounge-generate-all-shops.ps1"

$smsText = "건마천사보고 연락드립니다 예약 가능 한가요?"
$stamp = Get-Date -Format "yyyyMMdd-HHmmss"
$backupRoot = "E:\mong-lounge-detail-template-backup-$stamp"

foreach ($p in @($template,$previewCsv,$reportCsv,$xlsx)) {
    if (!(Test-Path -LiteralPath $p)) {
        throw "필수 파일을 찾을 수 없습니다: $p"
    }
}

function HtmlEncode([string]$text) {
    if ($null -eq $text) { return "" }
    return [System.Net.WebUtility]::HtmlEncode($text)
}

function Format-Price([string]$value) {
    if ([string]::IsNullOrWhiteSpace($value)) { return "" }
    $digits = $value -replace "[^\d]",""
    if ($digits) {
        try { return ("{0:N0}원" -f [long]$digits) } catch {}
    }
    return $value.Trim()
}

function Get-PhoneLink([string]$phone) {
    if ([string]::IsNullOrWhiteSpace($phone)) { return "" }
    return ($phone -replace "[^0-9+]","")
}

function Get-RelativeToRoot([string]$fullPath) {
    $full = [System.IO.Path]::GetFullPath($fullPath)
    $base = [System.IO.Path]::GetFullPath($root)
    if ($full.StartsWith($base,[System.StringComparison]::OrdinalIgnoreCase)) {
        return $full.Substring($base.Length).TrimStart("\")
    }
    return [System.IO.Path]::GetFileName($full)
}

function Backup-File([string]$filePath) {
    if (!(Test-Path -LiteralPath $filePath)) { return }
    $rel = Get-RelativeToRoot $filePath
    $dst = Join-Path $backupRoot $rel
    New-Item -ItemType Directory -Force -Path (Split-Path -Parent $dst) | Out-Null
    Copy-Item -LiteralPath $filePath -Destination $dst -Force
}

function Replace-GroupText([string]$html,[string]$pattern,[string]$newValue) {
    return [regex]::Replace(
        $html,
        $pattern,
        {
            param($m)
            return $m.Groups[1].Value + $newValue + $m.Groups[2].Value
        },
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor
        [System.Text.RegularExpressions.RegexOptions]::Singleline
    )
}

# ---------------------------------------------------------
# 엑셀 코스정보 읽기
# ---------------------------------------------------------

Add-Type -AssemblyName System.IO.Compression.FileSystem

function Read-ZipText($zip,[string]$name) {
    $entry = $zip.Entries | Where-Object FullName -eq $name | Select-Object -First 1
    if ($null -eq $entry) { return $null }
    $sr = New-Object System.IO.StreamReader($entry.Open())
    try { return $sr.ReadToEnd() } finally { $sr.Close() }
}

function Get-CellValue($cell,$sharedStrings) {
    $v = [string]$cell.v
    $t = [string]$cell.t
    if ($t -eq "s" -and $v -ne "") { return [string]$sharedStrings[[int]$v] }
    if ($t -eq "inlineStr") {
        if ($cell.is.t) { return [string]$cell.is.t }
        if ($cell.is.r) { return (($cell.is.r | ForEach-Object { [string]$_.t }) -join "") }
    }
    return $v
}

function Read-SheetRows($zip,[string]$sheetPath,$sharedStrings) {
    $text = Read-ZipText $zip $sheetPath
    if (!$text) { throw "엑셀 시트를 읽지 못했습니다: $sheetPath" }
    [xml]$xml = $text
    $rows = @()
    foreach ($row in $xml.worksheet.sheetData.row) {
        if ([int]$row.r -le 1) { continue }
        $h = @{}
        foreach ($cell in $row.c) {
            $col = ([regex]::Match([string]$cell.r,'^[A-Z]+')).Value
            $h[$col] = Get-CellValue $cell $sharedStrings
        }
        $rows += ,$h
    }
    return $rows
}

$zip = [System.IO.Compression.ZipFile]::OpenRead($xlsx)
try {
    $shared = @()
    $ssText = Read-ZipText $zip "xl/sharedStrings.xml"
    if ($ssText) {
        [xml]$ssXml = $ssText
        foreach ($si in $ssXml.sst.si) {
            if ($si.t) { $shared += [string]$si.t }
            elseif ($si.r) { $shared += (($si.r | ForEach-Object { [string]$_.t }) -join "") }
            else { $shared += "" }
        }
    }
    $courseRows = @(Read-SheetRows $zip "xl/worksheets/sheet3.xml" $shared)
}
finally {
    $zip.Dispose()
}

$courses = @(
    $courseRows |
    Where-Object {
        ([string]$_.A).Trim().ToUpperInvariant() -ne "N" -and
        -not [string]::IsNullOrWhiteSpace([string]$_.B) -and
        -not [string]::IsNullOrWhiteSpace([string]$_.C)
    } |
    ForEach-Object {
        $ord = 9999
        [void][int]::TryParse(([string]$_.G),[ref]$ord)
        [pscustomobject]@{
            업체명 = ([string]$_.B).Trim()
            코스명 = ([string]$_.C).Trim()
            시간 = ([string]$_.D).Trim()
            가격 = Format-Price ([string]$_.E)
            정렬 = $ord
        }
    }
)

$courseMap = @{}
foreach ($g in ($courses | Group-Object 업체명)) {
    $courseMap[$g.Name] = @($g.Group | Sort-Object 정렬,코스명,시간)
}

# ---------------------------------------------------------
# 현재 3831개 생성대상과 상세파일 연결
# ---------------------------------------------------------

$preview = @(Import-Csv -LiteralPath $previewCsv)
$report  = @(Import-Csv -LiteralPath $reportCsv)

$previewMap = @{}
foreach ($r in $preview) {
    $key = "$($r.업체명)|$($r.광역지역)|$($r.상위지역)|$($r.하위지역)"
    $previewMap[$key] = $r
}

$targets = New-Object System.Collections.Generic.List[object]
$missing = New-Object System.Collections.Generic.List[object]

foreach ($rr in $report) {
    if ([string]$rr.상태 -like "실패*") { continue }

    $key = "$($rr.업체명)|$($rr.광역지역)|$($rr.상위지역)|$($rr.하위지역)"
    if (!$previewMap.ContainsKey($key)) {
        $missing.Add($rr)
        continue
    }

    $p = $previewMap[$key]
    $targets.Add([pscustomobject]@{
        업체명=$p.업체명
        영업시간=$p.영업시간
        전화번호=$p.전화번호
        광역지역=$p.광역지역
        상위지역=$p.상위지역
        하위지역=$p.하위지역
        대표역=$p.대표역
        대표역주소=$p.대표역주소
        한줄소개=$p.한줄소개
        공지사항=([string]$p.공지사항).Replace("인근 주변","인근 생활권")
        업체소개=([string]$p.업체소개).Replace("인근 주변","인근 생활권")
        상세페이지=$rr.상세페이지
    })
}

if ($missing.Count -gt 0) {
    $mfile = Join-Path $dataDir "detail-template-missing-preview.csv"
    $missing | Export-Csv -LiteralPath $mfile -NoTypeInformation -Encoding UTF8
    Write-Host "미리보기 연결 누락:" $missing.Count
    Write-Host "저장:" $mfile
    Write-Host "상세페이지 수정: 0"
    exit 2
}

$vendorsNoCourse = @(
    $targets |
    Select-Object -ExpandProperty 업체명 -Unique |
    Where-Object { !$courseMap.ContainsKey($_) -or @($courseMap[$_]).Count -eq 0 }
)

if ($vendorsNoCourse.Count -gt 0) {
    Write-Host "코스 없는 업체가 있어 중단합니다."
    $vendorsNoCourse | ForEach-Object { Write-Host " - $_" }
    Write-Host "상세페이지 수정: 0"
    exit 3
}

$baseTemplate = Get-Content -LiteralPath $template -Raw -Encoding UTF8

# 템플릿 필수 구조 검사
$checks = @(
    'ml-course-row',
    '<b>지역</b>',
    '<b>대표역</b>',
    '<b>주소</b>',
    '<b>영업시간</b>',
    '<b>전화번호</b>',
    '<h2>공지사항</h2>',
    '<h2>업체소개</h2>',
    'ml-sticky-call',
    'ml-faq-section'
)

$badStructure = @($checks | Where-Object { $baseTemplate -notlike "*$_*" })
if ($badStructure.Count -gt 0) {
    Write-Host "템플릿 구조 확인 실패:"
    $badStructure | ForEach-Object { Write-Host " - $_" }
    Write-Host "상세페이지 수정: 0"
    exit 4
}

# ---------------------------------------------------------
# 템플릿 그대로 두고 데이터만 교체
# ---------------------------------------------------------

New-Item -ItemType Directory -Force -Path $backupRoot | Out-Null

Write-Host ""
Write-Host "=== 기준 템플릿 상세페이지 적용 시작 ==="
Write-Host "대상:" $targets.Count
Write-Host "백업:" $backupRoot

$updated = 0
$failed = 0
$i = 0

$oldSmsEncoded = [uri]::EscapeDataString($smsText)

foreach ($r in $targets) {
    $i++
    $file = [string]$r.상세페이지

    if ([string]::IsNullOrWhiteSpace($file)) {
        $failed++
        continue
    }

    try {
        $html = $baseTemplate

        $vendor = [string]$r.업체명
        $hours = [string]$r.영업시간
        $phone = [string]$r.전화번호
        $phoneLink = Get-PhoneLink $phone
        $smsHref = "sms:$phoneLink?body=$([uri]::EscapeDataString($smsText))"

        $geo = [string]$r.광역지역
        $parent = [string]$r.상위지역
        $area = [string]$r.하위지역

        $station = if (-not [string]::IsNullOrWhiteSpace([string]$r.대표역)) {
            [string]$r.대표역
        } else {
            "$area 중심"
        }

        $address = if (-not [string]::IsNullOrWhiteSpace([string]$r.대표역주소)) {
            [string]$r.대표역주소
        } else {
            "$area 중심"
        }

        $oneLine = [string]$r.한줄소개
        $notice = [string]$r.공지사항
        $about = [string]$r.업체소개

        $title = "$vendor | $parent $area 이용 안내"
        $meta = "$parent $area 지역의 $vendor 안내입니다. 영업시간, 코스와 가격, 전화·문자 문의, 주변 위치 정보를 확인할 수 있습니다."

        # title / meta / OG
        $html = [regex]::Replace($html,'(?is)<title>.*?</title>',"<title>$(HtmlEncode $title)</title>")
        $html = Replace-GroupText $html '(<meta\s+name=["'']description["''][^>]*content=["''])[^"'']*(["''][^>]*>)' (HtmlEncode $meta)
        $html = Replace-GroupText $html '(<meta\s+property=["'']og:title["''][^>]*content=["''])[^"'']*(["''][^>]*>)' (HtmlEncode $title)
        $html = Replace-GroupText $html '(<meta\s+property=["'']og:description["''][^>]*content=["''])[^"'']*(["''][^>]*>)' (HtmlEncode $meta)

        # 코스행만 교체: 디자인/CSS/섹션 구조는 그대로 유지
        $courseHtml = @(
            foreach ($c in @($courseMap[$vendor])) {
@"
        <div class="ml-course-row">
          <div class="ml-course-name">$(HtmlEncode $c.코스명)</div>
          <div class="ml-course-time">$(HtmlEncode $c.시간)</div>
          <div class="ml-course-price">$(HtmlEncode $c.가격)</div>
        </div>
"@
            }
        ) -join ""

        $coursePattern = '(?is)(?:\s*<div class="ml-course-row">\s*<div class="ml-course-name">.*?</div>\s*<div class="ml-course-time">.*?</div>\s*<div class="ml-course-price">.*?</div>\s*</div>)+'
        $html = [regex]::Replace($html,$coursePattern,"`r`n$courseHtml",1)

        # 업체정보 행
        $html = Replace-GroupText $html '(<b>지역</b>\s*<span>).*?(</span>)' (HtmlEncode "$geo $parent $area")
        $html = Replace-GroupText $html '(<b>대표역</b>\s*<span>).*?(</span>)' (HtmlEncode $station)
        $html = Replace-GroupText $html '(<b>주소</b>\s*<span>).*?(</span>)' (HtmlEncode $address)
        $html = Replace-GroupText $html '(<b>영업시간</b>\s*<span>).*?(</span>)' (HtmlEncode $hours)
        $html = Replace-GroupText $html '(<b>전화번호</b>\s*<span>).*?(</span>)' (HtmlEncode $phone)

        # 공지사항 / 업체소개
        $html = Replace-GroupText $html '(<h2>공지사항</h2>\s*<p class="ml-copy">).*?(</p>)' (HtmlEncode $notice)
        $html = Replace-GroupText $html '(<h2>업체소개</h2>\s*<p class="ml-copy">).*?(</p>)' (HtmlEncode $about)

        # FAQ 영업시간
        $faqHours = if ($hours -eq "24시간") {
            "24시간, 휴일에도 가능합니다."
        } else {
            "영업시간은 $hours 이며, 야간·휴일 가능 여부는 이용 전 문의해 주세요."
        }
        $html = $html.Replace("24시간, 휴일에도 가능합니다.",(HtmlEncode $faqHours))

        # 전화 / 문자 href
        $html = [regex]::Replace($html,'href="tel:[^"]*"',"href=`"tel:$(HtmlEncode $phoneLink)`"")
        $html = [regex]::Replace($html,'href="sms:[^"]*"',"href=`"$(HtmlEncode $smsHref)`"")

        # 템플릿 원본의 화면 텍스트를 현재 데이터로 교체
        # 긴 문자열부터 바꿔서 디자인 마크업은 건드리지 않음
        $html = $html.Replace("강남구 역삼동 중심 한국미인테라피 이용 안내",(HtmlEncode $oneLine))
        $html = $html.Replace("서울 강남구 역삼동",(HtmlEncode "$geo $parent $area"))
        $html = $html.Replace("강남구 역삼동",(HtmlEncode "$parent $area"))
        $html = $html.Replace("서울특별시 강남구 테헤란로 지하156(역삼동)",(HtmlEncode $address))
        $html = $html.Replace("0507-1280-3345",(HtmlEncode $phone))
        $html = $html.Replace("050712803345",(HtmlEncode $phoneLink))
        $html = $html.Replace("역삼역",(HtmlEncode $station))
        $html = $html.Replace("한국미인테라피",(HtmlEncode $vendor))
        $html = $html.Replace("24시간",(HtmlEncode $hours))
        $html = $html.Replace("역삼동",(HtmlEncode $area))
        $html = $html.Replace("강남구",(HtmlEncode $parent))
        $html = $html.Replace("서울",(HtmlEncode $geo))

        # 고정 SMS 안내문은 그대로
        # 생성 마커 추가
        if ($html -notmatch 'ML-DETAIL-TEMPLATE-PAGE') {
            $html = $html.Replace("<body>","<body>`r`n<!-- ML-DETAIL-TEMPLATE-PAGE -->")
        }

        if (Test-Path -LiteralPath $file) {
            Backup-File $file
        } else {
            New-Item -ItemType Directory -Force -Path (Split-Path -Parent $file) | Out-Null
        }

        Set-Content -LiteralPath $file -Value $html -Encoding UTF8
        $updated++
    }
    catch {
        $failed++
        Write-Host "실패: $file"
        Write-Host $_.Exception.Message
    }

    if (($i % 250) -eq 0 -or $i -eq $targets.Count) {
        Write-Host "진행: $i / $($targets.Count)"
    }
}

# ---------------------------------------------------------
# 검증
# ---------------------------------------------------------

$markerOk = 0
$courseOk = 0
$smsOk = 0
$faqOk = 0
$pinkStyleOk = 0

foreach ($r in $targets) {
    $file = [string]$r.상세페이지
    if (!(Test-Path -LiteralPath $file)) { continue }

    $t = Get-Content -LiteralPath $file -Raw -Encoding UTF8
    if ($t -match 'ML-DETAIL-TEMPLATE-PAGE') { $markerOk++ }
    if ($t -match 'ml-course-row') { $courseOk++ }
    if ($t -match 'href="sms:' -and $t -match [regex]::Escape([uri]::EscapeDataString($smsText))) { $smsOk++ }
    if ($t -match 'ml-faq-section') { $faqOk++ }

    # 템플릿의 분홍 영역 관련 CSS/마크업이 유지됐는지:
    # 전체 템플릿에 존재하는 'pink' 또는 pink 계열 색/클래스가 페이지에도 남아있는지 확인
    if (
        ($baseTemplate -match '(?i)pink|#f[0-9a-f]{5}|#e[0-9a-f]{5}') -and
        ($t -match '(?i)pink|#f[0-9a-f]{5}|#e[0-9a-f]{5}')
    ) { $pinkStyleOk++ }
}

# ---------------------------------------------------------
# 앞으로 전체 자동생성기 실행 후 이 템플릿 자동 적용
# ---------------------------------------------------------

$hookApplied = $false
if (Test-Path -LiteralPath $generator) {
    $g = Get-Content -LiteralPath $generator -Raw -Encoding UTF8

    if ($g -notmatch 'ML-DETAIL-TEMPLATE-HOOK') {
        Backup-File $generator
        $hook = @'

# ML-DETAIL-TEMPLATE-HOOK
$mlDetailTemplateScript = Join-Path $dataDir "mong-lounge-apply-detail-template.ps1"
if (Test-Path -LiteralPath $mlDetailTemplateScript) {
    Write-Host ""
    Write-Host "=== 기준 상세페이지 템플릿 재적용 ==="
    & $mlDetailTemplateScript -FromGenerator
}
'@
        Add-Content -LiteralPath $generator -Value $hook -Encoding UTF8
        $hookApplied = $true
    }
    else {
        $hookApplied = $true
    }
}

Write-Host ""
Write-Host "=== 기준 상세페이지 디자인 적용 완료 ==="
Write-Host "대상:" $targets.Count
Write-Host "수정:" $updated
Write-Host "실패:" $failed
Write-Host ""
Write-Host "템플릿 마커:" "$markerOk / $($targets.Count)"
Write-Host "코스영역:" "$courseOk / $($targets.Count)"
Write-Host "문자버튼/문구:" "$smsOk / $($targets.Count)"
Write-Host "FAQ:" "$faqOk / $($targets.Count)"
Write-Host "분홍 디자인 유지 확인:" "$pinkStyleOk / $($targets.Count)"
Write-Host "자동생성기 템플릿 후처리 연결:" $hookApplied
Write-Host ""
Write-Host "백업 폴더:" $backupRoot
Write-Host "지역페이지 수정: 0"
