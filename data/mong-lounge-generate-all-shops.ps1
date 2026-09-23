param(
    [switch]$SkipPreviewRefresh
)

$ErrorActionPreference = "Stop"

# =========================================================
# 몽라운지 전체 업체 자동생성기
# - 엑셀 최신값 읽기
# - 업체 상세페이지 지역별 생성/업데이트
# - 지역 페이지에는 같은 업체 카드 1개만 표시
# - 기존 파일 백업
# - 재실행 시 자동 업데이트
# =========================================================

$root            = "E:\mong-lounge-site"
$dataDir         = Join-Path $root "data"
$massageRoot     = Join-Path $root "massage"
$shopsRoot       = Join-Path $root "shops"

$masterXlsx      = Join-Path $dataDir "몽라운지_업체자동등록_마스터_v2.xlsx"
$readXlsx        = Join-Path $dataDir "몽라운지_업체자동등록_읽기용.xlsx"
$previewScript   = Join-Path $dataDir "mong-lounge-build-shop-preview.ps1"
$previewCsv      = Join-Path $dataDir "shop-generation-preview.csv"
$locationCsv     = Join-Path $dataDir "region-location-master-final-v2.csv"

$siteAreaCsv     = Join-Path $dataDir "site-area-links.csv"
$otherFinalCsv   = Join-Path $dataDir "other-region-structured-final.csv"
$otherAreasCsv   = Join-Path $dataDir "other-region-areas.csv"

$runStamp        = Get-Date -Format "yyyyMMdd-HHmmss"
$backupRoot      = "E:\mong-lounge-full-generator-backup-$runStamp"
$reportCsv       = Join-Path $dataDir "shop-generation-run-report.csv"
$unresolvedCsv   = Join-Path $dataDir "shop-generation-unresolved-region-pages.csv"

$smsText         = "건마천사보고 연락드립니다 예약 가능 한가요?"

foreach ($p in @($masterXlsx,$previewScript,$locationCsv)) {
    if (!(Test-Path -LiteralPath $p)) {
        throw "필수 파일을 찾을 수 없습니다: $p"
    }
}

New-Item -ItemType Directory -Force -Path $shopsRoot | Out-Null

# ----------------------------
# 공통 함수
# ----------------------------

function HtmlEncode([string]$text) {
    if ($null -eq $text) { return "" }
    return [System.Net.WebUtility]::HtmlEncode($text)
}

function Safe-PathPart([string]$text) {
    if ([string]::IsNullOrWhiteSpace($text)) { return "unknown" }

    $s = $text.Trim()
    foreach ($ch in [System.IO.Path]::GetInvalidFileNameChars()) {
        $s = $s.Replace([string]$ch, "-")
    }
    $s = $s.Trim(" ",".")
    if ([string]::IsNullOrWhiteSpace($s)) { return "unknown" }
    return $s
}

function Format-Price([string]$value) {
    if ([string]::IsNullOrWhiteSpace($value)) { return "" }

    $digits = $value -replace "[^\d]",""
    if (-not [string]::IsNullOrWhiteSpace($digits)) {
        try {
            return ("{0:N0}원" -f [long]$digits)
        } catch {}
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
    $dest = Join-Path $backupRoot $rel
    $destDir = Split-Path -Parent $dest

    New-Item -ItemType Directory -Force -Path $destDir | Out-Null
    Copy-Item -LiteralPath $filePath -Destination $dest -Force
}

function Get-ApplyRoot([string]$applyRegion) {
    switch ($applyRegion) {
        "서울"       { return Join-Path $massageRoot "seoul" }
        "경기"       { return Join-Path $massageRoot "gyeonggi" }
        "인천"       { return Join-Path $massageRoot "incheon" }
        "천안"       { return Join-Path $massageRoot "other\cheonan" }
        "아산"       { return Join-Path $massageRoot "other\asan" }
        "대전"       { return Join-Path $massageRoot "other\daejeon" }
        "청주"       { return Join-Path $massageRoot "other\cheongju" }
        "창원"       { return Join-Path $massageRoot "other\changwon" }
        "대구"       { return Join-Path $massageRoot "other\daegu" }
        "공주"       { return Join-Path $massageRoot "other\gongju" }
        "계룡"       { return Join-Path $massageRoot "other\gyeryong" }
        "전주"       { return Join-Path $massageRoot "other\jeonju" }
        "광주광역시" { return Join-Path $massageRoot "other\gwangju" }
        "익산"       { return Join-Path $massageRoot "other\iksan" }
        "김해"       { return Join-Path $massageRoot "other\gimhae" }
        default      { return $null }
    }
}

function Get-ParentTokens([string]$parent,[string]$applyRegion) {
    $tokens = New-Object System.Collections.Generic.List[string]

    if (-not [string]::IsNullOrWhiteSpace($parent)) {
        $tokens.Add($parent)

        if ($parent.Contains("-")) {
            $last = ($parent -split "-")[-1]
            if (-not [string]::IsNullOrWhiteSpace($last)) {
                $tokens.Add($last)
            }
        }

        if ($parent -notmatch "(시|군|구)$") {
            $tokens.Add("${parent}시")
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($applyRegion)) {
        $tokens.Add($applyRegion)

        if ($applyRegion -in @("천안","아산","공주","계룡","청주","전주","익산","김해","창원")) {
            $tokens.Add("${applyRegion}시")
        }
    }

    return @($tokens | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
}

# ----------------------------
# 엑셀 읽기
# ----------------------------

Add-Type -AssemblyName System.IO.Compression.FileSystem

function Read-ZipText($zip,[string]$name) {
    $entry = $zip.Entries | Where-Object FullName -eq $name | Select-Object -First 1
    if ($null -eq $entry) { return $null }

    $sr = New-Object System.IO.StreamReader($entry.Open())
    try { return $sr.ReadToEnd() }
    finally { $sr.Close() }
}

function Get-CellValue($cell,$sharedStrings) {
    $v = [string]$cell.v
    $t = [string]$cell.t

    if ($t -eq "s" -and $v -ne "") {
        return [string]$sharedStrings[[int]$v]
    }

    if ($t -eq "inlineStr") {
        if ($cell.is.t) { return [string]$cell.is.t }
        if ($cell.is.r) {
            return (($cell.is.r | ForEach-Object { [string]$_.t }) -join "")
        }
    }

    return $v
}

function Read-SheetRows($zip,[string]$sheetPath,$sharedStrings) {
    $text = Read-ZipText $zip $sheetPath
    if ([string]::IsNullOrWhiteSpace($text)) {
        throw "엑셀 시트를 읽지 못했습니다: $sheetPath"
    }

    [xml]$xml = $text
    $result = @()

    foreach ($row in $xml.worksheet.sheetData.row) {
        if ([int]$row.r -le 1) { continue }

        $h = @{}
        foreach ($cell in $row.c) {
            $col = ([regex]::Match([string]$cell.r,'^[A-Z]+')).Value
            $h[$col] = Get-CellValue $cell $sharedStrings
        }

        $result += ,$h
    }

    return $result
}

function Read-MasterWorkbook([string]$xlsxPath) {
    $zip = [System.IO.Compression.ZipFile]::OpenRead($xlsxPath)

    try {
        $shared = @()
        $ssText = Read-ZipText $zip "xl/sharedStrings.xml"

        if ($ssText) {
            [xml]$ssXml = $ssText
            foreach ($si in $ssXml.sst.si) {
                if ($si.t) {
                    $shared += [string]$si.t
                }
                elseif ($si.r) {
                    $shared += (($si.r | ForEach-Object { [string]$_.t }) -join "")
                }
                else {
                    $shared += ""
                }
            }
        }

        $basicRows  = @(Read-SheetRows $zip "xl/worksheets/sheet2.xml" $shared)
        $courseRows = @(Read-SheetRows $zip "xl/worksheets/sheet3.xml" $shared)

        return [pscustomobject]@{
            BasicRows  = $basicRows
            CourseRows = $courseRows
        }
    }
    finally {
        $zip.Dispose()
    }
}

# ----------------------------
# 최신 엑셀 -> 읽기용 복사 -> 미리보기 재생성
# ----------------------------

if (-not $SkipPreviewRefresh) {
    Write-Host ""
    Write-Host "=== 1. 최신 엑셀 읽기 ==="

    $srcStream = [System.IO.File]::Open(
        $masterXlsx,
        [System.IO.FileMode]::Open,
        [System.IO.FileAccess]::Read,
        [System.IO.FileShare]::ReadWrite
    )

    $dstStream = [System.IO.File]::Create($readXlsx)

    try {
        $srcStream.CopyTo($dstStream)
    }
    finally {
        $dstStream.Close()
        $srcStream.Close()
    }

    Write-Host "읽기용 엑셀 갱신 완료"

    Write-Host ""
    Write-Host "=== 2. 업체 생성 미리보기 갱신 ==="
    & $previewScript
}

if (!(Test-Path -LiteralPath $previewCsv)) {
    throw "미리보기 CSV가 없습니다: $previewCsv"
}

# ----------------------------
# 업체/코스/지역 데이터 준비
# ----------------------------

Write-Host ""
Write-Host "=== 3. 실제 생성 데이터 준비 ==="

$workbook = Read-MasterWorkbook $readXlsx

$shops = @(
    $workbook.BasicRows |
    Where-Object {
        ([string]$_.A).Trim().ToUpperInvariant() -ne "N" -and
        -not [string]::IsNullOrWhiteSpace([string]$_.B)
    } |
    ForEach-Object {
        [pscustomobject]@{
            업체명   = ([string]$_.B).Trim()
            영업시간 = ([string]$_.C).Trim()
            전화번호 = ([string]$_.D).Trim()
        }
    }
)

$courses = @(
    $workbook.CourseRows |
    Where-Object {
        ([string]$_.A).Trim().ToUpperInvariant() -ne "N" -and
        -not [string]::IsNullOrWhiteSpace([string]$_.B) -and
        -not [string]::IsNullOrWhiteSpace([string]$_.C)
    } |
    ForEach-Object {
        $order = 9999
        [void][int]::TryParse(([string]$_.G),[ref]$order)

        [pscustomobject]@{
            업체명   = ([string]$_.B).Trim()
            코스명   = ([string]$_.C).Trim()
            시간     = ([string]$_.D).Trim()
            가격     = Format-Price ([string]$_.E)
            코스설명 = ([string]$_.F).Trim()
            정렬순서 = $order
        }
    }
)

$previewRows = @(Import-Csv -LiteralPath $previewCsv)
$locationRows = @(Import-Csv -LiteralPath $locationCsv)

Write-Host "업체:" $shops.Count
Write-Host "코스:" $courses.Count
Write-Host "상세페이지 대상:" $previewRows.Count
Write-Host "공용 지역 마스터:" $locationRows.Count

$shopMap = @{}
foreach ($s in $shops) {
    $shopMap[$s.업체명] = $s
}

$courseMap = @{}
foreach ($g in ($courses | Group-Object 업체명)) {
    $courseMap[$g.Name] = @($g.Group | Sort-Object 정렬순서,코스명,시간)
}

# 지역키 -> 적용지역
$locationMap = @{}
foreach ($l in $locationRows) {
    $key = "$([string]$l.광역지역)|$([string]$l.상위지역)|$([string]$l.하위지역)"
    if (!$locationMap.ContainsKey($key)) {
        $locationMap[$key] = $l
    }
}

# ----------------------------
# 명시적 지역페이지 매핑
# ----------------------------

$explicitPageMap = @{}

function Add-ExplicitPageMappings([string]$csvPath) {
    if (!(Test-Path -LiteralPath $csvPath)) { return }

    $rows = @(Import-Csv -LiteralPath $csvPath)

    foreach ($r in $rows) {
        $props = @($r.PSObject.Properties.Name)

        $page = ""
        foreach ($name in @("원본페이지","페이지","SourcePage")) {
            if ($props -contains $name) {
                $page = [string]$r.$name
                break
            }
        }

        if ([string]::IsNullOrWhiteSpace($page) -or !(Test-Path -LiteralPath $page)) {
            continue
        }

        $parent = ""
        foreach ($name in @("상위지역","시군구","구군")) {
            if ($props -contains $name) {
                $parent = [string]$r.$name
                break
            }
        }

        $apply = ""
        foreach ($name in @("적용지역","지역")) {
            if ($props -contains $name) {
                $apply = [string]$r.$name
                break
            }
        }

        $geo = ""
        if ($props -contains "광역지역") {
            $geo = [string]$r.광역지역
        }

        if ([string]::IsNullOrWhiteSpace($apply) -and $geo -in @("서울","경기","인천")) {
            $apply = $geo
        }

        if (-not [string]::IsNullOrWhiteSpace($apply) -and
            -not [string]::IsNullOrWhiteSpace($parent)) {
            $key = "$apply|$parent"
            if (!$explicitPageMap.ContainsKey($key)) {
                $explicitPageMap[$key] = $page
            }
        }
    }
}

Add-ExplicitPageMappings $siteAreaCsv
Add-ExplicitPageMappings $otherFinalCsv
Add-ExplicitPageMappings $otherAreasCsv

# ----------------------------
# 지역 페이지 검색 인덱스
# ----------------------------

$rootFileCache = @{}
$fileTextCache = @{}
$regionPageCache = @{}

function Get-RootFiles([string]$rootPath) {
    if ([string]::IsNullOrWhiteSpace($rootPath) -or !(Test-Path -LiteralPath $rootPath)) {
        return @()
    }

    if (!$rootFileCache.ContainsKey($rootPath)) {
        $rootFileCache[$rootPath] = @(
            Get-ChildItem -LiteralPath $rootPath -Recurse -File -Filter "index.html"
        )
    }

    return @($rootFileCache[$rootPath])
}

function Get-PageInfo([string]$filePath) {
    if ($fileTextCache.ContainsKey($filePath)) {
        return $fileTextCache[$filePath]
    }

    $html = Get-Content -LiteralPath $filePath -Raw -Encoding UTF8

    $title = ""
    $h1 = ""

    $tm = [regex]::Match($html,'(?is)<title\b[^>]*>(.*?)</title>')
    if ($tm.Success) {
        $title = [System.Net.WebUtility]::HtmlDecode(
            [regex]::Replace($tm.Groups[1].Value,'(?is)<[^>]+>',' ')
        )
    }

    $hm = [regex]::Match($html,'(?is)<h1\b[^>]*>(.*?)</h1>')
    if ($hm.Success) {
        $h1 = [System.Net.WebUtility]::HtmlDecode(
            [regex]::Replace($hm.Groups[1].Value,'(?is)<[^>]+>',' ')
        )
    }

    $info = [pscustomobject]@{
        Html  = $html
        Title = ([regex]::Replace($title,'\s+',' ').Trim())
        H1    = ([regex]::Replace($h1,'\s+',' ').Trim())
    }

    $fileTextCache[$filePath] = $info
    return $info
}

function Find-RegionPage([string]$applyRegion,[string]$parent) {
    $cacheKey = "$applyRegion|$parent"

    if ($regionPageCache.ContainsKey($cacheKey)) {
        return $regionPageCache[$cacheKey]
    }

    if ($explicitPageMap.ContainsKey($cacheKey)) {
        $p = [string]$explicitPageMap[$cacheKey]
        if (Test-Path -LiteralPath $p) {
            $regionPageCache[$cacheKey] = $p
            return $p
        }
    }

    $searchRoot = Get-ApplyRoot $applyRegion
    if ([string]::IsNullOrWhiteSpace($searchRoot) -or !(Test-Path -LiteralPath $searchRoot)) {
        $regionPageCache[$cacheKey] = $null
        return $null
    }

    $tokens = @(Get-ParentTokens $parent $applyRegion)
    $bestPath = $null
    $bestScore = -999999

    foreach ($f in @(Get-RootFiles $searchRoot)) {
        $info = Get-PageInfo $f.FullName
        $score = 0
        $matched = $false

        foreach ($token in $tokens) {
            if ([string]::IsNullOrWhiteSpace($token)) { continue }

            if ($info.H1 -eq $token -or $info.H1 -like "$token *" -or $info.H1 -like "* $token*") {
                $score += 220
                $matched = $true
            }
            elseif ($info.H1 -like "*$token*") {
                $score += 180
                $matched = $true
            }

            if ($info.Title -like "*$token*") {
                $score += 120
                $matched = $true
            }

            if ($f.DirectoryName -like "*$token*") {
                $score += 35
                $matched = $true
            }
        }

        if (-not $matched) { continue }

        if ($info.Title -match "출장마사지") { $score += 20 }
        if ($info.H1 -match "출장마사지") { $score += 20 }

        # 메인 루트 페이지보다 구/시 전용 페이지를 우선
        $depth = ($f.FullName.Substring($searchRoot.Length).TrimStart("\") -split "\\").Count
        $score += [Math]::Min($depth,4)

        if ($score -gt $bestScore) {
            $bestScore = $score
            $bestPath = $f.FullName
        }
    }

    $regionPageCache[$cacheKey] = $bestPath
    return $bestPath
}

function Get-DetailBaseFromRegionPage([string]$regionPage) {
    $dir = Split-Path -Parent $regionPage

    if (!$dir.StartsWith($massageRoot,[System.StringComparison]::OrdinalIgnoreCase)) {
        return $shopsRoot
    }

    $rel = $dir.Substring($massageRoot.Length).TrimStart("\")
    $parts = @($rel -split "\\")

    if ($parts.Count -gt 0) {
        $last = $parts[-1]

        # SEO용 긴 한글 폴더는 제거하고 구조 폴더까지만 사용
        if ($last -match "[가-힣]" -and $last -match "-") {
            if ($parts.Count -gt 1) {
                $parts = @($parts[0..($parts.Count-2)])
            }
        }
    }

    if ($parts.Count -eq 0) { return $shopsRoot }

    return Join-Path $shopsRoot ($parts -join "\")
}

# ----------------------------
# 생성 대상 풍부화 + 검증
# ----------------------------

$enriched = New-Object System.Collections.Generic.List[object]
$unresolved = New-Object System.Collections.Generic.List[object]

$rowNo = 0

foreach ($r in $previewRows) {
    $rowNo++

    $vendor = ([string]$r.업체명).Trim()
    $geo    = ([string]$r.광역지역).Trim()
    $parent = ([string]$r.상위지역).Trim()
    $area   = ([string]$r.하위지역).Trim()

    if (!$shopMap.ContainsKey($vendor)) {
        $unresolved.Add([pscustomobject]@{
            종류="업체기본정보없음"; 업체명=$vendor; 광역지역=$geo; 상위지역=$parent; 하위지역=$area; 적용지역=""; 원인="업체기본정보 시트에 없음"
        })
        continue
    }

    $locKey = "$geo|$parent|$area"
    $loc = $null

    if ($locationMap.ContainsKey($locKey)) {
        $loc = $locationMap[$locKey]
    }

    $apply = if ($loc) { ([string]$loc.적용지역).Trim() } else { "" }

    if ([string]::IsNullOrWhiteSpace($apply)) {
        if ($geo -in @("서울","경기","인천","대전","대구","광주광역시")) {
            $apply = $geo
        }
        elseif ($geo -eq "충북") { $apply = "청주" }
        elseif ($geo -eq "전북") {
            if ($parent -match "전주|덕진구|완산구") { $apply = "전주" }
            elseif ($parent -match "익산") { $apply = "익산" }
        }
        elseif ($geo -eq "경남") {
            if ($parent -match "창원|의창구|성산구|마산|진해구") { $apply = "창원" }
            elseif ($parent -match "김해") { $apply = "김해" }
        }
        elseif ($geo -eq "충남") {
            if ($parent -match "천안") { $apply = "천안" }
            elseif ($parent -match "아산") { $apply = "아산" }
            elseif ($parent -match "공주") { $apply = "공주" }
            elseif ($parent -match "계룡") { $apply = "계룡" }
        }
    }

    if ([string]::IsNullOrWhiteSpace($apply)) {
        $unresolved.Add([pscustomobject]@{
            종류="적용지역확인필요"; 업체명=$vendor; 광역지역=$geo; 상위지역=$parent; 하위지역=$area; 적용지역=""; 원인="적용지역을 결정하지 못함"
        })
        continue
    }

    $regionPage = Find-RegionPage $apply $parent

    if ([string]::IsNullOrWhiteSpace($regionPage)) {
        $unresolved.Add([pscustomobject]@{
            종류="지역페이지없음"; 업체명=$vendor; 광역지역=$geo; 상위지역=$parent; 하위지역=$area; 적용지역=$apply; 원인="연결할 지역 index.html을 찾지 못함"
        })
        continue
    }

    $detailBase = Get-DetailBaseFromRegionPage $regionPage
    $detailDir = Join-Path (Join-Path $detailBase (Safe-PathPart $area)) (Safe-PathPart $vendor)
    $detailFile = Join-Path $detailDir "index.html"

    $detailRel = Get-RelativeToRoot $detailFile
    $detailUrl = "/" + ($detailRel -replace "\\","/")

    $regionRel = Get-RelativeToRoot $regionPage
    $regionUrl = "/" + ($regionRel -replace "\\","/")

    $enriched.Add([pscustomobject]@{
        업체명       = $vendor
        영업시간     = [string]$r.영업시간
        전화번호     = [string]$r.전화번호
        광역지역     = $geo
        적용지역     = $apply
        상위지역     = $parent
        하위지역     = $area
        대표역       = [string]$r.대표역
        대표역주소   = [string]$r.대표역주소
        한줄소개     = [string]$r.한줄소개
        공지사항     = ([string]$r.공지사항).Replace("인근 주변","인근 생활권")
        업체소개     = ([string]$r.업체소개).Replace("인근 주변","인근 생활권")
        지역페이지   = $regionPage
        지역URL      = $regionUrl
        상세파일     = $detailFile
        상세URL      = $detailUrl
    })
}

if ($unresolved.Count -gt 0) {
    $unresolved |
        Export-Csv -LiteralPath $unresolvedCsv -NoTypeInformation -Encoding UTF8

    Write-Host ""
    Write-Host "=== 생성 중단: 지역 연결 확인 필요 ==="
    Write-Host "확인 필요:" $unresolved.Count
    Write-Host "저장:" $unresolvedCsv
    Write-Host ""
    $unresolved |
        Group-Object 종류 |
        ForEach-Object { Write-Host "$($_.Name) : $($_.Count)개" }

    Write-Host ""
    Write-Host "사이트 HTML 수정: 0"
    exit 2
}

if ($enriched.Count -ne $previewRows.Count) {
    throw "생성 대상 개수가 미리보기와 다릅니다. 미리보기=$($previewRows.Count), 연결완료=$($enriched.Count)"
}

# 업체 코스 확인
$vendorsWithoutCourses = @(
    $enriched |
    Select-Object -ExpandProperty 업체명 -Unique |
    Where-Object { !$courseMap.ContainsKey($_) -or @($courseMap[$_]).Count -eq 0 }
)

if ($vendorsWithoutCourses.Count -gt 0) {
    Write-Host ""
    Write-Host "코스가 없는 업체가 있어 생성하지 않습니다:"
    $vendorsWithoutCourses | ForEach-Object { Write-Host " - $_" }
    Write-Host "사이트 HTML 수정: 0"
    exit 3
}

Write-Host "지역페이지 연결:" (@($enriched | Select-Object -ExpandProperty 지역페이지 -Unique).Count)
Write-Host "상세페이지 연결:" $enriched.Count
Write-Host "검증 오류: 0"

# ----------------------------
# 상세페이지 그룹/내부링크 준비
# ----------------------------

$vendorRegionGroups = @{}
foreach ($g in ($enriched | Group-Object 업체명,지역페이지)) {
    $vendorRegionGroups[$g.Name] = @($g.Group | Sort-Object 하위지역)
}

function Get-SiblingLinks($row) {
    $groupName = "$($row.업체명), $($row.지역페이지)"
    if (!$vendorRegionGroups.ContainsKey($groupName)) { return "" }

    $siblings = @(
        $vendorRegionGroups[$groupName] |
        Where-Object { $_.상세URL -ne $row.상세URL } |
        Select-Object -First 8
    )

    if ($siblings.Count -eq 0) { return "" }

    $items = foreach ($s in $siblings) {
        '<a href="' + (HtmlEncode $s.상세URL) + '">' + (HtmlEncode $s.하위지역) + '</a>'
    }

    return @"
<section class="ml-nearby">
  <h2>같은 업체 이용 가능 지역</h2>
  <div class="ml-nearby-links">
    $($items -join "`r`n    ")
  </div>
</section>
"@
}

function Build-CourseHtml([string]$vendor) {
    $list = @($courseMap[$vendor])

    $cards = foreach ($c in $list) {
        $desc = ""
        if (-not [string]::IsNullOrWhiteSpace([string]$c.코스설명)) {
            $desc = '<p class="ml-course-desc">' + (HtmlEncode $c.코스설명) + '</p>'
        }

@"
<div class="ml-course-card">
  <div>
    <h3>$(HtmlEncode $c.코스명)</h3>
    <p class="ml-course-time">$(HtmlEncode $c.시간)</p>
    $desc
  </div>
  <strong>$(HtmlEncode $c.가격)</strong>
</div>
"@
    }

    return ($cards -join "`r`n")
}

function Build-DetailHtml($row) {
    $shop = $shopMap[$row.업체명]
    $phone = [string]$shop.전화번호
    if ([string]::IsNullOrWhiteSpace($phone)) { $phone = [string]$row.전화번호 }

    $hours = [string]$shop.영업시간
    if ([string]::IsNullOrWhiteSpace($hours)) { $hours = [string]$row.영업시간 }

    $phoneLink = Get-PhoneLink $phone
    $smsLink = "sms:$phoneLink?body=$([uri]::EscapeDataString($smsText))"

    $stationLine = ""
    if (-not [string]::IsNullOrWhiteSpace([string]$row.대표역)) {
        $stationLine = "$(HtmlEncode $row.대표역) 인근"
    }
    else {
        $stationLine = "$(HtmlEncode $row.하위지역) 중심"
    }

    $addressLine = ""
    if (-not [string]::IsNullOrWhiteSpace([string]$row.대표역주소)) {
        $addressLine = '<p class="ml-location-address">' + (HtmlEncode $row.대표역주소) + '</p>'
    }

    $courseHtml = Build-CourseHtml $row.업체명
    $siblingHtml = Get-SiblingLinks $row

    $faqHours = if ($hours -eq "24시간") {
        "24시간, 휴일에도 가능합니다."
    } else {
        "영업시간은 $hours 이며, 야간·휴일 가능 여부는 이용 전 문의해 주세요."
    }

    $title = "$($row.업체명) | $($row.상위지역) $($row.하위지역) 이용 안내"
    $meta = "$($row.상위지역) $($row.하위지역)에서 이용 가능한 $($row.업체명) 안내입니다. 영업시간, 코스와 가격, 전화·문자 예약, 주변 위치 정보를 확인할 수 있습니다."

@"
<!doctype html>
<html lang="ko">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <title>$(HtmlEncode $title)</title>
  <meta name="description" content="$(HtmlEncode $meta)">
  <meta property="og:title" content="$(HtmlEncode $title)">
  <meta property="og:description" content="$(HtmlEncode $meta)">
  <meta property="og:type" content="website">
  <link rel="icon" href="/favicon.ico" sizes="any">
  <style>
    :root{--green:#173f38;--green2:#214f46;--pink:#ee8ca7;--paper:#f7f5f1;--ink:#202624;--muted:#6c7471;--line:#e7e2db;--gold:#f3d28b}
    *{box-sizing:border-box}
    html{scroll-behavior:smooth}
    body{margin:0;background:var(--paper);color:var(--ink);font-family:Arial,"Noto Sans KR","Apple SD Gothic Neo",sans-serif;line-height:1.65;padding-bottom:86px}
    a{text-decoration:none;color:inherit}
    .ml-wrap{width:min(960px,calc(100% - 28px));margin:0 auto}
    .ml-hero{background:linear-gradient(145deg,var(--green),#0f302b);color:white;padding:46px 0 32px}
    .ml-back{display:inline-flex;align-items:center;gap:6px;color:#d8e5e2;font-size:14px;margin-bottom:20px}
    .ml-kicker{margin:0 0 6px;color:#d9ece7;font-size:14px}
    .ml-hero h1{margin:0;color:var(--gold);font-size:clamp(30px,6vw,48px);letter-spacing:-1px}
    .ml-hero-copy{margin:12px 0 0;color:#eef5f3;font-size:16px}
    .ml-location{margin-top:22px;padding:15px 16px;border:1px solid rgba(255,255,255,.16);border-radius:14px;background:rgba(255,255,255,.06)}
    .ml-location strong{display:block;color:white}
    .ml-location-address{margin:3px 0 0;color:#dce8e5;font-size:13px}
    .ml-band{background:var(--pink);color:#fff;padding:18px 0}
    .ml-band-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;text-align:center}
    .ml-band-item{padding:4px}
    .ml-band-item span{display:block;font-size:12px;opacity:.9}
    .ml-band-item strong{display:block;font-size:16px}
    main{padding:30px 0 20px}
    .ml-section{margin:0 0 28px}
    .ml-section h2,.ml-nearby h2{font-size:22px;margin:0 0 14px;color:#193d36}
    .ml-card{background:white;border:1px solid var(--line);border-radius:18px;padding:20px;box-shadow:0 8px 24px rgba(30,45,40,.05)}
    .ml-course-list{display:grid;gap:12px}
    .ml-course-card{display:flex;justify-content:space-between;gap:18px;align-items:center;background:white;border:1px solid var(--line);border-radius:16px;padding:17px 18px}
    .ml-course-card h3{margin:0 0 2px;font-size:17px}
    .ml-course-card strong{white-space:nowrap;color:#b45f79;font-size:18px}
    .ml-course-time{margin:0;color:var(--muted);font-size:14px}
    .ml-course-desc{margin:5px 0 0;color:#707875;font-size:13px}
    .ml-info-grid{display:grid;grid-template-columns:1fr 1fr;gap:14px}
    .ml-info-grid .ml-card h3{margin:0 0 8px;font-size:15px;color:#6c7471}
    .ml-info-grid .ml-card p{margin:0;font-weight:700}
    .ml-notice-box{background:#fff9ed;border:1px solid #f0d49a;border-radius:18px;padding:20px}
    .ml-notice-box h2{margin:0 0 8px;color:#705021;font-size:20px}
    .ml-notice-box ul{margin:9px 0 0;padding-left:21px}
    .ml-faq-section{margin-top:30px}
    .ml-faq-item{background:white;border:1px solid var(--line);border-radius:15px;padding:17px 18px;margin-bottom:10px}
    .ml-faq-item h3{margin:0 0 6px;font-size:16px;color:#193d36}
    .ml-faq-item p{margin:0;color:#59625f}
    .ml-nearby{margin-top:30px}
    .ml-nearby-links{display:flex;flex-wrap:wrap;gap:8px}
    .ml-nearby-links a{background:white;border:1px solid var(--line);border-radius:999px;padding:8px 13px;font-size:13px}
    .ml-actions{display:flex;gap:10px;margin-top:22px}
    .ml-btn{flex:1;text-align:center;border-radius:13px;padding:13px 15px;font-weight:800}
    .ml-btn-call{background:#fff;color:#173f38}
    .ml-btn-sms{background:#f2a1b7;color:#532536}
    .ml-sticky-call{position:fixed;left:0;right:0;bottom:0;z-index:50;background:rgba(18,42,37,.96);backdrop-filter:blur(8px);padding:10px 12px calc(10px + env(safe-area-inset-bottom));box-shadow:0 -6px 24px rgba(0,0,0,.14)}
    .ml-sticky-inner{width:min(960px,100%);margin:0 auto;display:grid;grid-template-columns:1fr 1fr;gap:9px}
    .ml-sticky-inner a{display:block;text-align:center;border-radius:12px;padding:12px;font-weight:900}
    .ml-sticky-inner .call{background:white;color:#173f38}
    .ml-sticky-inner .sms{background:#ee8ca7;color:white}
    @media(max-width:640px){
      .ml-band-grid{grid-template-columns:1fr}
      .ml-info-grid{grid-template-columns:1fr}
      .ml-course-card{align-items:flex-start}
      .ml-hero{padding-top:32px}
    }
  </style>
</head>
<body>
<!-- ML-GENERATED-SHOP-PAGE -->
<header class="ml-hero">
  <div class="ml-wrap">
    <a class="ml-back" href="$(HtmlEncode $row.지역URL)">← $(HtmlEncode $row.상위지역) 지역 안내</a>
    <p class="ml-kicker">$(HtmlEncode $row.광역지역) · $(HtmlEncode $row.상위지역) · $(HtmlEncode $row.하위지역)</p>
    <h1>$(HtmlEncode $row.업체명)</h1>
    <p class="ml-hero-copy">$(HtmlEncode $row.한줄소개)</p>

    <div class="ml-location">
      <strong>$stationLine</strong>
      $addressLine
    </div>

    <div class="ml-actions">
      <a class="ml-btn ml-btn-call" href="tel:$(HtmlEncode $phoneLink)">전화 문의</a>
      <a class="ml-btn ml-btn-sms" href="$(HtmlEncode $smsLink)">문자 문의</a>
    </div>
  </div>
</header>

<div class="ml-band">
  <div class="ml-wrap ml-band-grid">
    <div class="ml-band-item"><span>업체</span><strong>$(HtmlEncode $row.업체명)</strong></div>
    <div class="ml-band-item"><span>영업시간</span><strong>$(HtmlEncode $hours)</strong></div>
    <div class="ml-band-item"><span>이용지역</span><strong>$(HtmlEncode $row.하위지역)</strong></div>
  </div>
</div>

<main class="ml-wrap">
  <section class="ml-section">
    <h2>코스 안내</h2>
    <div class="ml-course-list">
      $courseHtml
    </div>
  </section>

  <section class="ml-section">
    <h2>업체 정보</h2>
    <div class="ml-info-grid">
      <div class="ml-card">
        <h3>영업시간</h3>
        <p>$(HtmlEncode $hours)</p>
      </div>
      <div class="ml-card">
        <h3>전화번호</h3>
        <p>$(HtmlEncode $phone)</p>
      </div>
    </div>
  </section>

  <section class="ml-section">
    <h2>공지사항</h2>
    <div class="ml-card">
      <p>$(HtmlEncode $row.공지사항)</p>
    </div>
  </section>

  <section class="ml-section">
    <h2>업체소개</h2>
    <div class="ml-card">
      <p>$(HtmlEncode $row.업체소개)</p>
    </div>
  </section>

  <section class="ml-notice-box">
    <h2>이용 전 확인해 주세요</h2>
    <ul>
      <li>예약 가능 여부는 전화 또는 문자로 먼저 확인해 주세요.</li>
      <li>당일 일정과 이동 상황에 따라 가능한 시간이 달라질 수 있습니다.</li>
      <li>정확한 이용 가능 지역과 자세한 안내는 상담 시 확인해 주세요.</li>
    </ul>
  </section>

  <section class="ml-faq-section">
    <h2>자주 묻는 질문</h2>
    <div class="ml-faq-item">
      <h3>야간이나 새벽에도 가능한가요?</h3>
      <p>$(HtmlEncode $faqHours)</p>
    </div>
    <div class="ml-faq-item">
      <h3>관리는 어떻게 이용하나요?</h3>
      <p>이용 전 문의하시면 됩니다.</p>
    </div>
    <div class="ml-faq-item">
      <h3>준비해야 할 게 있나요?</h3>
      <p>편안하게 이용받을 수 있는 공간이면 됩니다.</p>
    </div>
  </section>

  $siblingHtml
</main>

<div class="ml-sticky-call">
  <div class="ml-sticky-inner">
    <a class="call" href="tel:$(HtmlEncode $phoneLink)">전화하기</a>
    <a class="sms" href="$(HtmlEncode $smsLink)">문자하기</a>
  </div>
</div>
</body>
</html>
"@
}

# ----------------------------
# 실제 생성 시작
# ----------------------------

Write-Host ""
Write-Host "=== 4. 실제 HTML 자동생성 시작 ==="
Write-Host "백업 폴더:" $backupRoot

New-Item -ItemType Directory -Force -Path $backupRoot | Out-Null

$detailCreated = 0
$detailUpdated = 0
$detailFailed = 0
$report = New-Object System.Collections.Generic.List[object]

$detailNo = 0

foreach ($row in $enriched) {
    $detailNo++

    $file = [string]$row.상세파일
    $dir = Split-Path -Parent $file

    New-Item -ItemType Directory -Force -Path $dir | Out-Null

    $wasExisting = Test-Path -LiteralPath $file
    if ($wasExisting) {
        Backup-File $file
    }

    try {
        $html = Build-DetailHtml $row
        Set-Content -LiteralPath $file -Value $html -Encoding UTF8

        if ($wasExisting) { $detailUpdated++ }
        else { $detailCreated++ }

        $report.Add([pscustomobject]@{
            업체명=$row.업체명
            적용지역=$row.적용지역
            광역지역=$row.광역지역
            상위지역=$row.상위지역
            하위지역=$row.하위지역
            지역페이지=$row.지역페이지
            상세페이지=$file
            상태=if($wasExisting){"업데이트"}else{"신규"}
        })
    }
    catch {
        $detailFailed++

        $report.Add([pscustomobject]@{
            업체명=$row.업체명
            적용지역=$row.적용지역
            광역지역=$row.광역지역
            상위지역=$row.상위지역
            하위지역=$row.하위지역
            지역페이지=$row.지역페이지
            상세페이지=$file
            상태="실패: $($_.Exception.Message)"
        })
    }

    if (($detailNo % 250) -eq 0 -or $detailNo -eq $enriched.Count) {
        Write-Host "상세페이지 진행: $detailNo / $($enriched.Count)"
    }
}

# ----------------------------
# 지역페이지 업체 카드 생성
# ----------------------------

function Get-FirstCourseSummary([string]$vendor) {
    $list = @($courseMap[$vendor])
    if ($list.Count -eq 0) { return "" }

    $c = $list[0]
    return "$($c.코스명) $($c.시간) $($c.가격)"
}

function Build-RegionCardBlock($rowsForPage) {
    $parent = [string]($rowsForPage | Select-Object -First 1).상위지역
    $groups = @($rowsForPage | Group-Object 업체명 | Sort-Object Name)
    $cards = New-Object System.Collections.Generic.List[string]

    foreach ($g in $groups) {
        $vendorRows = @($g.Group | Sort-Object 하위지역)
        $rep = $vendorRows | Select-Object -First 1
        $areas = @($vendorRows | Select-Object -ExpandProperty 하위지역 -Unique)

        $areaText = if ($areas.Count -le 3) {
            ($areas -join " · ")
        } else {
            "$($areas[0]) · $($areas[1]) · $($areas[2]) 외 $($areas.Count - 3)개 지역"
        }

        $cards.Add(@"
<a class="ml-top-shop-card" href="$(HtmlEncode $rep.상세URL)">
  <div class="ml-top-shop-copy">$(HtmlEncode $areaText) 중심 빠른 예약과 편안한 방문 관리</div>
  <div class="ml-top-shop-line"></div>
  <div class="ml-top-shop-name">$(HtmlEncode $g.Name)</div>
</a>
"@)
    }

@"
<!-- ML-TOP-SHOP-START -->
<style id="ml-top-shop-style">
.ml-top-shop-section{margin:28px 0 34px}
.ml-top-shop-head{display:flex;align-items:flex-end;justify-content:space-between;gap:16px;margin-bottom:14px}
.ml-top-shop-kicker{margin:0 0 4px;color:#9b7b68;font-size:12px;letter-spacing:.16em}
.ml-top-shop-title{margin:0;font-size:30px;line-height:1.2;letter-spacing:-.7px}
.ml-top-shop-count{white-space:nowrap;color:#7d6d65;font-size:13px;font-weight:700}
.ml-top-shop-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:14px}
.ml-top-shop-card{display:flex;min-height:190px;flex-direction:column;justify-content:center;align-items:center;text-align:center;background:#083f35;border-radius:15px;padding:26px 34px;text-decoration:none;box-shadow:0 7px 20px rgba(11,58,49,.08)}
.ml-top-shop-copy{color:#b8e0d6;font-size:13px;line-height:1.55}
.ml-top-shop-line{width:88%;height:1px;background:#d8b23b;margin:20px 0 12px}
.ml-top-shop-name{color:#ffc928;font-size:17px;font-weight:900}
.ml-top-shop-card:hover{transform:translateY(-1px)}
@media(max-width:700px){.ml-top-shop-grid{grid-template-columns:1fr}.ml-top-shop-title{font-size:26px}.ml-top-shop-card{min-height:170px}}
</style>
<section class="ml-top-shop-section" aria-label="등록 업체">
  <div class="ml-top-shop-head">
    <div>
      <p class="ml-top-shop-kicker">REGISTERED SHOP</p>
      <h2 class="ml-top-shop-title">$(HtmlEncode $parent) 등록 업체</h2>
    </div>
    <div class="ml-top-shop-count">$($groups.Count)개 업체</div>
  </div>
  <div class="ml-top-shop-grid">
    $($cards -join "`r`n    ")
  </div>
</section>
<!-- ML-TOP-SHOP-END -->
"@
}
$rowsByRegionPage = @{}
foreach ($g in ($enriched | Group-Object 지역페이지)) {
    $rowsByRegionPage[$g.Name] = @($g.Group)
}

# 기존 생성 카드가 있으나 현재 업체가 없는 페이지도 정리 대상에 포함
$pagesWithOldCards = @(
    Get-ChildItem -LiteralPath $massageRoot -Recurse -File -Filter "index.html" |
    Where-Object {
        try {
            $t = Get-Content -LiteralPath $_.FullName -Raw -Encoding UTF8
            return ($t -match "<!-- ML-SHOP-CARDS-START -->" -or $t -match "<!-- SHOP-TEST-V68-START -->")
        } catch { return $false }
    }
)

$allRegionPages = New-Object System.Collections.Generic.HashSet[string]([System.StringComparer]::OrdinalIgnoreCase)

foreach ($p in $rowsByRegionPage.Keys) {
    [void]$allRegionPages.Add($p)
}

foreach ($f in $pagesWithOldCards) {
    [void]$allRegionPages.Add($f.FullName)
}

$regionModified = 0
$regionNo = 0

foreach ($page in $allRegionPages) {
    $regionNo++

    if (!(Test-Path -LiteralPath $page)) { continue }

    $html = Get-Content -LiteralPath $page -Raw -Encoding UTF8
    $original = $html

    # 이전 테스트 카드 제거
    $html = [regex]::Replace(
        $html,
        '(?is)\s*<!-- SHOP-TEST-V68-START -->.*?<!-- SHOP-TEST-V68-END -->\s*',
        "`r`n"
    )

    # 기존 자동생성 카드 제거
    $html = [regex]::Replace(
        $html,
        '(?is)\s*<!-- ML-SHOP-CARDS-START -->.*?<!-- ML-SHOP-CARDS-END -->\s*',
        "`r`n"
    )

    # 예전 하단 자동카드와 현재 상단 자동카드 제거
    $html = [regex]::Replace(
        $html,
        '(?is)\s*<!-- ML-TOP-SHOP-START -->.*?<!-- ML-TOP-SHOP-END -->\s*',
        "`r`n"
    )

    if ($rowsByRegionPage.ContainsKey($page)) {
        $block = Build-RegionCardBlock $rowsByRegionPage[$page]

        # 기존 REGISTERED SHOP 영역을 상단 자동업체 영역으로 교체
        $registeredPattern = '(?is)<section\b[^>]*>.*?REGISTERED SHOP.*?</section>'
        $registeredMatch = [regex]::Match($html,$registeredPattern)

        if ($registeredMatch.Success) {
            $html = $html.Substring(0,$registeredMatch.Index) + $block + $html.Substring($registeredMatch.Index + $registeredMatch.Length)
        }
        else {
            # 기존 섹션이 없으면 main의 가장 위에 삽입
            $mainMatch = [regex]::Match($html,'(?is)<main\b[^>]*>')
            if ($mainMatch.Success) {
                $insertAt = $mainMatch.Index + $mainMatch.Length
                $html = $html.Substring(0,$insertAt) + "`r`n$block`r`n" + $html.Substring($insertAt)
            }
            elseif ($html -match '(?is)</body>') {
                $html = [regex]::Replace($html,'(?is)</body>',"$block`r`n</body>",1)
            }
        }
    }

    if ($html -ne $original) {
        Backup-File $page
        Set-Content -LiteralPath $page -Value $html -Encoding UTF8
        $regionModified++
    }

    if (($regionNo % 25) -eq 0 -or $regionNo -eq $allRegionPages.Count) {
        Write-Host "지역페이지 진행: $regionNo / $($allRegionPages.Count)"
    }
}

# ----------------------------
# 실행 리포트 / 검증
# ----------------------------

$report | Export-Csv -LiteralPath $reportCsv -NoTypeInformation -Encoding UTF8

$generatedFiles = @(
    $enriched |
    Where-Object { Test-Path -LiteralPath $_.상세파일 }
)

$generatedMarkerOk = 0
$smsOk = 0
$phoneOk = 0
$faqOk = 0

foreach ($r in $generatedFiles) {
    $t = Get-Content -LiteralPath $r.상세파일 -Raw -Encoding UTF8

    if ($t -match "<!-- ML-GENERATED-SHOP-PAGE -->") { $generatedMarkerOk++ }
    if ($t -like "*$smsText*") { $smsOk++ }
    if ($t -match "전화하기") { $phoneOk++ }
    if ($t -match "자주 묻는 질문") { $faqOk++ }
}

$regionCardPagesOk = 0
foreach ($page in $rowsByRegionPage.Keys) {
    if (!(Test-Path -LiteralPath $page)) { continue }
    $t = Get-Content -LiteralPath $page -Raw -Encoding UTF8
    if ($t -match "<!-- ML-SHOP-CARDS-START -->") { $regionCardPagesOk++ }
}

Write-Host ""
Write-Host "=== 전체 업체 자동생성 완료 ==="
Write-Host "업체 수:" (@($enriched | Select-Object -ExpandProperty 업체명 -Unique).Count)
Write-Host "상세페이지 목표:" $enriched.Count
Write-Host "상세페이지 신규:" $detailCreated
Write-Host "상세페이지 업데이트:" $detailUpdated
Write-Host "상세페이지 실패:" $detailFailed
Write-Host ""
Write-Host "지역페이지 카드 대상:" $rowsByRegionPage.Count
Write-Host "지역페이지 실제 수정:" $regionModified
Write-Host "카드 블록 확인:" $regionCardPagesOk
Write-Host ""
Write-Host "상세페이지 마커 확인:" "$generatedMarkerOk / $($enriched.Count)"
Write-Host "문자문구 확인:" "$smsOk / $($enriched.Count)"
Write-Host "하단 전화버튼 확인:" "$phoneOk / $($enriched.Count)"
Write-Host "FAQ 확인:" "$faqOk / $($enriched.Count)"
Write-Host ""
Write-Host "백업 폴더:" $backupRoot
Write-Host "실행 리포트:" $reportCsv
Write-Host ""
Write-Host "사이트 HTML 수정:" ($detailCreated + $detailUpdated + $regionModified)
Write-Host ""
Write-Host "※ 사용=N 또는 적용지역에서 제외된 기존 생성 상세페이지는 자동 삭제하지 않습니다."


# ML-DETAIL-TEMPLATE-HOOK
$mlDetailTemplateScript = Join-Path $dataDir "mong-lounge-apply-detail-template.ps1"
if (Test-Path -LiteralPath $mlDetailTemplateScript) {
    Write-Host ""
    Write-Host "=== 기준 상세페이지 템플릿 재적용 ==="
    & $mlDetailTemplateScript -FromGenerator
}
