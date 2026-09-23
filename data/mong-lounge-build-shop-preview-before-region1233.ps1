$ErrorActionPreference = "Stop"

$root = "E:\mong-lounge-site"
$xlsx = Join-Path $root "data\몽라운지_업체자동등록_읽기용.xlsx"
$locFile = Join-Path $root "data\shop-region-location-final-v2.csv"
$outFile = Join-Path $root "data\shop-generation-preview.csv"

if (!(Test-Path -LiteralPath $xlsx)) {
    throw "엑셀 파일 없음: $xlsx"
}
if (!(Test-Path -LiteralPath $locFile)) {
    throw "지역 마스터 파일 없음: $locFile"
}

Add-Type -AssemblyName System.IO.Compression.FileSystem

function Read-ZipText($zip, [string]$name) {
    $entry = $zip.Entries | Where-Object FullName -eq $name | Select-Object -First 1
    if ($null -eq $entry) { return $null }

    $sr = New-Object System.IO.StreamReader($entry.Open())
    try { return $sr.ReadToEnd() }
    finally { $sr.Close() }
}

function Get-CellValue($cell, $sharedStrings) {
    $v = [string]$cell.v
    $t = [string]$cell.t

    if ($t -eq "s" -and $v -ne "") {
        return [string]$sharedStrings[[int]$v]
    }

    if ($t -eq "inlineStr") {
        if ($cell.is.t) {
            return [string]$cell.is.t
        }
        if ($cell.is.r) {
            return (($cell.is.r | ForEach-Object { [string]$_.t }) -join "")
        }
    }

    return $v
}

function Read-SheetRows($zip, [string]$sheetPath, $sharedStrings) {
    $text = Read-ZipText $zip $sheetPath
    if ([string]::IsNullOrWhiteSpace($text)) {
        throw "시트를 읽지 못했습니다: $sheetPath"
    }

    [xml]$xml = $text
    $result = @()

    foreach ($row in $xml.worksheet.sheetData.row) {
        if ([int]$row.r -le 1) { continue }

        $h = @{}

        foreach ($cell in $row.c) {
            $col = ([regex]::Match([string]$cell.r, '^[A-Z]+')).Value
            $h[$col] = Get-CellValue $cell $sharedStrings
        }

        $result += ,$h
    }

    return $result
}

function Format-Price([string]$value) {
    if ([string]::IsNullOrWhiteSpace($value)) { return "" }

    $digits = $value -replace '[^\d]',''
    if ($digits) {
        $n = 0L
        if ([long]::TryParse($digits, [ref]$n)) {
            return ("{0:N0}원" -f $n)
        }
    }

    return $value
}

function Matches-ApplyRegion($loc, [string]$region) {
    $region = $region.Trim()
    if ([string]::IsNullOrWhiteSpace($region)) { return $false }

    if ($region -in @("서울","경기","인천")) {
        return ([string]$loc.광역지역 -eq $region)
    }

    if ($region -in @("천안","아산")) {
        return (
            [string]$loc.적용지역 -eq $region -or
            [string]$loc.상위지역 -eq $region
        )
    }

    $parent = [string]$loc.상위지역

    return (
        $parent -eq $region -or
        $parent -like "$region-*" -or
        $parent -like "*-$region"
    )
}

# -----------------------------
# 엑셀 직접 읽기
# -----------------------------
$zip = [System.IO.Compression.ZipFile]::OpenRead($xlsx)

try {
    $sharedStrings = @()

    $ssText = Read-ZipText $zip "xl/sharedStrings.xml"

    if ($ssText) {
        [xml]$ssXml = $ssText

        foreach ($si in $ssXml.sst.si) {
            if ($si.t) {
                $sharedStrings += [string]$si.t
            }
            elseif ($si.r) {
                $sharedStrings += (($si.r | ForEach-Object { [string]$_.t }) -join "")
            }
            else {
                $sharedStrings += ""
            }
        }
    }

    # v2 파일 시트 순서:
    # sheet2 = 업체기본정보
    # sheet3 = 코스정보
    # sheet4 = 적용지역
    $shopRows   = @(Read-SheetRows $zip "xl/worksheets/sheet2.xml" $sharedStrings)
    $courseRows = @(Read-SheetRows $zip "xl/worksheets/sheet3.xml" $sharedStrings)
    $applyRows  = @(Read-SheetRows $zip "xl/worksheets/sheet4.xml" $sharedStrings)
}
finally {
    $zip.Dispose()
}

$shops = @(
    $shopRows |
    Where-Object {
        $_.A -ne "N" -and
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
    $courseRows |
    Where-Object {
        $_.A -ne "N" -and
        -not [string]::IsNullOrWhiteSpace([string]$_.B) -and
        -not [string]::IsNullOrWhiteSpace([string]$_.C)
    } |
    ForEach-Object {
        $order = 9999
        if ($_.G -match '^\d+$') { $order = [int]$_.G }

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

$applies = @(
    $applyRows |
    Where-Object {
        $_.A -ne "N" -and
        -not [string]::IsNullOrWhiteSpace([string]$_.B) -and
        -not [string]::IsNullOrWhiteSpace([string]$_.C)
    } |
    ForEach-Object {
        [pscustomobject]@{
            업체명   = ([string]$_.B).Trim()
            적용지역 = ([string]$_.C).Trim()
        }
    }
)

$locations = @(Import-Csv -LiteralPath $locFile)

Write-Host ""
Write-Host "=== 업체 생성 미리보기 준비 ==="
Write-Host "업체:" $shops.Count
Write-Host "코스:" $courses.Count
Write-Host "적용지역 설정:" $applies.Count
Write-Host "지역 마스터:" $locations.Count

$result = @()
$rowNo = 0

foreach ($shop in $shops) {

    $shopApplies = @(
        $applies | Where-Object 업체명 -eq $shop.업체명
    )

    if ($shopApplies.Count -eq 0) {
        Write-Host "주의: $($shop.업체명) 적용지역이 없습니다."
        continue
    }

    $shopCourses = @(
        $courses |
        Where-Object 업체명 -eq $shop.업체명 |
        Sort-Object 정렬순서,코스명
    )

    $courseSummary = (
        $shopCourses |
        ForEach-Object {
            $parts = @($_.코스명)

            if (-not [string]::IsNullOrWhiteSpace($_.시간)) {
                $parts += $_.시간
            }
            if (-not [string]::IsNullOrWhiteSpace($_.가격)) {
                $parts += $_.가격
            }

            $parts -join " "
        }
    ) -join " / "

    $selected = @()

    foreach ($a in $shopApplies) {
        $selected += @(
            $locations | Where-Object {
                Matches-ApplyRegion $_ $a.적용지역
            }
        )
    }

    # 같은 지역이 여러 적용범위에 겹쳐도 한 번만 생성
    $selected = @(
        $selected |
        Sort-Object 광역지역,상위지역,하위지역 -Unique
    )

    foreach ($loc in $selected) {

        $rowNo++

        $parentText = ([string]$loc.상위지역).Replace("-"," ")
        $areaText = ("{0} {1}" -f $parentText,[string]$loc.하위지역).Trim()

        $stationText = if (-not [string]::IsNullOrWhiteSpace([string]$loc.대표역)) {
            "$($loc.대표역) 인근"
        }
        else {
            "$($loc.하위지역) 중심"
        }

        # 여러 문형을 돌려가며 사용
        $variant = $rowNo % 4

        switch ($variant) {
            0 {
                $intro = "$areaText $stationText 전화·문자 예약 안내"
                $notice = "예약은 전화 또는 문자로 접수합니다. $areaText 지역의 예약 가능 여부와 방문 일정은 연락 시 확인해 주세요. 운영시간은 $($shop.영업시간)이며 현장 상황에 따라 시간이 조정될 수 있습니다."
                $about = "$($shop.업체명)은(는) $areaText 이용 고객을 위한 지역 안내 페이지입니다. $stationText 생활권을 기준으로 예약 정보를 확인할 수 있으며, 코스와 이용 시간은 예약 전에 전화 또는 문자로 안내받을 수 있습니다."
            }
            1 {
                $intro = "$stationText $($shop.업체명) 예약 및 코스 안내"
                $notice = "$areaText 예약 문의는 전화와 문자로 가능합니다. $($shop.영업시간) 운영 기준이며 원하는 시간대의 가능 여부는 예약 시 확인해 주세요. 지역별 이동 일정에 따라 방문 순서가 달라질 수 있습니다."
                $about = "$areaText 지역에서 $($shop.업체명) 이용 정보를 확인할 수 있도록 구성한 안내입니다. $stationText 기준의 위치 정보와 코스 구성을 함께 제공하며 실제 예약 일정은 상담 후 확정됩니다."
            }
            2 {
                $intro = "$areaText 중심 $($shop.업체명) 이용 안내"
                $notice = "$areaText 지역 예약은 전화 또는 문자 문의 후 진행됩니다. 운영시간은 $($shop.영업시간)이며 당일 일정이나 이동 상황에 따라 가능한 시간이 달라질 수 있으므로 먼저 예약 가능 여부를 확인해 주세요."
                $about = "$($shop.업체명)의 $areaText 지역 안내입니다. $stationText 주변을 포함한 해당 생활권의 예약 정보와 코스를 한 페이지에서 확인할 수 있도록 구성하며, 자세한 이용 내용은 상담 시 안내합니다."
            }
            default {
                $intro = "$stationText 중심 빠른 예약 확인과 코스 안내"
                $notice = "전화 또는 문자로 $areaText 지역의 예약 가능 여부를 확인할 수 있습니다. $($shop.영업시간) 운영 기준으로 접수하며, 방문 시간은 당일 예약 현황과 이동 상황을 확인한 뒤 안내합니다."
                $about = "$areaText 생활권에서 이용할 수 있는 $($shop.업체명) 지역 페이지입니다. $stationText 위치 정보를 참고할 수 있고, 등록된 코스와 시간·가격을 확인한 뒤 전화나 문자로 예약 문의할 수 있습니다."
            }
        }

        $result += [pscustomobject]@{
            업체명       = $shop.업체명
            영업시간     = $shop.영업시간
            전화번호     = $shop.전화번호
            문자번호     = $shop.전화번호
            문자문구     = "건마천사보고 연락드립니다 예약 가능 한가요?"

            광역지역     = $loc.광역지역
            상위지역     = $loc.상위지역
            하위지역     = $loc.하위지역

            대표역       = $loc.대표역
            대표역주소   = $loc.대표역주소
            위치표시     = $loc.위치표시
            주소표시     = $loc.주소표시
            역선정방식   = $loc.역선정방식

            한줄소개     = $intro
            공지사항     = $notice
            업체소개     = $about

            코스개수     = $shopCourses.Count
            코스요약     = $courseSummary
        }
    }
}

$result |
    Export-Csv -LiteralPath $outFile -NoTypeInformation -Encoding UTF8

Write-Host ""
Write-Host "=== 업체 생성 미리보기 결과 ==="
Write-Host "저장:" $outFile
Write-Host "생성 예정 페이지:" $result.Count

foreach ($g in ($result | Group-Object 업체명)) {
    Write-Host "$($g.Name): $($g.Count)"
}

Write-Host ""
Write-Host "=== 지역별 개수 ==="
foreach ($g in ($result | Group-Object 광역지역)) {
    Write-Host "$($g.Name): $($g.Count)"
}

Write-Host ""
Write-Host "=== 역삼동 샘플 ==="
$result |
    Where-Object 하위지역 -eq "역삼동" |
    Select-Object -First 1 |
    Format-List 업체명,영업시간,전화번호,광역지역,상위지역,하위지역,대표역,대표역주소,한줄소개,공지사항,업체소개,코스개수,코스요약

Write-Host ""
Write-Host "사이트 HTML 수정: 0"
Write-Host "업체 생성 미리보기 완료"


