$ErrorActionPreference = "Stop"

$root = "E:\mong-lounge-site"
$src  = Join-Path $root "data\몽라운지_업체자동등록_마스터_v2.xlsx"
$tmp  = Join-Path $root "data\몽라운지_업체자동등록_진단용.xlsx"

if (!(Test-Path -LiteralPath $src)) {
    throw "마스터 엑셀을 찾을 수 없습니다: $src"
}

$f1 = [System.IO.File]::Open($src,[System.IO.FileMode]::Open,[System.IO.FileAccess]::Read,[System.IO.FileShare]::ReadWrite)
$f2 = [System.IO.File]::Create($tmp)
try { $f1.CopyTo($f2) } finally { $f2.Close(); $f1.Close() }

Add-Type -AssemblyName System.IO.Compression.FileSystem

function Read-ZipText($zip, [string]$name) {
    $entry = $zip.Entries | Where-Object FullName -eq $name | Select-Object -First 1
    if ($null -eq $entry) { return $null }
    $sr = New-Object System.IO.StreamReader($entry.Open())
    try { return $sr.ReadToEnd() } finally { $sr.Close() }
}

function Get-CellValue($cell, $sharedStrings) {
    $v = [string]$cell.v
    $t = [string]$cell.t
    if ($t -eq "s" -and $v -ne "") { return [string]$sharedStrings[[int]$v] }
    if ($t -eq "inlineStr") {
        if ($cell.is.t) { return [string]$cell.is.t }
        if ($cell.is.r) { return (($cell.is.r | ForEach-Object { [string]$_.t }) -join "") }
    }
    return $v
}

function Read-SheetRows($zip, [string]$sheetPath, $sharedStrings) {
    $text = Read-ZipText $zip $sheetPath
    if ([string]::IsNullOrWhiteSpace($text)) { throw "시트를 읽지 못했습니다: $sheetPath" }
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

$zip = [System.IO.Compression.ZipFile]::OpenRead($tmp)
try {
    $sharedStrings = @()
    $ssText = Read-ZipText $zip "xl/sharedStrings.xml"
    if ($ssText) {
        [xml]$ssXml = $ssText
        foreach ($si in $ssXml.sst.si) {
            if ($si.t) { $sharedStrings += [string]$si.t }
            elseif ($si.r) { $sharedStrings += (($si.r | ForEach-Object { [string]$_.t }) -join "") }
            else { $sharedStrings += "" }
        }
    }
    $shopRows   = @(Read-SheetRows $zip "xl/worksheets/sheet2.xml" $sharedStrings)
    $courseRows = @(Read-SheetRows $zip "xl/worksheets/sheet3.xml" $sharedStrings)
    $applyRows  = @(Read-SheetRows $zip "xl/worksheets/sheet4.xml" $sharedStrings)
}
finally { $zip.Dispose() }

Write-Host ""
Write-Host "=== 업체기본정보 실제 저장값 ==="
foreach ($r in $shopRows) {
    if (-not [string]::IsNullOrWhiteSpace([string]$r.B)) {
        Write-Host "사용=[$($r.A)] | 업체명=[$($r.B)] | 영업시간=[$($r.C)] | 전화번호=[$($r.D)]"
    }
}

Write-Host ""
Write-Host "=== 코스정보 실제 저장값 ==="
foreach ($r in $courseRows) {
    if (-not [string]::IsNullOrWhiteSpace([string]$r.B)) {
        Write-Host "사용=[$($r.A)] | 업체=[$($r.B)] | 코스=[$($r.C)] | 시간=[$($r.D)] | 가격=[$($r.E)]"
    }
}

Write-Host ""
Write-Host "=== 적용지역 실제 저장값 ==="
foreach ($r in $applyRows) {
    if (-not [string]::IsNullOrWhiteSpace([string]$r.B) -or -not [string]::IsNullOrWhiteSpace([string]$r.C)) {
        Write-Host "사용=[$($r.A)] | 업체=[$($r.B)] | 적용지역=[$($r.C)]"
    }
}

$activeShops = @($shopRows | Where-Object { $_.A -eq "Y" -and -not [string]::IsNullOrWhiteSpace([string]$_.B) }).Count
$activeCourses = @($courseRows | Where-Object { $_.A -eq "Y" -and -not [string]::IsNullOrWhiteSpace([string]$_.B) }).Count
$activeApplies = @($applyRows | Where-Object { $_.A -eq "Y" -and -not [string]::IsNullOrWhiteSpace([string]$_.B) -and -not [string]::IsNullOrWhiteSpace([string]$_.C) }).Count

Write-Host ""
Write-Host "=== 진단 요약 ==="
Write-Host "사용=Y 업체 수:" $activeShops
Write-Host "사용=Y 코스 행:" $activeCourses
Write-Host "사용=Y 적용지역 행:" $activeApplies
Write-Host ""
Write-Host "사이트 HTML 수정: 0"
