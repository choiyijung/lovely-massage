$ErrorActionPreference = "Stop"

$master = "E:\mong-lounge-site\data\몽라운지_업체자동등록_마스터_v2.xlsx"
$stamp = Get-Date -Format "yyyyMMdd-HHmmss"
$backup = "E:\mong-lounge-site\data\몽라운지_업체자동등록_마스터_v2_backup_$stamp.xlsx"
$temp = Join-Path $env:TEMP "mong_lounge_master_upgrade_$stamp.xlsx"

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

if (!(Test-Path -LiteralPath $master)) {
    throw "마스터 엑셀을 찾을 수 없습니다: $master"
}

# 엑셀에서 현재 마스터 파일이 열려 있으면 안전하게 중단
try {
    $lockTest = [System.IO.File]::Open(
        $master,
        [System.IO.FileMode]::Open,
        [System.IO.FileAccess]::ReadWrite,
        [System.IO.FileShare]::None
    )
    $lockTest.Close()
}
catch {
    Write-Host ""
    Write-Host "=== 작업 중단 ===" -ForegroundColor Yellow
    Write-Host "현재 마스터 엑셀이 열려 있습니다."
    Write-Host "엑셀에서 '몽라운지_업체자동등록_마스터_v2.xlsx' 파일만 닫은 뒤 다시 실행하세요."
    exit 1
}

Copy-Item -LiteralPath $master -Destination $backup -Force
Copy-Item -LiteralPath $master -Destination $temp -Force

function Get-ZipText {
    param(
        [System.IO.Compression.ZipArchive]$Zip,
        [string]$Name
    )
    $entry = $Zip.GetEntry($Name)
    if (!$entry) { throw "XLSX 내부 파일 누락: $Name" }

    $stream = $entry.Open()
    try {
        $reader = [System.IO.StreamReader]::new(
            $stream,
            [System.Text.Encoding]::UTF8,
            $true
        )
        try { return $reader.ReadToEnd() }
        finally { $reader.Dispose() }
    }
    finally { $stream.Dispose() }
}

function Set-ZipText {
    param(
        [System.IO.Compression.ZipArchive]$Zip,
        [string]$Name,
        [string]$Text
    )

    $old = $Zip.GetEntry($Name)
    if ($old) { $old.Delete() }

    $entry = $Zip.CreateEntry(
        $Name,
        [System.IO.Compression.CompressionLevel]::Optimal
    )
    $stream = $entry.Open()
    try {
        $utf8 = [System.Text.UTF8Encoding]::new($false)
        $writer = [System.IO.StreamWriter]::new($stream, $utf8)
        try {
            $writer.Write($Text)
            $writer.Flush()
        }
        finally { $writer.Dispose() }
    }
    finally { $stream.Dispose() }
}

function New-RecommendSheetXml {
    param([string]$TypeName)

    return @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:H5"/>
  <sheetViews>
    <sheetView workbookViewId="0">
      <pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>
    </sheetView>
  </sheetViews>
  <sheetFormatPr defaultRowHeight="18"/>
  <cols>
    <col min="1" max="1" width="9" customWidth="1"/>
    <col min="2" max="2" width="24" customWidth="1"/>
    <col min="3" max="3" width="18" customWidth="1"/>
    <col min="4" max="4" width="14" customWidth="1"/>
    <col min="5" max="5" width="32" customWidth="1"/>
    <col min="6" max="6" width="42" customWidth="1"/>
    <col min="7" max="7" width="40" customWidth="1"/>
    <col min="8" max="8" width="10" customWidth="1"/>
  </cols>
  <sheetData>
    <row r="1" ht="24" customHeight="1">
      <c r="A1" t="inlineStr"><is><t>사용</t></is></c>
      <c r="B1" t="inlineStr"><is><t>업체명</t></is></c>
      <c r="C1" t="inlineStr"><is><t>지역</t></is></c>
      <c r="D1" t="inlineStr"><is><t>가격</t></is></c>
      <c r="E1" t="inlineStr"><is><t>이미지</t></is></c>
      <c r="F1" t="inlineStr"><is><t>상세페이지링크</t></is></c>
      <c r="G1" t="inlineStr"><is><t>소개문구</t></is></c>
      <c r="H1" t="inlineStr"><is><t>정렬순서</t></is></c>
    </row>
    <row r="2"><c r="A2" t="inlineStr"><is><t>Y</t></is></c><c r="H2"><v>1</v></c></row>
    <row r="3"><c r="A3" t="inlineStr"><is><t>Y</t></is></c><c r="H3"><v>2</v></c></row>
    <row r="4"><c r="A4" t="inlineStr"><is><t>Y</t></is></c><c r="H4"><v>3</v></c></row>
    <row r="5"><c r="A5" t="inlineStr"><is><t>Y</t></is></c><c r="H5"><v>4</v></c></row>
  </sheetData>
  <dataValidations count="1">
    <dataValidation type="list" allowBlank="1" showErrorMessage="1" sqref="A2:A100">
      <formula1>&quot;Y,N&quot;</formula1>
    </dataValidation>
  </dataValidations>
</worksheet>
"@
}

$zip = $null
$added = New-Object System.Collections.Generic.List[string]

try {
    $zip = [System.IO.Compression.ZipFile]::Open(
        $temp,
        [System.IO.Compression.ZipArchiveMode]::Update
    )

    [xml]$workbookXml = Get-ZipText $zip "xl/workbook.xml"
    [xml]$relsXml = Get-ZipText $zip "xl/_rels/workbook.xml.rels"
    [xml]$contentXml = Get-ZipText $zip "[Content_Types].xml"

    $mainNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    $officeRelNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    $packageRelNs = "http://schemas.openxmlformats.org/package/2006/relationships"
    $contentNs = "http://schemas.openxmlformats.org/package/2006/content-types"

    $wbMgr = [System.Xml.XmlNamespaceManager]::new($workbookXml.NameTable)
    $wbMgr.AddNamespace("m",$mainNs)

    $relMgr = [System.Xml.XmlNamespaceManager]::new($relsXml.NameTable)
    $relMgr.AddNamespace("r",$packageRelNs)

    $ctMgr = [System.Xml.XmlNamespaceManager]::new($contentXml.NameTable)
    $ctMgr.AddNamespace("c",$contentNs)

    $existingSheets = @(
        $workbookXml.SelectNodes("//m:sheets/m:sheet",$wbMgr) |
        ForEach-Object { $_.GetAttribute("name") }
    )

    $sheetIds = @(
        $workbookXml.SelectNodes("//m:sheets/m:sheet",$wbMgr) |
        ForEach-Object { [int]$_.GetAttribute("sheetId") }
    )
    $nextSheetId = if ($sheetIds.Count) {
        ($sheetIds | Measure-Object -Maximum).Maximum + 1
    } else { 1 }

    $ridNums = @(
        $relsXml.SelectNodes("//r:Relationship",$relMgr) |
        ForEach-Object {
            if ($_.GetAttribute("Id") -match '^rId(\d+)$') { [int]$Matches[1] }
        }
    )
    $nextRid = if ($ridNums.Count) {
        ($ridNums | Measure-Object -Maximum).Maximum + 1
    } else { 1 }

    $sheetFileNums = @(
        $zip.Entries |
        Where-Object { $_.FullName -match '^xl/worksheets/sheet(\d+)\.xml$' } |
        ForEach-Object {
            [void]($_.FullName -match '^xl/worksheets/sheet(\d+)\.xml$')
            [int]$Matches[1]
        }
    )
    $nextSheetFile = if ($sheetFileNums.Count) {
        ($sheetFileNums | Measure-Object -Maximum).Maximum + 1
    } else { 1 }

    $sheetDefs = @(
        @{ Name = "메인VIP추천"; Type = "VIP" },
        @{ Name = "메인프리미엄추천"; Type = "PREMIUM" }
    )

    foreach ($def in $sheetDefs) {
        $sheetName = $def.Name

        if ($existingSheets -contains $sheetName) {
            Write-Host "이미 존재 - 유지:" $sheetName
            continue
        }

        $rid = "rId$nextRid"
        $sheetPath = "xl/worksheets/sheet$nextSheetFile.xml"
        $target = "worksheets/sheet$nextSheetFile.xml"

        # worksheet xml 추가
        Set-ZipText $zip $sheetPath (New-RecommendSheetXml $def.Type)

        # workbook.xml 에 sheet 등록
        $sheetNode = $workbookXml.CreateElement("sheet",$mainNs)
        [void]$sheetNode.SetAttribute("name",$sheetName)
        [void]$sheetNode.SetAttribute("sheetId",[string]$nextSheetId)
        [void]$sheetNode.SetAttribute("id",$officeRelNs,$rid)
        [void]$workbookXml.SelectSingleNode("//m:sheets",$wbMgr).AppendChild($sheetNode)

        # workbook relationship 등록
        $relNode = $relsXml.CreateElement("Relationship",$packageRelNs)
        [void]$relNode.SetAttribute("Id",$rid)
        [void]$relNode.SetAttribute(
            "Type",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
        )
        [void]$relNode.SetAttribute("Target",$target)
        [void]$relsXml.DocumentElement.AppendChild($relNode)

        # content types 등록
        $override = $contentXml.CreateElement("Override",$contentNs)
        [void]$override.SetAttribute("PartName","/$sheetPath")
        [void]$override.SetAttribute(
            "ContentType",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
        )
        [void]$contentXml.DocumentElement.AppendChild($override)

        $added.Add($sheetName)

        $nextSheetId++
        $nextRid++
        $nextSheetFile++
    }

    if ($added.Count -gt 0) {
        Set-ZipText $zip "xl/workbook.xml" $workbookXml.OuterXml
        Set-ZipText $zip "xl/_rels/workbook.xml.rels" $relsXml.OuterXml
        Set-ZipText $zip "[Content_Types].xml" $contentXml.OuterXml
    }
}
finally {
    if ($zip) { $zip.Dispose() }
}

if ($added.Count -gt 0) {
    Copy-Item -LiteralPath $temp -Destination $master -Force
}
Remove-Item -LiteralPath $temp -Force -ErrorAction SilentlyContinue

# 최종 검증
$verifyZip = [System.IO.Compression.ZipFile]::OpenRead($master)
try {
    [xml]$verifyWb = Get-ZipText $verifyZip "xl/workbook.xml"
    $vm = [System.Xml.XmlNamespaceManager]::new($verifyWb.NameTable)
    $vm.AddNamespace("m","http://schemas.openxmlformats.org/spreadsheetml/2006/main")
    $names = @(
        $verifyWb.SelectNodes("//m:sheets/m:sheet",$vm) |
        ForEach-Object { $_.GetAttribute("name") }
    )
}
finally {
    $verifyZip.Dispose()
}

Write-Host ""
Write-Host "=== 기존 마스터 엑셀 업그레이드 완료 ==="
Write-Host "원본 유지 경로:" $master
Write-Host "백업 파일:" $backup
Write-Host "추가된 시트:" ($(if($added.Count){$added -join ", "}else{"추가 없음 - 이미 존재"}))
Write-Host "메인VIP추천 확인:" ($names -contains "메인VIP추천")
Write-Host "메인프리미엄추천 확인:" ($names -contains "메인프리미엄추천")
Write-Host ""
Write-Host "기존 업체기본정보/코스정보/적용지역/자동생성규칙 데이터는 그대로 유지됩니다."
Write-Host "추천 시트는 각각 4줄 입력용으로 추가했습니다."
