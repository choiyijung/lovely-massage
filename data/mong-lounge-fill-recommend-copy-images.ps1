$ErrorActionPreference = "Stop"

$master = "E:\mong-lounge-site\data\몽라운지_업체자동등록_마스터_v2.xlsx"
$stamp = Get-Date -Format "yyyyMMdd-HHmmss"
$backup = "E:\mong-lounge-site\data\몽라운지_업체자동등록_마스터_v2_backup_recommend_$stamp.xlsx"
$imageDir = "E:\mong-lounge-site\assets\images\main-recommend"
$temp = Join-Path $env:TEMP "mong_lounge_recommend_fill_$stamp.xlsx"

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

if (!(Test-Path -LiteralPath $master)) {
    throw "마스터 엑셀을 찾을 수 없습니다: $master"
}

# 엑셀 파일이 열려 있으면 안전하게 중단
try {
    $lock = [System.IO.File]::Open(
        $master,
        [System.IO.FileMode]::Open,
        [System.IO.FileAccess]::ReadWrite,
        [System.IO.FileShare]::None
    )
    $lock.Close()
}
catch {
    Write-Host ""
    Write-Host "=== 작업 중단 ===" -ForegroundColor Yellow
    Write-Host "몽라운지_업체자동등록_마스터_v2.xlsx 파일이 열려 있습니다."
    Write-Host "이 엑셀 파일만 닫고 다시 실행하세요."
    exit 1
}

Copy-Item -LiteralPath $master -Destination $backup -Force
Copy-Item -LiteralPath $master -Destination $temp -Force
New-Item -ItemType Directory -Force -Path $imageDir | Out-Null

function HtmlEncode([string]$text) {
    if ($null -eq $text) { return "" }
    return [System.Net.WebUtility]::HtmlEncode($text)
}

function Get-ZipText {
    param([System.IO.Compression.ZipArchive]$Zip,[string]$Name)
    $e = $Zip.GetEntry($Name)
    if (!$e) { throw "XLSX 내부 파일 누락: $Name" }
    $s = $e.Open()
    try {
        $r = [System.IO.StreamReader]::new($s,[System.Text.Encoding]::UTF8,$true)
        try { return $r.ReadToEnd() } finally { $r.Dispose() }
    } finally { $s.Dispose() }
}

function Set-ZipText {
    param([System.IO.Compression.ZipArchive]$Zip,[string]$Name,[string]$Text)
    $old = $Zip.GetEntry($Name)
    if ($old) { $old.Delete() }
    $e = $Zip.CreateEntry($Name,[System.IO.Compression.CompressionLevel]::Optimal)
    $s = $e.Open()
    try {
        $enc = [System.Text.UTF8Encoding]::new($false)
        $w = [System.IO.StreamWriter]::new($s,$enc)
        try { $w.Write($Text); $w.Flush() } finally { $w.Dispose() }
    } finally { $s.Dispose() }
}

function Get-SharedStrings {
    param([System.IO.Compression.ZipArchive]$Zip)
    $entry = $Zip.GetEntry("xl/sharedStrings.xml")
    $list = New-Object System.Collections.Generic.List[string]
    if (!$entry) { return $list }
    [xml]$xml = Get-ZipText $Zip "xl/sharedStrings.xml"
    $ns = [System.Xml.XmlNamespaceManager]::new($xml.NameTable)
    $ns.AddNamespace("m","http://schemas.openxmlformats.org/spreadsheetml/2006/main")
    foreach ($si in $xml.SelectNodes("//m:si",$ns)) {
        $parts = @($si.SelectNodes(".//m:t",$ns) | ForEach-Object { $_.'#text' })
        $list.Add(($parts -join ""))
    }
    return $list
}

function Get-CellText {
    param(
        [System.Xml.XmlElement]$Cell,
        [System.Collections.Generic.List[string]]$Shared
    )
    if (!$Cell) { return "" }
    $t = $Cell.GetAttribute("t")
    if ($t -eq "inlineStr") {
        $texts = @($Cell.SelectNodes(".//*[local-name()='t']") | ForEach-Object { $_.'#text' })
        return ($texts -join "")
    }
    $v = $Cell.SelectSingleNode("./*[local-name()='v']")
    if (!$v) { return "" }
    $raw = $v.InnerText
    if ($t -eq "s") {
        $idx = 0
        if ([int]::TryParse($raw,[ref]$idx) -and $idx -ge 0 -and $idx -lt $Shared.Count) {
            return $Shared[$idx]
        }
    }
    return $raw
}

function Set-InlineCell {
    param(
        [xml]$SheetXml,
        [int]$RowNum,
        [string]$Col,
        [string]$Text
    )
    $nsUri = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    $sheetData = $SheetXml.SelectSingleNode("/*[local-name()='worksheet']/*[local-name()='sheetData']")
    if (!$sheetData) { throw "sheetData 없음" }

    $row = $SheetXml.SelectSingleNode("//*[local-name()='row' and @r='$RowNum']")
    if (!$row) {
        $row = $SheetXml.CreateElement("row",$nsUri)
        [void]$row.SetAttribute("r",[string]$RowNum)
        [void]$sheetData.AppendChild($row)
    }

    $ref = "$Col$RowNum"
    $cell = $SheetXml.SelectSingleNode("//*[local-name()='c' and @r='$ref']")
    if (!$cell) {
        $cell = $SheetXml.CreateElement("c",$nsUri)
        [void]$cell.SetAttribute("r",$ref)
        [void]$row.AppendChild($cell)
    }

    while ($cell.HasChildNodes) { [void]$cell.RemoveChild($cell.FirstChild) }
    [void]$cell.SetAttribute("t","inlineStr")

    $is = $SheetXml.CreateElement("is",$nsUri)
    $tx = $SheetXml.CreateElement("t",$nsUri)
    $tx.InnerText = $Text
    [void]$is.AppendChild($tx)
    [void]$cell.AppendChild($is)
}

function Find-SheetPath {
    param(
        [xml]$WorkbookXml,
        [xml]$RelsXml,
        [string]$SheetName
    )
    $wbNs = [System.Xml.XmlNamespaceManager]::new($WorkbookXml.NameTable)
    $wbNs.AddNamespace("m","http://schemas.openxmlformats.org/spreadsheetml/2006/main")
    $wbNs.AddNamespace("r","http://schemas.openxmlformats.org/officeDocument/2006/relationships")

    $sheet = $WorkbookXml.SelectSingleNode("//m:sheet[@name='$SheetName']",$wbNs)
    if (!$sheet) { throw "시트를 찾을 수 없습니다: $SheetName" }

    $rid = $sheet.GetAttribute("id","http://schemas.openxmlformats.org/officeDocument/2006/relationships")

    $relNs = [System.Xml.XmlNamespaceManager]::new($RelsXml.NameTable)
    $relNs.AddNamespace("r","http://schemas.openxmlformats.org/package/2006/relationships")
    $rel = $RelsXml.SelectSingleNode("//r:Relationship[@Id='$rid']",$relNs)
    if (!$rel) { throw "시트 관계를 찾을 수 없습니다: $SheetName" }

    $target = $rel.GetAttribute("Target").Replace("\","/")
    if ($target.StartsWith("/")) { return $target.TrimStart("/") }
    return "xl/" + $target.TrimStart("/")
}

function New-RecommendSvg {
    param(
        [string]$Path,
        [string]$Label,
        [string]$Shop,
        [string]$Region,
        [string]$C1,
        [string]$C2
    )

    $safeLabel = HtmlEncode $Label
    $safeShop = HtmlEncode ($(if([string]::IsNullOrWhiteSpace($Shop)){$Label}else{$Shop}))
    $safeRegion = HtmlEncode ($(if([string]::IsNullOrWhiteSpace($Region)){"추천 업체"}else{$Region}))

    $svg = @"
<svg xmlns="http://www.w3.org/2000/svg" width="900" height="520" viewBox="0 0 900 520">
  <defs>
    <linearGradient id="g" x1="0" y1="0" x2="1" y2="1">
      <stop offset="0%" stop-color="$C1"/>
      <stop offset="100%" stop-color="$C2"/>
    </linearGradient>
    <filter id="shadow" x="-20%" y="-20%" width="140%" height="140%">
      <feDropShadow dx="0" dy="14" stdDeviation="20" flood-color="#000000" flood-opacity=".28"/>
    </filter>
  </defs>
  <rect width="900" height="520" rx="36" fill="url(#g)"/>
  <circle cx="760" cy="110" r="150" fill="#ffffff" opacity=".07"/>
  <circle cx="110" cy="430" r="210" fill="#ffffff" opacity=".05"/>
  <g filter="url(#shadow)">
    <rect x="66" y="64" width="768" height="392" rx="30" fill="#081f1b" opacity=".68"/>
  </g>
  <rect x="94" y="94" width="145" height="42" rx="21" fill="#ffc928"/>
  <text x="166" y="122" text-anchor="middle" font-family="Arial, sans-serif" font-size="20" font-weight="800" fill="#083f35">$safeLabel</text>
  <text x="94" y="240" font-family="Arial, sans-serif" font-size="50" font-weight="800" fill="#ffffff">$safeShop</text>
  <text x="96" y="300" font-family="Arial, sans-serif" font-size="25" font-weight="500" fill="#d6f5ec">$safeRegion</text>
  <line x1="96" y1="344" x2="800" y2="344" stroke="#ffc928" stroke-width="2" opacity=".88"/>
  <text x="96" y="395" font-family="Arial, sans-serif" font-size="20" font-weight="600" fill="#ffffff" opacity=".86">MONG LOUNGE RECOMMEND</text>
</svg>
"@
    [System.IO.File]::WriteAllText($Path,$svg,[System.Text.UTF8Encoding]::new($false))
}

$zip = $null
$vipChanged = 0
$premiumChanged = 0
$imageCount = 0

$vipTexts = @(
    "지역과 이용 조건을 한눈에 확인하기 좋은 추천 업체입니다. 원하는 일정과 코스를 비교한 뒤 편하게 문의해보세요.",
    "깔끔한 이용 안내와 편리한 예약 확인을 중심으로 살펴볼 수 있는 추천 업체입니다. 위치와 시간을 확인해 선택해보세요.",
    "지역별 이용 범위와 코스 정보를 간편하게 비교할 수 있는 추천 업체입니다. 방문 전 필요한 내용을 확인해보세요.",
    "일정에 맞춰 업체 정보와 이용 조건을 빠르게 확인할 수 있는 추천 업체입니다. 상세페이지에서 코스와 안내를 확인해보세요."
)
$premiumTexts = @(
    "차분한 분위기와 편리한 이용 정보 확인을 중심으로 구성한 프리미엄 추천 업체입니다. 지역과 코스를 비교해보세요.",
    "업체 정보와 이용 조건을 보기 쉽게 확인할 수 있는 프리미엄 추천 업체입니다. 원하는 일정에 맞춰 살펴보세요.",
    "지역, 가격, 코스 정보를 한 번에 비교하기 좋은 프리미엄 추천 업체입니다. 상세 안내를 확인한 뒤 선택해보세요.",
    "편리한 예약 확인과 명확한 이용 안내를 기준으로 살펴볼 수 있는 프리미엄 추천 업체입니다. 필요한 정보를 확인해보세요."
)

$palettes = @(
    @("#0a584a","#062f28"),
    @("#73501f","#302114"),
    @("#4d315e","#201626"),
    @("#244f65","#102934"),
    @("#0d4d43","#082722"),
    @("#5c3b21","#281b12"),
    @("#43305e","#21172c"),
    @("#3f5560","#1e2a30")
)

try {
    $zip = [System.IO.Compression.ZipFile]::Open(
        $temp,
        [System.IO.Compression.ZipArchiveMode]::Update
    )

    [xml]$wbXml = Get-ZipText $zip "xl/workbook.xml"
    [xml]$relXml = Get-ZipText $zip "xl/_rels/workbook.xml.rels"
    $shared = Get-SharedStrings $zip

    $defs = @(
        @{ Sheet="메인VIP추천"; Prefix="vip"; Label="VIP 추천"; Texts=$vipTexts; StartPalette=0 },
        @{ Sheet="메인프리미엄추천"; Prefix="premium"; Label="PREMIUM"; Texts=$premiumTexts; StartPalette=4 }
    )

    foreach ($def in $defs) {
        $sheetPath = Find-SheetPath $wbXml $relXml $def.Sheet
        [xml]$sx = Get-ZipText $zip $sheetPath

        for ($row=2; $row -le 5; $row++) {
            $shopCell = $sx.SelectSingleNode("//*[local-name()='c' and @r='B$row']")
            $regionCell = $sx.SelectSingleNode("//*[local-name()='c' and @r='C$row']")
            $imageCell = $sx.SelectSingleNode("//*[local-name()='c' and @r='E$row']")
            $descCell = $sx.SelectSingleNode("//*[local-name()='c' and @r='G$row']")

            $shop = Get-CellText $shopCell $shared
            $region = Get-CellText $regionCell $shared
            $currentImage = Get-CellText $imageCell $shared
            $currentDesc = Get-CellText $descCell $shared

            $idx = $row - 2
            $fileName = "{0}-{1:00}.svg" -f $def.Prefix,($idx+1)
            $webPath = "/assets/images/main-recommend/$fileName"
            $diskPath = Join-Path $imageDir $fileName
            $palette = $palettes[$def.StartPalette + $idx]

            New-RecommendSvg `
                -Path $diskPath `
                -Label ("{0} {1}" -f $def.Label,($idx+1)) `
                -Shop $shop `
                -Region $region `
                -C1 $palette[0] `
                -C2 $palette[1]
            $imageCount++

            if ([string]::IsNullOrWhiteSpace($currentImage)) {
                Set-InlineCell $sx $row "E" $webPath
            }

            if ([string]::IsNullOrWhiteSpace($currentDesc)) {
                Set-InlineCell $sx $row "G" $def.Texts[$idx]
            }

            if ($def.Prefix -eq "vip") { $vipChanged++ } else { $premiumChanged++ }
        }

        Set-ZipText $zip $sheetPath $sx.OuterXml
    }
}
finally {
    if ($zip) { $zip.Dispose() }
}

Copy-Item -LiteralPath $temp -Destination $master -Force
Remove-Item -LiteralPath $temp -Force -ErrorAction SilentlyContinue

Write-Host ""
Write-Host "=== 메인 추천 소개문구 + 이미지 자동 입력 완료 ==="
Write-Host "VIP 추천 행 처리:" $vipChanged
Write-Host "프리미엄 추천 행 처리:" $premiumChanged
Write-Host "추천 이미지 생성:" $imageCount
Write-Host "이미지 폴더:" $imageDir
Write-Host ""
Write-Host "소개문구: 비어 있는 칸만 자동 입력"
Write-Host "이미지: 비어 있는 칸만 자동 경로 입력"
Write-Host "업체명/지역이 입력돼 있으면 이미지에도 자동 반영"
Write-Host "업체명/지역이 아직 비어 있어도 나중에 스크립트를 다시 실행하면 이미지가 갱신됩니다."
Write-Host ""
Write-Host "백업 파일:" $backup
Write-Host "마스터 파일:" $master
