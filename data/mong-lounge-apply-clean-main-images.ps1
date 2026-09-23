$ErrorActionPreference = "Stop"

$root      = "E:\mong-lounge-site"
$dataDir   = Join-Path $root "data"
$master    = Join-Path $dataDir "몽라운지_업체자동등록_마스터_v2.xlsx"
$appJs     = Join-Path $root "assets\js\app.js"
$index     = Join-Path $root "index.html"
$imageDir  = Join-Path $root "assets\images\main-recommend"
$stamp     = Get-Date -Format "yyyyMMdd-HHmmss"
$backup    = "E:\mong-lounge-clean-main-images-backup-$stamp"
$tempDir   = Join-Path $env:TEMP "mong_lounge_clean_images_$stamp"
$tempXlsx  = Join-Path $env:TEMP "mong_lounge_master_images_$stamp.xlsx"

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$zipFile = Get-ChildItem "$env:USERPROFILE\Downloads" -File |
    Where-Object { $_.Name -like "mong-lounge-clean-main-images*.zip" } |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1

if (!$zipFile) {
    throw "다운로드 폴더에서 mong-lounge-clean-main-images.zip 파일을 찾지 못했습니다."
}

if (!(Test-Path -LiteralPath $master)) { throw "마스터 엑셀 없음: $master" }
if (!(Test-Path -LiteralPath $appJs))  { throw "app.js 없음: $appJs" }

# 마스터 엑셀은 닫은 상태에서 진행
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
    Write-Host "몽라운지_업체자동등록_마스터_v2.xlsx 파일을 닫고 다시 실행하세요."
    exit 1
}

New-Item -ItemType Directory -Force -Path $backup,$tempDir,$imageDir | Out-Null

Copy-Item -LiteralPath $master -Destination (Join-Path $backup "몽라운지_업체자동등록_마스터_v2.xlsx") -Force
Copy-Item -LiteralPath $appJs  -Destination (Join-Path $backup "app.js") -Force
if (Test-Path -LiteralPath $index) {
    Copy-Item -LiteralPath $index -Destination (Join-Path $backup "index.html") -Force
}

# 기존 추천 이미지 백업
foreach($name in @(
    "vip-01.svg","vip-02.svg","vip-03.svg","vip-04.svg",
    "premium-01.svg","premium-02.svg","premium-03.svg","premium-04.svg",
    "vip-01.png","vip-02.png","vip-03.png","vip-04.png",
    "premium-01.png","premium-02.png","premium-03.png","premium-04.png"
)){
    $p = Join-Path $imageDir $name
    if(Test-Path -LiteralPath $p){
        Copy-Item -LiteralPath $p -Destination (Join-Path $backup $name) -Force
    }
}

Expand-Archive -LiteralPath $zipFile.FullName -DestinationPath $tempDir -Force

$names = @(
    "vip-01.png","vip-02.png","vip-03.png","vip-04.png",
    "premium-01.png","premium-02.png","premium-03.png","premium-04.png"
)

foreach($name in $names){
    $src = Join-Path $tempDir $name
    if(!(Test-Path -LiteralPath $src)){ throw "압축파일 내부 이미지 누락: $name" }
    Copy-Item -LiteralPath $src -Destination (Join-Path $imageDir $name) -Force
}

# app.js와 관련 HTML의 정확한 추천 이미지 경로만 png로 변경
$filesToPatch = @(
    $appJs,
    (Join-Path $root "index.html"),
    (Join-Path $root "vip.html"),
    (Join-Path $root "premium.html")
) | Where-Object { Test-Path -LiteralPath $_ }

foreach($f in $filesToPatch){
    $text = Get-Content -LiteralPath $f -Raw -Encoding UTF8
    $before = $text

    foreach($base in @(
        "vip-01","vip-02","vip-03","vip-04",
        "premium-01","premium-02","premium-03","premium-04"
    )){
        $text = $text.Replace("assets/images/main-recommend/$base.svg","assets/images/main-recommend/$base.png")
        $text = $text.Replace("/assets/images/main-recommend/$base.svg","/assets/images/main-recommend/$base.png")
    }

    if($text -ne $before){
        Set-Content -LiteralPath $f -Value $text -Encoding UTF8
    }
}

# index의 app.js 캐시 버전 갱신
if(Test-Path -LiteralPath $index){
    $html = Get-Content -LiteralPath $index -Raw -Encoding UTF8
    $html = [regex]::Replace(
        $html,
        'assets/js/app\.js(?:\?v=[^"'']*)?',
        "assets/js/app.js?v=$stamp"
    )
    Set-Content -LiteralPath $index -Value $html -Encoding UTF8
}

# ===== 엑셀 이미지 경로 E2:E5 업데이트 =====
Copy-Item -LiteralPath $master -Destination $tempXlsx -Force

function Get-ZipText {
    param([System.IO.Compression.ZipArchive]$Zip,[string]$Name)
    $e = $Zip.GetEntry($Name)
    if(!$e){ throw "XLSX 내부 파일 누락: $Name" }
    $s = $e.Open()
    try{
        $r = [System.IO.StreamReader]::new($s,[System.Text.Encoding]::UTF8,$true)
        try { return $r.ReadToEnd() } finally { $r.Dispose() }
    } finally { $s.Dispose() }
}

function Set-ZipText {
    param([System.IO.Compression.ZipArchive]$Zip,[string]$Name,[string]$Text)
    $old=$Zip.GetEntry($Name)
    if($old){ $old.Delete() }
    $e=$Zip.CreateEntry($Name,[System.IO.Compression.CompressionLevel]::Optimal)
    $s=$e.Open()
    try{
        $w=[System.IO.StreamWriter]::new($s,[System.Text.UTF8Encoding]::new($false))
        try { $w.Write($Text); $w.Flush() } finally { $w.Dispose() }
    } finally { $s.Dispose() }
}

function Get-SheetPath {
    param([xml]$Wb,[xml]$Rels,[string]$Name)
    $sheet=$Wb.SelectSingleNode("//*[local-name()='sheet' and @name='$Name']")
    if(!$sheet){ throw "시트 없음: $Name" }
    $rid=$sheet.GetAttribute("id","http://schemas.openxmlformats.org/officeDocument/2006/relationships")
    $rel=$Rels.SelectSingleNode("//*[local-name()='Relationship' and @Id='$rid']")
    if(!$rel){ throw "시트 관계 없음: $Name" }
    $target=$rel.GetAttribute("Target").Replace("\","/")
    if($target.StartsWith("/")){ return $target.TrimStart("/") }
    return "xl/"+$target.TrimStart("/")
}

function Set-InlineCell {
    param([xml]$Sheet,[int]$Row,[string]$Col,[string]$Text)

    $ns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    $sheetData=$Sheet.SelectSingleNode("/*[local-name()='worksheet']/*[local-name()='sheetData']")
    $rowNode=$Sheet.SelectSingleNode("//*[local-name()='row' and @r='$Row']")

    if(!$rowNode){
        $rowNode=$Sheet.CreateElement("row",$ns)
        [void]$rowNode.SetAttribute("r",[string]$Row)
        [void]$sheetData.AppendChild($rowNode)
    }

    $ref="$Col$Row"
    $cell=$Sheet.SelectSingleNode("//*[local-name()='c' and @r='$ref']")

    if(!$cell){
        $cell=$Sheet.CreateElement("c",$ns)
        [void]$cell.SetAttribute("r",$ref)
        [void]$rowNode.AppendChild($cell)
    }

    while($cell.HasChildNodes){ [void]$cell.RemoveChild($cell.FirstChild) }
    [void]$cell.SetAttribute("t","inlineStr")

    $is=$Sheet.CreateElement("is",$ns)
    $tx=$Sheet.CreateElement("t",$ns)
    $tx.InnerText=$Text
    [void]$is.AppendChild($tx)
    [void]$cell.AppendChild($is)
}

$xlsxZip = [System.IO.Compression.ZipFile]::Open(
    $tempXlsx,
    [System.IO.Compression.ZipArchiveMode]::Update
)

try{
    [xml]$wb   = Get-ZipText $xlsxZip "xl/workbook.xml"
    [xml]$rels = Get-ZipText $xlsxZip "xl/_rels/workbook.xml.rels"

    $vipPath = Get-SheetPath $wb $rels "메인VIP추천"
    $prePath = Get-SheetPath $wb $rels "메인프리미엄추천"

    [xml]$vipXml = Get-ZipText $xlsxZip $vipPath
    [xml]$preXml = Get-ZipText $xlsxZip $prePath

    for($i=1;$i -le 4;$i++){
        $row = $i + 1
        Set-InlineCell $vipXml $row "E" ("/assets/images/main-recommend/vip-{0:00}.png" -f $i)
        Set-InlineCell $preXml $row "E" ("/assets/images/main-recommend/premium-{0:00}.png" -f $i)
    }

    Set-ZipText $xlsxZip $vipPath $vipXml.OuterXml
    Set-ZipText $xlsxZip $prePath $preXml.OuterXml
}
finally{
    $xlsxZip.Dispose()
}

Copy-Item -LiteralPath $tempXlsx -Destination $master -Force
Remove-Item -LiteralPath $tempXlsx -Force -ErrorAction SilentlyContinue
Remove-Item -LiteralPath $tempDir -Recurse -Force -ErrorAction SilentlyContinue

# 검증
$pngCount = @(
    Get-ChildItem -LiteralPath $imageDir -File -Filter "*.png" |
    Where-Object { $_.Name -match '^(vip|premium)-0[1-4]\.png$' }
).Count

$appCheck = Get-Content -LiteralPath $appJs -Raw -Encoding UTF8
$pngRefs = ([regex]::Matches(
    $appCheck,
    '(?:vip|premium)-0[1-4]\.png'
)).Count

$svgRefs = ([regex]::Matches(
    $appCheck,
    '(?:vip|premium)-0[1-4]\.svg'
)).Count

Write-Host ""
Write-Host "=== 메인 추천 이미지 교체 완료 ==="
Write-Host "PNG 이미지 설치:" "$pngCount / 8"
Write-Host "app.js PNG 경로 확인:" $pngRefs
Write-Host "app.js 기존 SVG 경로 남음:" $svgRefs
Write-Host "엑셀 추천 이미지 경로: PNG로 갱신"
Write-Host "백업 폴더:" $backup
Write-Host ""
Write-Host "이제 5512 메인에서 Ctrl+F5 하세요."
