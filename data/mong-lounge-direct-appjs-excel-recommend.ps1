$ErrorActionPreference = "Stop"

$root    = "E:\mong-lounge-site"
$dataDir = Join-Path $root "data"
$master  = Join-Path $dataDir "몽라운지_업체자동등록_마스터_v2.xlsx"
$appJs   = Join-Path $root "assets\js\app.js"
$index   = Join-Path $root "index.html"

$stamp   = Get-Date -Format "yyyyMMdd-HHmmss"
$backup  = "E:\mong-lounge-appjs-main-recommend-backup-$stamp"
$temp    = Join-Path $env:TEMP "mong_lounge_master_read_$stamp.xlsx"

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

if (!(Test-Path -LiteralPath $master)) { throw "마스터 엑셀 없음: $master" }
if (!(Test-Path -LiteralPath $appJs))  { throw "app.js 없음: $appJs" }
if (!(Test-Path -LiteralPath $index))  { throw "index.html 없음: $index" }

New-Item -ItemType Directory -Force -Path $backup | Out-Null
Copy-Item -LiteralPath $appJs -Destination (Join-Path $backup "app.js") -Force
Copy-Item -LiteralPath $index -Destination (Join-Path $backup "index.html") -Force

# 엑셀이 열려 있어도 읽을 수 있게 공유 읽기로 임시 복사
$inStream = [System.IO.File]::Open(
    $master,
    [System.IO.FileMode]::Open,
    [System.IO.FileAccess]::Read,
    [System.IO.FileShare]::ReadWrite
)
try {
    $outStream = [System.IO.File]::Create($temp)
    try { $inStream.CopyTo($outStream) }
    finally { $outStream.Dispose() }
}
finally { $inStream.Dispose() }

function Get-ZipText {
    param([System.IO.Compression.ZipArchive]$Zip,[string]$Name)

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

function Get-SharedStrings {
    param([System.IO.Compression.ZipArchive]$Zip)

    $list = New-Object System.Collections.Generic.List[string]
    if (!$Zip.GetEntry("xl/sharedStrings.xml")) { return $list }

    [xml]$xml = Get-ZipText $Zip "xl/sharedStrings.xml"
    foreach ($si in $xml.SelectNodes("//*[local-name()='si']")) {
        $parts = @(
            $si.SelectNodes(".//*[local-name()='t']") |
            ForEach-Object { $_.InnerText }
        )
        $list.Add(($parts -join ""))
    }
    return $list
}

function Get-CellText {
    param([System.Xml.XmlElement]$Cell,$Shared)

    if (!$Cell) { return "" }

    $type = $Cell.GetAttribute("t")

    if ($type -eq "inlineStr") {
        return (
            @(
                $Cell.SelectNodes(".//*[local-name()='t']") |
                ForEach-Object { $_.InnerText }
            ) -join ""
        )
    }

    $v = $Cell.SelectSingleNode("./*[local-name()='v']")
    if (!$v) { return "" }

    if ($type -eq "s") {
        $i = 0
        if ([int]::TryParse($v.InnerText,[ref]$i) -and
            $i -ge 0 -and $i -lt $Shared.Count) {
            return [string]$Shared[$i]
        }
    }

    return [string]$v.InnerText
}

function Get-SheetPath {
    param([xml]$Workbook,[xml]$Rels,[string]$SheetName)

    $sheet = $Workbook.SelectSingleNode(
        "//*[local-name()='sheet' and @name='$SheetName']"
    )
    if (!$sheet) { throw "시트 없음: $SheetName" }

    $rid = $sheet.GetAttribute(
        "id",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    )

    $rel = $Rels.SelectSingleNode(
        "//*[local-name()='Relationship' and @Id='$rid']"
    )
    if (!$rel) { throw "시트 관계 없음: $SheetName" }

    $target = $rel.GetAttribute("Target").Replace("\","/")
    if ($target.StartsWith("/")) { return $target.TrimStart("/") }

    return "xl/" + $target.TrimStart("/")
}

function Read-RecommendSheet {
    param(
        [System.IO.Compression.ZipArchive]$Zip,
        [xml]$Workbook,
        [xml]$Rels,
        $Shared,
        [string]$SheetName,
        [string]$TypeName
    )

    $path = Get-SheetPath $Workbook $Rels $SheetName
    [xml]$sheet = Get-ZipText $Zip $path

    $rows = @()

    foreach ($row in $sheet.SelectNodes(
        "//*[local-name()='row' and number(@r)>=2]"
    )) {
        $r = [int]$row.GetAttribute("r")
        $v = @{}

        foreach ($col in @("A","B","C","D","E","F","G","H")) {
            $cell = $sheet.SelectSingleNode(
                "//*[local-name()='c' and @r='$col$r']"
            )
            $v[$col] = (Get-CellText $cell $Shared).Trim()
        }

        # blank / Y = 사용, N만 제외
        if ($v["A"].ToUpperInvariant() -eq "N") { continue }
        if ([string]::IsNullOrWhiteSpace($v["B"])) { continue }

        $sort = 999
        if ($v["H"]) {
            [void][int]::TryParse($v["H"],[ref]$sort)
        }

        $image = $v["E"]
        if ($image -and $image -notmatch '^(?:https?://|/)') {
            $image = "/" + $image.TrimStart("/")
        }

        $link = $v["F"]
        if ($link -and $link -notmatch '^(?:https?://|/)') {
            $link = "/" + $link.TrimStart("/")
        }

        $desc = $v["G"]
        if ([string]::IsNullOrWhiteSpace($desc)) {
            $desc = "$($v["C"]) · 상세 안내 확인"
        }

        $rows += [pscustomobject]@{
            name    = $v["B"]
            area    = $(if($v["C"]){$v["C"]}else{"추천 지역"})
            address = $desc
            price   = $(if($v["D"]){$v["D"]}else{"가격 문의"})
            image   = $image
            link    = $(if($link){$link}else{"#"})
            type    = $TypeName
            sort    = $sort
        }
    }

    return @(
        $rows |
        Sort-Object sort,name |
        Select-Object -First 4
    )
}

$zip = [System.IO.Compression.ZipFile]::OpenRead($temp)

try {
    [xml]$workbook = Get-ZipText $zip "xl/workbook.xml"
    [xml]$rels     = Get-ZipText $zip "xl/_rels/workbook.xml.rels"
    $shared        = Get-SharedStrings $zip

    $vip = @(
        Read-RecommendSheet `
            $zip $workbook $rels $shared `
            "메인VIP추천" "VIP"
    )

    $premium = @(
        Read-RecommendSheet `
            $zip $workbook $rels $shared `
            "메인프리미엄추천" "PICK"
    )
}
finally {
    $zip.Dispose()
    Remove-Item -LiteralPath $temp -Force -ErrorAction SilentlyContinue
}

if ($vip.Count -ne 4) {
    throw "메인VIP추천에 사용 가능한 업체가 4개가 아닙니다. 현재: $($vip.Count)"
}
if ($premium.Count -ne 4) {
    throw "메인프리미엄추천에 사용 가능한 업체가 4개가 아닙니다. 현재: $($premium.Count)"
}

$vipJson = $vip | ConvertTo-Json -Depth 5 -Compress
$preJson = $premium | ConvertTo-Json -Depth 5 -Compress

$js = Get-Content -LiteralPath $appJs -Raw -Encoding UTF8
$before = $js

# 1) 기존 임시 shops 배열을 VIP / PREMIUM 실제 데이터로 직접 교체
$arrayBlock = @"
      const vipShops = $vipJson;

      const premiumShops = $preJson;

      const shops = [...vipShops, ...premiumShops];
"@

$js = [regex]::Replace(
    $js,
    '(?is)\s*const\s+shops\s*=\s*\[.*?\];',
    "`r`n$arrayBlock",
    1
)

# 2) 카드 이미지: 엑셀 이미지 사용, 비어 있으면 기존 room 이미지 사용
$oldImg = '<img src="assets/images/shops/room-${(idx % 8) + 1}.jpg" alt="${s.name} 실내 이미지">'
$newImg = '<img src="${s.image || (''assets/images/shops/room-'' + ((idx % 8) + 1) + ''.jpg'')}" alt="${s.name} 추천 이미지">'

if ($js.Contains($oldImg)) {
    $js = $js.Replace($oldImg,$newImg)
}

# 3) 가짜 조회/후기 숫자를 표시하지 않고 추천 상태로 표시
$oldStats = '<div class="stats"><span>◉ ${s.views}</span><span>♥ ${s.reviews}</span></div>'
$newStats = '<div class="stats"><span>${type}</span><span>추천 업체</span></div>'

if ($js.Contains($oldStats)) {
    $js = $js.Replace($oldStats,$newStats)
}

# 4) 기존 지역명 조건식 링크를 엑셀 상세페이지 링크로 직접 교체
$js = [regex]::Replace(
    $js,
    '(?m)^\s*<a href="\$\{s\.area\.includes\(''강남''\).*?class="detail-btn">자세히 보기</a>\s*$',
    '            <a href="${s.link || ''#''}" class="detail-btn">자세히 보기</a>',
    1
)

# 혹시 큰따옴표/공백 형태가 달라도 두 번째 방식으로 한 번 더 보정
$js = [regex]::Replace(
    $js,
    '(?m)^\s*<a href="\$\{s\.area\.includes\("강남"\).*?class="detail-btn">자세히 보기</a>\s*$',
    '            <a href="${s.link || ''#''}" class="detail-btn">자세히 보기</a>',
    1
)

# 5) VIP 4개 / 프리미엄 4개를 각각 직접 렌더링
$js = [regex]::Replace(
    $js,
    'document\.querySelector\("#vipCards"\)\.innerHTML\s*=\s*shops\.slice\(0,4\)\.map\(\(s,i\)=>card\(s,"VIP",i\)\)\.join\(""\);',
    'document.querySelector("#vipCards").innerHTML = vipShops.map((s,i)=>card(s,"VIP",i)).join("");',
    1
)

$js = [regex]::Replace(
    $js,
    'document\.querySelector\("#premiumCards"\)\.innerHTML\s*=\s*shops\.slice\(0,9\)\.map\(\(s,i\)=>card\(s,"PICK",i\+2\)\)\.join\(""\);',
    'document.querySelector("#premiumCards").innerHTML = premiumShops.map((s,i)=>card(s,"PICK",i+4)).join("");',
    1
)

if ($js -eq $before) {
    throw "app.js 구조를 찾지 못해 수정되지 않았습니다. 파일은 변경하지 않았습니다."
}

Set-Content -LiteralPath $appJs -Value $js -Encoding UTF8

# 6) 이전에 임시로 연결했던 추천용 JS가 있다면 메인에서 모두 제거
$html = Get-Content -LiteralPath $index -Raw -Encoding UTF8

$markerPatterns = @(
    '(?is)\s*<!--\s*ML-MAIN-RECOMMEND-AUTO-START\s*-->.*?<!--\s*ML-MAIN-RECOMMEND-AUTO-END\s*-->\s*',
    '(?is)\s*<!--\s*ML-MAIN-RECOMMEND-CLEANUP-START\s*-->.*?<!--\s*ML-MAIN-RECOMMEND-CLEANUP-END\s*-->\s*',
    '(?is)\s*<!--\s*ML-MAIN-RECOMMEND-FINAL-START\s*-->.*?<!--\s*ML-MAIN-RECOMMEND-FINAL-END\s*-->\s*',
    '(?is)\s*<!--\s*ML-EXISTING-RECOMMEND-PATCH-START\s*-->.*?<!--\s*ML-EXISTING-RECOMMEND-PATCH-END\s*-->\s*'
)

foreach ($p in $markerPatterns) {
    $html = [regex]::Replace($html,$p,"`r`n")
}

$html = [regex]::Replace(
    $html,
    '(?is)\s*<script\b[^>]*src=["'']assets/js/main-recommend-(?:auto|cleanup|final|existing-cards)\.js[^"'']*["''][^>]*>\s*</script>\s*',
    "`r`n"
)

# 7) app.js 캐시 버전 갱신
$html = [regex]::Replace(
    $html,
    'assets/js/app\.js(?:\?v=[^"'']*)?',
    "assets/js/app.js?v=$stamp"
)

Set-Content -LiteralPath $index -Value $html -Encoding UTF8

# 최종 검사
$checkJs = Get-Content -LiteralPath $appJs -Raw -Encoding UTF8
$checkIndex = Get-Content -LiteralPath $index -Raw -Encoding UTF8

$oldNames = '강남 골드케어|송파 밸런스룸|광교 리셋테라피|부평 온기케어|분당 슬로우케어|일산 리커버리|연수 힐링룸|마포 포근케어|안산 밸런스랩'

Write-Host ""
Write-Host "=== app.js 메인 추천 원본 직접 교체 완료 ==="
Write-Host ""
Write-Host "VIP 4개:"
$vip | ForEach-Object {
    Write-Host " -" $_.name "|" $_.area "|" $_.price
}
Write-Host ""
Write-Host "프리미엄 4개:"
$premium | ForEach-Object {
    Write-Host " -" $_.name "|" $_.area "|" $_.price
}
Write-Host ""
Write-Host "app.js 예전 임시 shops 업체명 남음:" ($checkJs -match $oldNames)
Write-Host "VIP가 vipShops 사용:" ($checkJs -match 'vipShops\.map')
Write-Host "프리미엄이 premiumShops 사용:" ($checkJs -match 'premiumShops\.map')
Write-Host "엑셀 상세링크 사용:" ($checkJs -match 's\.link')
Write-Host "엑셀 이미지 사용:" ($checkJs -match 's\.image')
Write-Host "메인 추천 임시 JS 연결 남음:" ($checkIndex -match 'main-recommend-(?:auto|cleanup|final|existing-cards)\.js')
Write-Host "app.js 캐시 버전 갱신:" ($checkIndex -match [regex]::Escape("app.js?v=$stamp"))
Write-Host ""
Write-Host "백업 폴더:" $backup
Write-Host ""
Write-Host "이제 http://127.0.0.1:5512/ 에서 Ctrl+F5 하세요."
