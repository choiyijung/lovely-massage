$ErrorActionPreference = "Stop"

$root       = "E:\mong-lounge-site"
$dataDir    = Join-Path $root "data"
$reportCsv  = Join-Path $dataDir "shop-generation-run-report.csv"
$generator  = Join-Path $dataDir "mong-lounge-generate-all-shops.ps1"

if (!(Test-Path -LiteralPath $reportCsv)) {
    throw "실행 리포트를 찾을 수 없습니다: $reportCsv"
}
if (!(Test-Path -LiteralPath $generator)) {
    throw "자동생성기를 찾을 수 없습니다: $generator"
}

$stamp = Get-Date -Format "yyyyMMdd-HHmmss"
$backupRoot = "E:\mong-lounge-region-shop-layout-backup-$stamp"
New-Item -ItemType Directory -Force -Path $backupRoot | Out-Null

function HtmlEncode([string]$text) {
    if ($null -eq $text) { return "" }
    return [System.Net.WebUtility]::HtmlEncode($text)
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
    New-Item -ItemType Directory -Force -Path (Split-Path -Parent $dest) | Out-Null
    Copy-Item -LiteralPath $filePath -Destination $dest -Force
}

function Get-DetailUrl([string]$detailFile) {
    $rel = Get-RelativeToRoot $detailFile
    return "/" + ($rel -replace "\\","/")
}

function Build-TopShopBlock($rowsForPage) {
    $parent = [string]($rowsForPage | Select-Object -First 1).상위지역
    $groups = @($rowsForPage | Group-Object 업체명 | Sort-Object Name)

    $cards = foreach ($g in $groups) {
        $rows = @($g.Group | Sort-Object 하위지역)
        $rep = $rows | Select-Object -First 1
        $detailUrl = Get-DetailUrl ([string]$rep.상세페이지)

        $areas = @($rows | Select-Object -ExpandProperty 하위지역 -Unique)
        $areaText = if ($areas.Count -le 3) {
            ($areas -join " · ")
        } else {
            "$($areas[0]) · $($areas[1]) · $($areas[2]) 외 $($areas.Count - 3)개 지역"
        }

@"
<a class="ml-top-shop-card" href="$(HtmlEncode $detailUrl)">
  <div class="ml-top-shop-copy">$(HtmlEncode $areaText) 중심 빠른 예약과 편안한 방문 관리</div>
  <div class="ml-top-shop-line"></div>
  <div class="ml-top-shop-name">$(HtmlEncode $g.Name)</div>
</a>
"@
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
@media(max-width:700px){
  .ml-top-shop-grid{grid-template-columns:1fr}
  .ml-top-shop-title{font-size:26px}
  .ml-top-shop-card{min-height:170px}
}
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

# ---------------------------------------------------------
# 1) 현재 120개 지역페이지 수정
# ---------------------------------------------------------

$report = @(Import-Csv -LiteralPath $reportCsv)

$byPage = @{}
foreach ($g in ($report | Where-Object { $_.상태 -notlike "실패*" } | Group-Object 지역페이지)) {
    $byPage[$g.Name] = @($g.Group)
}

Write-Host ""
Write-Host "=== 현재 지역페이지 업체영역 수정 ==="
Write-Host "대상 지역페이지:" $byPage.Count

$modified = 0
$replacedOld = 0
$fallbackTop = 0
$removedBottom = 0
$failed = New-Object System.Collections.Generic.List[object]

foreach ($page in $byPage.Keys) {
    if (!(Test-Path -LiteralPath $page)) {
        $failed.Add([pscustomobject]@{페이지=$page;원인="파일없음"})
        continue
    }

    $html = Get-Content -LiteralPath $page -Raw -Encoding UTF8
    $original = $html
    $block = Build-TopShopBlock $byPage[$page]

    # 기존 상단 자동영역이 있다면 먼저 제거
    $html = [regex]::Replace(
        $html,
        '(?is)\s*<!-- ML-TOP-SHOP-START -->.*?<!-- ML-TOP-SHOP-END -->\s*',
        "`r`n"
    )

    # 현재 하단 흰색 자동카드 영역 제거
    $beforeBottom = $html
    $html = [regex]::Replace(
        $html,
        '(?is)\s*<!-- ML-SHOP-CARDS-START -->.*?<!-- ML-SHOP-CARDS-END -->\s*',
        "`r`n"
    )
    if ($html -ne $beforeBottom) { $removedBottom++ }

    # 이전 1개짜리 REGISTERED SHOP 섹션을 통째로 교체
    # 대부분의 지역페이지는 이 구조이며, 없을 때만 main 상단에 삽입
    $registeredPattern = '(?is)<section\b[^>]*>.*?REGISTERED SHOP.*?</section>'
    $m = [regex]::Match($html,$registeredPattern)

    if ($m.Success) {
        $html = $html.Substring(0,$m.Index) + $block + $html.Substring($m.Index + $m.Length)
        $replacedOld++
    }
    else {
        $mainMatch = [regex]::Match($html,'(?is)<main\b[^>]*>')
        if ($mainMatch.Success) {
            $insertAt = $mainMatch.Index + $mainMatch.Length
            $html = $html.Substring(0,$insertAt) + "`r`n$block`r`n" + $html.Substring($insertAt)
            $fallbackTop++
        }
        else {
            $failed.Add([pscustomobject]@{페이지=$page;원인="REGISTERED SHOP 및 main 없음"})
            continue
        }
    }

    if ($html -ne $original) {
        Backup-File $page
        Set-Content -LiteralPath $page -Value $html -Encoding UTF8
        $modified++
    }
}

# ---------------------------------------------------------
# 2) 앞으로 자동생성기도 같은 상단 디자인으로 고정
# ---------------------------------------------------------

Write-Host ""
Write-Host "=== 자동생성기 규칙 수정 ==="

Backup-File $generator
$gen = Get-Content -LiteralPath $generator -Raw -Encoding UTF8

$newFunction = @'
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

'@

$fnPattern = '(?s)function Build-RegionCardBlock\(\$rowsForPage\) \{.*?\r?\n\}\r?\n\r?\n(?=\$rowsByRegionPage = @\{\})'
$fnMatch = [regex]::Match($gen,$fnPattern)

if (!$fnMatch.Success) {
    throw "자동생성기의 Build-RegionCardBlock 함수를 찾지 못했습니다. 지역페이지 수정은 완료됐지만 생성기 패치는 중단합니다."
}

$gen = $gen.Substring(0,$fnMatch.Index) + $newFunction + $gen.Substring($fnMatch.Index + $fnMatch.Length)

$oldInsertPattern = '(?s)    if \(\$rowsByRegionPage\.ContainsKey\(\$page\)\) \{.*?\r?\n    \}\r?\n\r?\n    if \(\$html -ne \$original\) \{'
$insertMatch = [regex]::Match($gen,$oldInsertPattern)

if (!$insertMatch.Success) {
    throw "자동생성기의 지역카드 삽입 구문을 찾지 못했습니다. 지역페이지 수정은 완료됐지만 생성기 삽입위치 패치는 중단합니다."
}

$newInsert = @'
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
'@

$gen = $gen.Substring(0,$insertMatch.Index) + $newInsert + $gen.Substring($insertMatch.Index + $insertMatch.Length)

Set-Content -LiteralPath $generator -Value $gen -Encoding UTF8

# ---------------------------------------------------------
# 3) 검증
# ---------------------------------------------------------

$topOk = 0
$bottomRemain = 0
$links = 0
$broken = 0

foreach ($page in $byPage.Keys) {
    if (!(Test-Path -LiteralPath $page)) { continue }

    $html = Get-Content -LiteralPath $page -Raw -Encoding UTF8
    if ($html -match '<!-- ML-TOP-SHOP-START -->') { $topOk++ }
    if ($html -match '<!-- ML-SHOP-CARDS-START -->') { $bottomRemain++ }

    foreach ($m in [regex]::Matches($html,'href="(/shops/[^"]+/index\.html)"')) {
        $links++
        $target = Join-Path $root ($m.Groups[1].Value.TrimStart("/") -replace "/","\")
        if (!(Test-Path -LiteralPath $target)) { $broken++ }
    }
}

$generatorTop = (Get-Content -LiteralPath $generator -Raw -Encoding UTF8) -match 'ML-TOP-SHOP-START'

Write-Host ""
Write-Host "=== 업체영역 위치/디자인 수정 완료 ==="
Write-Host "대상 지역페이지:" $byPage.Count
Write-Host "수정 페이지:" $modified
Write-Host "기존 REGISTERED SHOP 교체:" $replacedOld
Write-Host "main 상단 대체삽입:" $fallbackTop
Write-Host "하단 흰색 자동영역 제거:" $removedBottom
Write-Host "수정 실패:" $failed.Count
Write-Host ""
Write-Host "상단 녹색 업체영역 확인:" "$topOk / $($byPage.Count)"
Write-Host "하단 흰색 영역 남음:" $bottomRemain
Write-Host "상세보기 링크:" $links
Write-Host "깨진 상세링크:" $broken
Write-Host "자동생성기 상단방식 적용:" $generatorTop
Write-Host ""
Write-Host "백업 폴더:" $backupRoot

if ($failed.Count -gt 0) {
    $failCsv = Join-Path $dataDir "region-shop-layout-failed.csv"
    $failed | Export-Csv -LiteralPath $failCsv -NoTypeInformation -Encoding UTF8
    Write-Host "확인 필요 목록:" $failCsv
}

Write-Host ""
Write-Host "상세페이지 수정: 0"
