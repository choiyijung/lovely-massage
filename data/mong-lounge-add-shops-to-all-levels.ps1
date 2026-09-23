$ErrorActionPreference = "Stop"

# ============================================================
# 몽라운지 업체영역 전체 계층 적용
# 1) 서울/경기/인천 등 상위 지역 페이지에도 업체 표시
# 2) 구/시 지역 페이지 업체 표시 유지
# 3) 동/읍/면 페이지에도 해당 업체 표시
# 4) 오른쪽 하단 "업체 바로가기" 버튼 추가
# 5) 기존 상세페이지 디자인은 수정하지 않음
# ============================================================

$root       = "E:\mong-lounge-site"
$dataDir    = Join-Path $root "data"
$massageRoot= Join-Path $root "massage"
$reportCsv  = Join-Path $dataDir "shop-generation-run-report.csv"

if (!(Test-Path -LiteralPath $reportCsv)) {
    throw "실행 리포트를 찾을 수 없습니다: $reportCsv"
}

$stamp = Get-Date -Format "yyyyMMdd-HHmmss"
$backupRoot = "E:\mong-lounge-all-level-shop-backup-$stamp"
$missingCsv = Join-Path $dataDir "shop-dong-page-missing.csv"

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
    $dst = Join-Path $backupRoot $rel
    New-Item -ItemType Directory -Force -Path (Split-Path -Parent $dst) | Out-Null
    Copy-Item -LiteralPath $filePath -Destination $dst -Force
}

function Get-DetailUrl([string]$detailFile) {
    $rel = Get-RelativeToRoot $detailFile
    return "/" + ($rel -replace "\\","/")
}

function Strip-Html([string]$html) {
    if ([string]::IsNullOrWhiteSpace($html)) { return "" }

    $t = [regex]::Replace($html,'(?is)<[^>]+>',' ')
    $t = [System.Net.WebUtility]::HtmlDecode($t)
    return ([regex]::Replace($t,'\s+',' ').Trim())
}

function Resolve-Href([string]$parentPage,[string]$href) {
    if ([string]::IsNullOrWhiteSpace($href)) { return $null }

    $h = [System.Net.WebUtility]::HtmlDecode($href.Trim())

    # query / hash 제거
    $h = ($h -split '#')[0]
    $h = ($h -split '\?')[0]

    if ([string]::IsNullOrWhiteSpace($h)) { return $null }
    if ($h -match '^(mailto:|tel:|sms:|javascript:)') { return $null }

    try {
        if ($h -match '^https?://') {
            $u = [uri]$h
            $local = [uri]::UnescapeDataString($u.AbsolutePath).TrimStart("/")
            $candidate = Join-Path $root ($local -replace "/","\")
        }
        elseif ($h.StartsWith("/")) {
            $local = [uri]::UnescapeDataString($h).TrimStart("/")
            $candidate = Join-Path $root ($local -replace "/","\")
        }
        else {
            $local = [uri]::UnescapeDataString($h) -replace "/","\"
            $baseDir = Split-Path -Parent $parentPage
            $candidate = [System.IO.Path]::GetFullPath((Join-Path $baseDir $local))
        }
    }
    catch {
        return $null
    }

    if (Test-Path -LiteralPath $candidate -PathType Container) {
        $candidate = Join-Path $candidate "index.html"
    }
    elseif (!(Test-Path -LiteralPath $candidate -PathType Leaf)) {
        $candidate2 = Join-Path $candidate "index.html"
        if (Test-Path -LiteralPath $candidate2 -PathType Leaf) {
            $candidate = $candidate2
        }
    }

    if (!(Test-Path -LiteralPath $candidate -PathType Leaf)) {
        return $null
    }

    $full = [System.IO.Path]::GetFullPath($candidate)
    $massageFull = [System.IO.Path]::GetFullPath($massageRoot)

    if (!$full.StartsWith($massageFull,[System.StringComparison]::OrdinalIgnoreCase)) {
        return $null
    }

    return $full
}

function Get-LinkIndex([string]$parentPage) {
    $map = @{}

    if (!(Test-Path -LiteralPath $parentPage)) { return $map }

    $html = Get-Content -LiteralPath $parentPage -Raw -Encoding UTF8

    foreach ($m in [regex]::Matches(
        $html,
        '(?is)<a\b[^>]*href\s*=\s*["'']([^"'']+)["''][^>]*>(.*?)</a>'
    )) {
        $href = $m.Groups[1].Value
        $text = Strip-Html $m.Groups[2].Value

        if ([string]::IsNullOrWhiteSpace($text)) { continue }

        $target = Resolve-Href $parentPage $href
        if ([string]::IsNullOrWhiteSpace($target)) { continue }

        if (!$map.ContainsKey($text)) {
            $map[$text] = New-Object System.Collections.Generic.List[string]
        }
        if (!$map[$text].Contains($target)) {
            $map[$text].Add($target)
        }
    }

    return $map
}

function Find-DongPage([hashtable]$linkIndex,[string]$area) {
    if ([string]::IsNullOrWhiteSpace($area)) { return $null }

    # 1순위: 링크 텍스트 정확히 일치
    if ($linkIndex.ContainsKey($area)) {
        return [string]($linkIndex[$area] | Select-Object -First 1)
    }

    # 2순위: "도곡동 지역 안내"처럼 하위지역명이 포함된 링크
    $candidates = @(
        foreach ($k in $linkIndex.Keys) {
            if ($k -eq $area -or $k -like "*$area*") {
                foreach ($p in $linkIndex[$k]) {
                    [pscustomobject]@{
                        Text=$k
                        Path=$p
                        Score=if($k -eq $area){100}else{50}
                    }
                }
            }
        }
    )

    if ($candidates.Count -gt 0) {
        return [string](
            $candidates |
            Sort-Object Score -Descending |
            Select-Object -First 1 -ExpandProperty Path
        )
    }

    return $null
}

function Build-ShopBlock($rows,[string]$label,[string]$scopeMode) {
    $groups = @($rows | Group-Object 업체명 | Sort-Object Name)

    $cards = foreach ($g in $groups) {
        $vendorRows = @($g.Group | Sort-Object 상위지역,하위지역)
        $rep = $vendorRows | Select-Object -First 1
        $detailUrl = Get-DetailUrl ([string]$rep.상세페이지)

        if ($scopeMode -eq "dong") {
            $copy = "$($rep.상위지역) · $($rep.하위지역) 중심 빠른 예약과 편안한 방문 관리"
        }
        elseif ($scopeMode -eq "parent") {
            $areas = @($vendorRows | Select-Object -ExpandProperty 하위지역 -Unique)
            if ($areas.Count -le 3) {
                $areaText = $areas -join " · "
            } else {
                $areaText = "$($areas[0]) · $($areas[1]) · $($areas[2]) 외 $($areas.Count - 3)개 지역"
            }
            $copy = "$areaText 중심 빠른 예약과 편안한 방문 관리"
        }
        else {
            $parents = @($vendorRows | Select-Object -ExpandProperty 상위지역 -Unique)
            if ($parents.Count -le 3) {
                $parentText = $parents -join " · "
            } else {
                $parentText = "$($parents[0]) · $($parents[1]) · $($parents[2]) 외 $($parents.Count - 3)개 지역"
            }
            $copy = "$parentText 지역에서 확인 가능한 업체 안내"
        }

@"
<a class="ml-top-shop-card" href="$(HtmlEncode $detailUrl)">
  <div class="ml-top-shop-copy">$(HtmlEncode $copy)</div>
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
.ml-shop-jump{position:fixed;right:18px;bottom:20px;z-index:80;background:#083f35;color:#ffc928!important;border:1px solid rgba(255,201,40,.55);border-radius:999px;padding:12px 16px;font-size:13px;font-weight:900;text-decoration:none;box-shadow:0 7px 24px rgba(0,0,0,.18)}
@media(max-width:700px){
  .ml-top-shop-grid{grid-template-columns:1fr}
  .ml-top-shop-title{font-size:26px}
  .ml-top-shop-card{min-height:170px}
  .ml-shop-jump{right:12px;bottom:14px;padding:10px 13px}
}
</style>
<section id="registered-shop" class="ml-top-shop-section" aria-label="등록 업체">
  <div class="ml-top-shop-head">
    <div>
      <p class="ml-top-shop-kicker">REGISTERED SHOP</p>
      <h2 class="ml-top-shop-title">$(HtmlEncode $label) 등록 업체</h2>
    </div>
    <div class="ml-top-shop-count">$($groups.Count)개 업체</div>
  </div>
  <div class="ml-top-shop-grid">
    $($cards -join "`r`n    ")
  </div>
</section>
<a class="ml-shop-jump" href="#registered-shop">업체 바로가기</a>
<!-- ML-TOP-SHOP-END -->
"@
}

function Apply-ShopBlock([string]$page,[string]$block) {
    if (!(Test-Path -LiteralPath $page)) {
        return $false
    }

    $html = Get-Content -LiteralPath $page -Raw -Encoding UTF8
    $original = $html

    # 기존 자동 업체 블록 제거
    $html = [regex]::Replace(
        $html,
        '(?is)\s*<!-- ML-TOP-SHOP-START -->.*?<!-- ML-TOP-SHOP-END -->\s*',
        "`r`n"
    )

    # 과거 하단 자동 카드가 남아 있으면 제거
    $html = [regex]::Replace(
        $html,
        '(?is)\s*<!-- ML-SHOP-CARDS-START -->.*?<!-- ML-SHOP-CARDS-END -->\s*',
        "`r`n"
    )

    # 예전 REGISTERED SHOP 섹션이 있다면 그 자리에 교체
    $registeredPattern = '(?is)<section\b[^>]*>.*?REGISTERED SHOP.*?</section>'
    $registeredMatch = [regex]::Match($html,$registeredPattern)

    if ($registeredMatch.Success) {
        $html = $html.Substring(0,$registeredMatch.Index) +
                $block +
                $html.Substring($registeredMatch.Index + $registeredMatch.Length)
    }
    else {
        # 없으면 main 시작 바로 뒤에 삽입
        $mainMatch = [regex]::Match($html,'(?is)<main\b[^>]*>')
        if ($mainMatch.Success) {
            $at = $mainMatch.Index + $mainMatch.Length
            $html = $html.Substring(0,$at) +
                    "`r`n$block`r`n" +
                    $html.Substring($at)
        }
        elseif ($html -match '(?is)</body>') {
            $html = [regex]::Replace(
                $html,
                '(?is)</body>',
                "$block`r`n</body>",
                1
            )
        }
        else {
            return $false
        }
    }

    if ($html -ne $original) {
        Backup-File $page
        Set-Content -LiteralPath $page -Value $html -Encoding UTF8
    }

    return $true
}

# ------------------------------------------------------------
# 실행 데이터
# ------------------------------------------------------------

$report = @(
    Import-Csv -LiteralPath $reportCsv |
    Where-Object { [string]$_.상태 -notlike "실패*" }
)

Write-Host ""
Write-Host "=== 업체영역 전체 계층 적용 시작 ==="
Write-Host "업체-지역 데이터:" $report.Count

$pageRows = @{}
$pageLabels = @{}
$pageModes = @{}

function Add-PageRows([string]$page,$rows,[string]$label,[string]$mode) {
    if ([string]::IsNullOrWhiteSpace($page)) { return }
    if (!(Test-Path -LiteralPath $page)) { return }

    if (!$pageRows.ContainsKey($page)) {
        $pageRows[$page] = New-Object System.Collections.Generic.List[object]
    }

    foreach ($r in @($rows)) {
        $pageRows[$page].Add($r)
    }

    $pageLabels[$page] = $label
    $pageModes[$page] = $mode
}

# ------------------------------------------------------------
# A. 기존 구/시 지역 페이지
# ------------------------------------------------------------

$parentGroups = @($report | Group-Object 지역페이지)

foreach ($g in $parentGroups) {
    $rows = @($g.Group)
    $page = [string]$g.Name
    $label = [string]($rows | Select-Object -First 1).상위지역

    Add-PageRows $page $rows $label "parent"
}

# ------------------------------------------------------------
# B. 동/읍/면 페이지 찾기
#    각 구/시 페이지의 실제 링크를 따라가므로 동명이인 지역 혼동 방지
# ------------------------------------------------------------

$missing = New-Object System.Collections.Generic.List[object]
$dongFound = 0

foreach ($g in $parentGroups) {
    $parentPage = [string]$g.Name
    if (!(Test-Path -LiteralPath $parentPage)) { continue }

    $linkIndex = Get-LinkIndex $parentPage

    foreach ($areaGroup in @($g.Group | Group-Object 하위지역)) {
        $area = [string]$areaGroup.Name
        $rows = @($areaGroup.Group)

        $dongPage = Find-DongPage $linkIndex $area

        if ([string]::IsNullOrWhiteSpace($dongPage)) {
            $sample = $rows | Select-Object -First 1
            $missing.Add([pscustomobject]@{
                광역지역=[string]$sample.광역지역
                상위지역=[string]$sample.상위지역
                하위지역=$area
                부모페이지=$parentPage
            })
            continue
        }

        Add-PageRows $dongPage $rows $area "dong"
        $dongFound++
    }
}

# ------------------------------------------------------------
# C. 서울/경기/인천 상위 메인에도 업체 표시
# ------------------------------------------------------------

$metroRoots = [ordered]@{
    "서울" = Join-Path $massageRoot "seoul\index.html"
    "경기" = Join-Path $massageRoot "gyeonggi\index.html"
    "인천" = Join-Path $massageRoot "incheon\index.html"
}

foreach ($geo in $metroRoots.Keys) {
    $page = $metroRoots[$geo]

    if (Test-Path -LiteralPath $page) {
        $rows = @($report | Where-Object { [string]$_.광역지역 -eq $geo })
        if ($rows.Count -gt 0) {
            Add-PageRows $page $rows $geo "root"
        }
    }
}

# ------------------------------------------------------------
# D. 기타 광역시 메인 페이지가 직접 있으면 추가
# ------------------------------------------------------------

$otherRoots = [ordered]@{
    "대전"       = Join-Path $massageRoot "other\daejeon\index.html"
    "대구"       = Join-Path $massageRoot "other\daegu\index.html"
    "광주광역시" = Join-Path $massageRoot "other\gwangju\index.html"
}

foreach ($geo in $otherRoots.Keys) {
    $page = $otherRoots[$geo]
    if (Test-Path -LiteralPath $page) {
        $rows = @($report | Where-Object { [string]$_.광역지역 -eq $geo })
        if ($rows.Count -gt 0) {
            Add-PageRows $page $rows $geo "root"
        }
    }
}

# 각 페이지 내부 중복 행 제거
foreach ($page in @($pageRows.Keys)) {
    $pageRows[$page] = @(
        $pageRows[$page] |
        Sort-Object 업체명,광역지역,상위지역,하위지역,상세페이지 -Unique
    )
}

Write-Host "구/시 지역페이지:" $parentGroups.Count
Write-Host "동/읍/면 페이지 발견:" $dongFound
Write-Host "동/읍/면 페이지 미발견:" $missing.Count
Write-Host "최종 업체영역 적용 페이지:" $pageRows.Count

if ($missing.Count -gt 0) {
    $missing |
        Sort-Object 광역지역,상위지역,하위지역 -Unique |
        Export-Csv -LiteralPath $missingCsv -NoTypeInformation -Encoding UTF8
}

# ------------------------------------------------------------
# 실제 수정
# ------------------------------------------------------------

$modified = 0
$failed = 0
$i = 0

foreach ($page in @($pageRows.Keys)) {
    $i++

    $block = Build-ShopBlock `
        $pageRows[$page] `
        ([string]$pageLabels[$page]) `
        ([string]$pageModes[$page])

    if (Apply-ShopBlock $page $block) {
        $modified++
    }
    else {
        $failed++
    }

    if (($i % 100) -eq 0 -or $i -eq $pageRows.Count) {
        Write-Host "진행: $i / $($pageRows.Count)"
    }
}

# ------------------------------------------------------------
# 검증
# ------------------------------------------------------------

$blockOk = 0
$jumpOk = 0
$brokenDetail = 0
$detailLinks = 0
$seoulRootOk = $false
$dongBlockOk = 0

foreach ($page in @($pageRows.Keys)) {
    if (!(Test-Path -LiteralPath $page)) { continue }

    $html = Get-Content -LiteralPath $page -Raw -Encoding UTF8

    if ($html -match '<!-- ML-TOP-SHOP-START -->') {
        $blockOk++
    }

    if ($html -match 'class="ml-shop-jump"') {
        $jumpOk++
    }

    if ($pageModes[$page] -eq "dong" -and
        $html -match '<!-- ML-TOP-SHOP-START -->') {
        $dongBlockOk++
    }

    foreach ($m in [regex]::Matches(
        $html,
        'href="(/shops/[^"]+/index\.html)"'
    )) {
        $detailLinks++
        $target = Join-Path $root (
            $m.Groups[1].Value.TrimStart("/") -replace "/","\"
        )
        if (!(Test-Path -LiteralPath $target)) {
            $brokenDetail++
        }
    }
}

$seoulRoot = $metroRoots["서울"]
if (Test-Path -LiteralPath $seoulRoot) {
    $seoulHtml = Get-Content -LiteralPath $seoulRoot -Raw -Encoding UTF8
    $seoulRootOk = (
        $seoulHtml -match '<!-- ML-TOP-SHOP-START -->' -and
        $seoulHtml -match '서울 등록 업체'
    )
}

Write-Host ""
Write-Host "=== 업체영역 전체 계층 적용 완료 ==="
Write-Host "수정 페이지:" $modified
Write-Host "수정 실패:" $failed
Write-Host ""
Write-Host "업체영역 확인:" "$blockOk / $($pageRows.Count)"
Write-Host "오른쪽 하단 바로가기 확인:" "$jumpOk / $($pageRows.Count)"
Write-Host "동/읍/면 업체영역 확인:" "$dongBlockOk / $dongFound"
Write-Host "서울 메인 업체영역:" $seoulRootOk
Write-Host ""
Write-Host "상세페이지 링크:" $detailLinks
Write-Host "깨진 상세링크:" $brokenDetail
Write-Host ""
Write-Host "동/읍/면 페이지 미발견:" $missing.Count
if ($missing.Count -gt 0) {
    Write-Host "확인 목록:" $missingCsv
}
Write-Host ""
Write-Host "백업 폴더:" $backupRoot
Write-Host "상세페이지 수정: 0"

# ML-VISIBLE-SHOP-BUTTON-FUTURE
$mlVisibleButtonScript = Join-Path $dataDir "mong-lounge-make-shop-button-visible.ps1"
if (Test-Path -LiteralPath $mlVisibleButtonScript) {
    & $mlVisibleButtonScript -FromGenerator
}
