$ErrorActionPreference = "Stop"

$root        = "E:\mong-lounge-site"
$dataDir     = Join-Path $root "data"
$cheonanRoot = Join-Path $root "massage\other\cheonan"
$reportCsv   = Join-Path $dataDir "shop-generation-run-report.csv"
$missingCsv  = Join-Path $dataDir "shop-dong-page-missing.csv"

foreach ($p in @($reportCsv,$missingCsv)) {
    if (!(Test-Path -LiteralPath $p)) { throw "필수 파일을 찾을 수 없습니다: $p" }
}

$stamp = Get-Date -Format "yyyyMMdd-HHmmss"
$backupRoot = "E:\mong-lounge-cheonan9-backup-$stamp"
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
    $t = [regex]::Replace([string]$html,'(?is)<[^>]+>',' ')
    $t = [System.Net.WebUtility]::HtmlDecode($t)
    return ([regex]::Replace($t,'\s+',' ').Trim())
}

function Resolve-Href([string]$sourcePage,[string]$href) {
    if ([string]::IsNullOrWhiteSpace($href)) { return $null }
    $h = [System.Net.WebUtility]::HtmlDecode($href.Trim())
    $h = ($h -split '#')[0]
    $h = ($h -split '\?')[0]
    if ([string]::IsNullOrWhiteSpace($h)) { return $null }

    try {
        if ($h.StartsWith("/")) {
            $local = [uri]::UnescapeDataString($h).TrimStart("/")
            $candidate = Join-Path $root ($local -replace "/","\")
        } else {
            $local = [uri]::UnescapeDataString($h) -replace "/","\"
            $candidate = [System.IO.Path]::GetFullPath((Join-Path (Split-Path -Parent $sourcePage) $local))
        }

        if (Test-Path -LiteralPath $candidate -PathType Container) {
            $candidate = Join-Path $candidate "index.html"
        } elseif (!(Test-Path -LiteralPath $candidate -PathType Leaf)) {
            $c2 = Join-Path $candidate "index.html"
            if (Test-Path -LiteralPath $c2 -PathType Leaf) { $candidate = $c2 }
        }

        if (Test-Path -LiteralPath $candidate -PathType Leaf) {
            return [System.IO.Path]::GetFullPath($candidate)
        }
    } catch {}

    return $null
}

function Find-TargetPage([string]$area) {
    $hits = New-Object System.Collections.Generic.List[object]

    foreach ($f in @(Get-ChildItem -LiteralPath $cheonanRoot -Recurse -File -Filter "index.html")) {
        $html = Get-Content -LiteralPath $f.FullName -Raw -Encoding UTF8

        foreach ($m in [regex]::Matches($html,'(?is)<a\b[^>]*href=["'']([^"'']+)["''][^>]*>(.*?)</a>')) {
            $text = Strip-Html $m.Groups[2].Value
            if ($text -ne $area -and $text -notlike "*$area*") { continue }

            $target = Resolve-Href $f.FullName $m.Groups[1].Value
            if ([string]::IsNullOrWhiteSpace($target)) { continue }

            $score = if ($text -eq $area) { 100 } else { 50 }
            $hits.Add([pscustomobject]@{
                Source=$f.FullName
                Text=$text
                Target=$target
                Score=$score
            })
        }
    }

    if ($hits.Count -eq 0) { return $null }

    return ($hits | Sort-Object Score -Descending | Select-Object -First 1)
}

function Build-ShopBlock($rows,[string]$label) {
    $groups = @($rows | Group-Object 업체명 | Sort-Object Name)

    $cards = foreach ($g in $groups) {
        $rep = $g.Group | Select-Object -First 1
        $detailUrl = Get-DetailUrl ([string]$rep.상세페이지)

@"
<a class="ml-top-shop-card" href="$(HtmlEncode $detailUrl)">
  <div class="ml-top-shop-copy">$(HtmlEncode $rep.상위지역) · $(HtmlEncode $label) 중심 빠른 예약과 편안한 방문 관리</div>
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
.ml-shop-jump{position:fixed;right:18px;bottom:20px;z-index:80;background:#083f35;color:#ffc928!important;border:1px solid rgba(255,201,40,.55);border-radius:999px;padding:12px 16px;font-size:13px;font-weight:900;text-decoration:none;box-shadow:0 7px 24px rgba(0,0,0,.18)}
@media(max-width:700px){.ml-top-shop-grid{grid-template-columns:1fr}.ml-shop-jump{right:12px;bottom:14px}}
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

function Apply-Block([string]$page,[string]$block) {
    $html = Get-Content -LiteralPath $page -Raw -Encoding UTF8
    $original = $html

    $html = [regex]::Replace(
        $html,
        '(?is)\s*<!-- ML-TOP-SHOP-START -->.*?<!-- ML-TOP-SHOP-END -->\s*',
        "`r`n"
    )

    $registeredPattern = '(?is)<section\b[^>]*>.*?REGISTERED SHOP.*?</section>'
    $m = [regex]::Match($html,$registeredPattern)

    if ($m.Success) {
        $html = $html.Substring(0,$m.Index) + $block + $html.Substring($m.Index + $m.Length)
    } else {
        $main = [regex]::Match($html,'(?is)<main\b[^>]*>')
        if ($main.Success) {
            $at = $main.Index + $main.Length
            $html = $html.Substring(0,$at) + "`r`n$block`r`n" + $html.Substring($at)
        } elseif ($html -match '(?is)</body>') {
            $html = [regex]::Replace($html,'(?is)</body>',"$block`r`n</body>",1)
        } else {
            return $false
        }
    }

    if ($html -ne $original) {
        Backup-File $page
        Set-Content -LiteralPath $page -Value $html -Encoding UTF8
    }

    return $true
}

$missing = @(Import-Csv -LiteralPath $missingCsv)
$report  = @(Import-Csv -LiteralPath $reportCsv | Where-Object { [string]$_.상태 -notlike "실패*" })

Write-Host ""
Write-Host "=== 천안 누락 9개 페이지 연결 ==="
Write-Host "대상:" $missing.Count

$found = 0
$modified = 0
$stillMissing = New-Object System.Collections.Generic.List[object]
$mapped = New-Object System.Collections.Generic.List[object]

foreach ($m in $missing) {
    $area = [string]$m.하위지역
    $hit = Find-TargetPage $area

    if ($null -eq $hit) {
        $stillMissing.Add($m)
        continue
    }

    $found++

    $rows = @(
        $report | Where-Object {
            [string]$_.광역지역 -eq [string]$m.광역지역 -and
            [string]$_.상위지역 -eq [string]$m.상위지역 -and
            [string]$_.하위지역 -eq $area
        }
    )

    if ($rows.Count -eq 0) {
        $stillMissing.Add($m)
        continue
    }

    $block = Build-ShopBlock $rows $area

    if (Apply-Block ([string]$hit.Target) $block) {
        $modified++
        $mapped.Add([pscustomobject]@{
            하위지역=$area
            링크출처=$hit.Source
            실제페이지=$hit.Target
            업체수=@($rows | Select-Object -ExpandProperty 업체명 -Unique).Count
        })
    } else {
        $stillMissing.Add($m)
    }
}

$mapCsv = Join-Path $dataDir "shop-cheonan9-page-map.csv"
$mapped | Export-Csv -LiteralPath $mapCsv -NoTypeInformation -Encoding UTF8

$remainCsv = Join-Path $dataDir "shop-dong-page-missing-after-cheonan9.csv"
$stillMissing | Export-Csv -LiteralPath $remainCsv -NoTypeInformation -Encoding UTF8

$blockOk = 0
$jumpOk = 0
$broken = 0

foreach ($x in $mapped) {
    $p = [string]$x.실제페이지
    $html = Get-Content -LiteralPath $p -Raw -Encoding UTF8

    if ($html -match '<!-- ML-TOP-SHOP-START -->') { $blockOk++ }
    if ($html -match 'class="ml-shop-jump"') { $jumpOk++ }

    foreach ($mm in [regex]::Matches($html,'href="(/shops/[^"]+/index\.html)"')) {
        $target = Join-Path $root ($mm.Groups[1].Value.TrimStart("/") -replace "/","\")
        if (!(Test-Path -LiteralPath $target)) { $broken++ }
    }
}

Write-Host ""
Write-Host "=== 천안 9개 보정 완료 ==="
Write-Host "찾은 페이지:" $found
Write-Host "수정 페이지:" $modified
Write-Host "여전히 미발견:" $stillMissing.Count
Write-Host ""
Write-Host "업체영역 확인:" "$blockOk / $modified"
Write-Host "업체 바로가기 확인:" "$jumpOk / $modified"
Write-Host "깨진 상세링크:" $broken
Write-Host ""
Write-Host "연결표:" $mapCsv
Write-Host "백업:" $backupRoot
Write-Host "상세페이지 수정: 0"
