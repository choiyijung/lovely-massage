$ErrorActionPreference = "Stop"

$root = "E:\mong-lounge-site"
$massageRoot = Join-Path $root "massage"
$stamp = Get-Date -Format "yyyyMMdd-HHmmss"
$backupRoot = "E:\mong-lounge-shop-jump-backup-$stamp"

New-Item -ItemType Directory -Force -Path $backupRoot | Out-Null

function Get-RelativeToRoot([string]$fullPath) {
    $full = [System.IO.Path]::GetFullPath($fullPath)
    $base = [System.IO.Path]::GetFullPath($root)
    if ($full.StartsWith($base,[System.StringComparison]::OrdinalIgnoreCase)) {
        return $full.Substring($base.Length).TrimStart("\")
    }
    return [System.IO.Path]::GetFileName($full)
}

function Backup-File([string]$filePath) {
    $rel = Get-RelativeToRoot $filePath
    $dst = Join-Path $backupRoot $rel
    New-Item -ItemType Directory -Force -Path (Split-Path -Parent $dst) | Out-Null
    Copy-Item -LiteralPath $filePath -Destination $dst -Force
}

$pages = @(
    Get-ChildItem -LiteralPath $massageRoot -Recurse -File -Filter "index.html" |
    Where-Object {
        try {
            (Get-Content -LiteralPath $_.FullName -Raw -Encoding UTF8) -match '<!-- ML-TOP-SHOP-START -->'
        } catch { $false }
    }
)

$modified = 0
$failed = 0

$jump = @'
<!-- ML-SHOP-JUMP-START -->
<a href="#registered-shop" class="ml-shop-jump-fixed" style="position:fixed!important;right:18px!important;bottom:20px!important;z-index:99999!important;display:flex!important;align-items:center!important;justify-content:center!important;min-width:118px!important;height:46px!important;padding:0 18px!important;border-radius:999px!important;background:#083f35!important;color:#ffc928!important;border:1px solid rgba(255,201,40,.70)!important;box-shadow:0 8px 28px rgba(0,0,0,.28)!important;font-size:14px!important;font-weight:900!important;line-height:1!important;text-decoration:none!important;">업체 바로가기</a>
<!-- ML-SHOP-JUMP-END -->
'@

foreach ($f in $pages) {
    try {
        $html = Get-Content -LiteralPath $f.FullName -Raw -Encoding UTF8
        $original = $html

        # 기존 블록 안의 바로가기 링크 제거
        $html = [regex]::Replace(
            $html,
            '(?is)\s*<a\b[^>]*class=["''][^"'']*ml-shop-jump[^"'']*["''][^>]*>.*?</a>\s*',
            "`r`n"
        )

        # 과거 별도 고정버튼 제거
        $html = [regex]::Replace(
            $html,
            '(?is)\s*<!-- ML-SHOP-JUMP-START -->.*?<!-- ML-SHOP-JUMP-END -->\s*',
            "`r`n"
        )

        if ($html -match '(?is)</body>') {
            $html = [regex]::Replace($html,'(?is)</body>',"$jump`r`n</body>",1)
        } else {
            $html += "`r`n$jump`r`n"
        }

        if ($html -ne $original) {
            Backup-File $f.FullName
            Set-Content -LiteralPath $f.FullName -Value $html -Encoding UTF8
            $modified++
        }
    }
    catch {
        $failed++
    }
}

$ok = 0
$bad = 0
foreach ($f in $pages) {
    $t = Get-Content -LiteralPath $f.FullName -Raw -Encoding UTF8
    if ($t -match 'ML-SHOP-JUMP-START' -and
        $t -match 'class="ml-shop-jump-fixed"' -and
        $t -match 'z-index:99999') {
        $ok++
    } else {
        $bad++
    }
}

Write-Host ""
Write-Host "=== 우측 하단 업체 바로가기 보정 완료 ==="
Write-Host "대상 페이지:" $pages.Count
Write-Host "수정 페이지:" $modified
Write-Host "수정 실패:" $failed
Write-Host "바로가기 확인:" "$ok / $($pages.Count)"
Write-Host "확인 필요:" $bad
Write-Host ""
Write-Host "백업 폴더:" $backupRoot
Write-Host "상세페이지 수정: 0"
