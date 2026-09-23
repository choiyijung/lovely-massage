$ErrorActionPreference = "Stop"

$root = "E:\mong-lounge-site"
$massageRoot = Join-Path $root "massage"
$stamp = Get-Date -Format "yyyyMMdd-HHmmss"
$backupRoot = "E:\mong-lounge-remove-floating-shop-button-backup-$stamp"

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
            $t = Get-Content -LiteralPath $_.FullName -Raw -Encoding UTF8
            $t -match 'ML-BODYTOP-SHOP-BUTTON|ML-VISIBLE-SHOP-BUTTON|ML-SHOP-JUMP|ml-bodytop-shop-button|ml-visible-shop-button|ml-shop-jump-fixed'
        } catch { $false }
    }
)

$modified = 0
$failed = 0

foreach ($f in $pages) {
    try {
        $html = Get-Content -LiteralPath $f.FullName -Raw -Encoding UTF8
        $original = $html

        $html = [regex]::Replace($html,'(?is)\s*<!-- ML-BODYTOP-SHOP-BUTTON-START -->.*?<!-- ML-BODYTOP-SHOP-BUTTON-END -->\s*',"`r`n")
        $html = [regex]::Replace($html,'(?is)\s*<!-- ML-VISIBLE-SHOP-BUTTON-START -->.*?<!-- ML-VISIBLE-SHOP-BUTTON-END -->\s*',"`r`n")
        $html = [regex]::Replace($html,'(?is)\s*<!-- ML-SHOP-JUMP-START -->.*?<!-- ML-SHOP-JUMP-END -->\s*',"`r`n")

        # 혹시 마커 없이 남은 버튼이 있으면 제거
        $html = [regex]::Replace($html,'(?is)<button\b[^>]*id=["'']ml-bodytop-shop-button["''][^>]*>.*?</button>',"")
        $html = [regex]::Replace($html,'(?is)<a\b[^>]*id=["'']ml-visible-shop-button["''][^>]*>.*?</a>',"")
        $html = [regex]::Replace($html,'(?is)<a\b[^>]*class=["''][^"'']*ml-shop-jump-fixed[^"'']*["''][^>]*>.*?</a>',"")

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

$remaining = @(
    Get-ChildItem -LiteralPath $massageRoot -Recurse -File -Filter "index.html" |
    Where-Object {
        try {
            $t = Get-Content -LiteralPath $_.FullName -Raw -Encoding UTF8
            $t -match 'ML-BODYTOP-SHOP-BUTTON|ML-VISIBLE-SHOP-BUTTON|ML-SHOP-JUMP|ml-bodytop-shop-button|ml-visible-shop-button|ml-shop-jump-fixed'
        } catch { $false }
    }
).Count

Write-Host ""
Write-Host "=== 우측 하단 업체 바로가기 제거 완료 ==="
Write-Host "대상 페이지:" $pages.Count
Write-Host "수정 페이지:" $modified
Write-Host "수정 실패:" $failed
Write-Host "남은 우측 하단 버튼:" $remaining
Write-Host ""
Write-Host "등록 업체 카드 영역: 유지"
Write-Host "상세페이지 수정: 0"
Write-Host "백업 폴더:" $backupRoot
