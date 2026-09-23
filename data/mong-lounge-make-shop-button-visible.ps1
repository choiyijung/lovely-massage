$ErrorActionPreference = "Stop"

$root = "E:\mong-lounge-site"
$massageRoot = Join-Path $root "massage"
$dataDir = Join-Path $root "data"
$stamp = Get-Date -Format "yyyyMMdd-HHmmss"
$backupRoot = "E:\mong-lounge-visible-shop-button-backup-$stamp"

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
            (Get-Content -LiteralPath $_.FullName -Raw -Encoding UTF8) -match 'id=["'']registered-shop["'']'
        } catch { $false }
    }
)

$buttonBlock = @'
<!-- ML-VISIBLE-SHOP-BUTTON-START -->
<style id="ml-visible-shop-button-style">
#ml-visible-shop-button{
  position:fixed!important;
  right:24px!important;
  bottom:72px!important;
  z-index:2147483647!important;
  display:flex!important;
  align-items:center!important;
  justify-content:center!important;
  width:148px!important;
  height:52px!important;
  margin:0!important;
  padding:0 14px!important;
  border:2px solid #083f35!important;
  border-radius:999px!important;
  background:#ffc928!important;
  color:#083f35!important;
  box-shadow:0 10px 30px rgba(0,0,0,.30)!important;
  font-family:Arial,"Noto Sans KR","Apple SD Gothic Neo",sans-serif!important;
  font-size:15px!important;
  font-weight:900!important;
  line-height:1!important;
  letter-spacing:-.2px!important;
  text-decoration:none!important;
  visibility:visible!important;
  opacity:1!important;
  pointer-events:auto!important;
}
#ml-visible-shop-button:hover{
  transform:translateY(-2px)!important;
}
@media(max-width:700px){
  #ml-visible-shop-button{
    right:14px!important;
    bottom:70px!important;
    width:132px!important;
    height:48px!important;
    font-size:14px!important;
  }
}
</style>
<a id="ml-visible-shop-button" href="#registered-shop" aria-label="등록 업체 바로가기">업체 바로가기</a>
<!-- ML-VISIBLE-SHOP-BUTTON-END -->
'@

$modified = 0
$failed = 0
$i = 0

foreach ($f in $pages) {
    $i++
    try {
        $html = Get-Content -LiteralPath $f.FullName -Raw -Encoding UTF8
        $original = $html

        # 이전에 넣었던 업체 바로가기 계열은 모두 제거
        $html = [regex]::Replace(
            $html,
            '(?is)\s*<!-- ML-SHOP-JUMP-START -->.*?<!-- ML-SHOP-JUMP-END -->\s*',
            "`r`n"
        )
        $html = [regex]::Replace(
            $html,
            '(?is)\s*<!-- ML-VISIBLE-SHOP-BUTTON-START -->.*?<!-- ML-VISIBLE-SHOP-BUTTON-END -->\s*',
            "`r`n"
        )

        if ($html -match '(?is)</body>') {
            $html = [regex]::Replace(
                $html,
                '(?is)</body>',
                "$buttonBlock`r`n</body>",
                1
            )
        } else {
            $html += "`r`n$buttonBlock`r`n"
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

    if (($i % 200) -eq 0 -or $i -eq $pages.Count) {
        Write-Host "진행: $i / $($pages.Count)"
    }
}

$ok = 0
foreach ($f in $pages) {
    $t = Get-Content -LiteralPath $f.FullName -Raw -Encoding UTF8
    if ($t -match 'id="ml-visible-shop-button"' -and
        $t -match 'bottom:72px!important' -and
        $t -match 'background:#ffc928!important') {
        $ok++
    }
}

# 앞으로 전체 계층 적용 스크립트를 다시 돌려도 같은 버튼을 유지하도록
$allLevelsScript = Join-Path $dataDir "mong-lounge-add-shops-to-all-levels.ps1"
$patched = $false
if (Test-Path -LiteralPath $allLevelsScript) {
    $src = Get-Content -LiteralPath $allLevelsScript -Raw -Encoding UTF8
    if ($src -notmatch 'ML-VISIBLE-SHOP-BUTTON-FUTURE') {
        Backup-File $allLevelsScript
        Add-Content -LiteralPath $allLevelsScript -Encoding UTF8 -Value @'

# ML-VISIBLE-SHOP-BUTTON-FUTURE
$mlVisibleButtonScript = Join-Path $dataDir "mong-lounge-make-shop-button-visible.ps1"
if (Test-Path -LiteralPath $mlVisibleButtonScript) {
    & $mlVisibleButtonScript -FromGenerator
}
'@
    }
    $patched = $true
}

Write-Host ""
Write-Host "=== 업체 바로가기 버튼 표시 완료 ==="
Write-Host "대상 페이지:" $pages.Count
Write-Host "수정 페이지:" $modified
Write-Host "수정 실패:" $failed
Write-Host "버튼 확인:" "$ok / $($pages.Count)"
Write-Host "향후 전체적용 스크립트 연결:" $patched
Write-Host ""
Write-Host "버튼 위치: 우측 하단"
Write-Host "버튼 색상: 노란색"
Write-Host "버튼 문구: 업체 바로가기"
Write-Host "백업 폴더:" $backupRoot
Write-Host "상세페이지 수정: 0"
