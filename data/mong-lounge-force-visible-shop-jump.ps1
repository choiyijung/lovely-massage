$ErrorActionPreference = "Stop"

$root = "E:\mong-lounge-site"
$massageRoot = Join-Path $root "massage"
$stamp = Get-Date -Format "yyyyMMdd-HHmmss"
$backupRoot = "E:\mong-lounge-shop-jump-runtime-backup-$stamp"

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
            (Get-Content -LiteralPath $_.FullName -Raw -Encoding UTF8) -match '<!-- ML-SHOP-JUMP-START -->'
        } catch { $false }
    }
)

$block = @'
<!-- ML-SHOP-JUMP-START -->
<a id="ml-shop-jump-fixed" href="#registered-shop">업체 바로가기</a>
<script>
(function () {
  function mountShopJump() {
    var btn = document.getElementById('ml-shop-jump-fixed');
    if (!btn) return;

    /* 마크업이 비정상적으로 중첩돼 있어도 body 직속으로 이동 */
    if (btn.parentNode !== document.body) {
      document.body.appendChild(btn);
    }

    btn.style.setProperty('position','fixed','important');
    btn.style.setProperty('right','18px','important');
    btn.style.setProperty('bottom','20px','important');
    btn.style.setProperty('left','auto','important');
    btn.style.setProperty('top','auto','important');
    btn.style.setProperty('z-index','2147483647','important');
    btn.style.setProperty('display','flex','important');
    btn.style.setProperty('align-items','center','important');
    btn.style.setProperty('justify-content','center','important');
    btn.style.setProperty('min-width','118px','important');
    btn.style.setProperty('height','46px','important');
    btn.style.setProperty('padding','0 18px','important');
    btn.style.setProperty('margin','0','important');
    btn.style.setProperty('border-radius','999px','important');
    btn.style.setProperty('background','#083f35','important');
    btn.style.setProperty('color','#ffc928','important');
    btn.style.setProperty('border','1px solid rgba(255,201,40,.70)','important');
    btn.style.setProperty('box-shadow','0 8px 28px rgba(0,0,0,.28)','important');
    btn.style.setProperty('font-size','14px','important');
    btn.style.setProperty('font-weight','900','important');
    btn.style.setProperty('line-height','1','important');
    btn.style.setProperty('text-decoration','none','important');
    btn.style.setProperty('visibility','visible','important');
    btn.style.setProperty('opacity','1','important');
    btn.style.setProperty('transform','none','important');
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', mountShopJump);
  } else {
    mountShopJump();
  }
})();
</script>
<!-- ML-SHOP-JUMP-END -->
'@

$modified = 0
$failed = 0

foreach ($f in $pages) {
    try {
        $html = Get-Content -LiteralPath $f.FullName -Raw -Encoding UTF8
        $newHtml = [regex]::Replace(
            $html,
            '(?is)<!-- ML-SHOP-JUMP-START -->.*?<!-- ML-SHOP-JUMP-END -->',
            [System.Text.RegularExpressions.MatchEvaluator]{ param($m) $block },
            1
        )

        if ($newHtml -ne $html) {
            Backup-File $f.FullName
            Set-Content -LiteralPath $f.FullName -Value $newHtml -Encoding UTF8
            $modified++
        }
    }
    catch {
        $failed++
    }
}

$ok = 0
foreach ($f in $pages) {
    $t = Get-Content -LiteralPath $f.FullName -Raw -Encoding UTF8
    if ($t -match 'id="ml-shop-jump-fixed"' -and
        $t -match 'document\.body\.appendChild\(btn\)' -and
        $t -match '2147483647') {
        $ok++
    }
}

Write-Host ""
Write-Host "=== 업체 바로가기 화면표시 강제 보정 완료 ==="
Write-Host "대상 페이지:" $pages.Count
Write-Host "수정 페이지:" $modified
Write-Host "수정 실패:" $failed
Write-Host "런타임 고정 확인:" "$ok / $($pages.Count)"
Write-Host ""
Write-Host "백업 폴더:" $backupRoot
Write-Host "상세페이지 수정: 0"
