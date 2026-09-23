$ErrorActionPreference = "Stop"

$root = "E:\mong-lounge-site"
$massageRoot = Join-Path $root "massage"
$stamp = Get-Date -Format "yyyyMMdd-HHmmss"
$backupRoot = "E:\mong-lounge-shop-button-bodytop-backup-$stamp"

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

$button = @'
<!-- ML-BODYTOP-SHOP-BUTTON-START -->
<button id="ml-bodytop-shop-button"
type="button"
onclick="var x=document.getElementById('registered-shop');if(x){x.scrollIntoView({behavior:'smooth',block:'start'});}"
style="position:fixed!important;right:22px!important;bottom:22px!important;left:auto!important;top:auto!important;z-index:2147483647!important;display:block!important;width:150px!important;height:54px!important;margin:0!important;padding:0!important;border:3px solid #083f35!important;border-radius:14px!important;background:#ffcc22!important;color:#083f35!important;box-shadow:0 8px 28px rgba(0,0,0,.35)!important;font-family:Arial,'Noto Sans KR','Apple SD Gothic Neo',sans-serif!important;font-size:15px!important;font-weight:900!important;line-height:48px!important;text-align:center!important;visibility:visible!important;opacity:1!important;pointer-events:auto!important;cursor:pointer!important;transform:none!important;">
업체 바로가기
</button>
<!-- ML-BODYTOP-SHOP-BUTTON-END -->
'@

$modified = 0
$failed = 0
$i = 0

foreach ($f in $pages) {
    $i++
    try {
        $html = Get-Content -LiteralPath $f.FullName -Raw -Encoding UTF8
        $original = $html

        # 지금까지 넣었던 버튼 전부 제거
        $patterns = @(
            '(?is)\s*<!-- ML-SHOP-JUMP-START -->.*?<!-- ML-SHOP-JUMP-END -->\s*',
            '(?is)\s*<!-- ML-VISIBLE-SHOP-BUTTON-START -->.*?<!-- ML-VISIBLE-SHOP-BUTTON-END -->\s*',
            '(?is)\s*<!-- ML-BODYTOP-SHOP-BUTTON-START -->.*?<!-- ML-BODYTOP-SHOP-BUTTON-END -->\s*'
        )
        foreach ($pat in $patterns) {
            $html = [regex]::Replace($html,$pat,"`r`n")
        }

        # 중요: 문서 맨 끝이 아니라 <body>가 열리자마자 바로 삽입
        $body = [regex]::Match($html,'(?is)<body\b[^>]*>')
        if (!$body.Success) {
            $failed++
            continue
        }

        $at = $body.Index + $body.Length
        $html = $html.Substring(0,$at) + "`r`n$button`r`n" + $html.Substring($at)

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
    $body = [regex]::Match($t,'(?is)<body\b[^>]*>')
    $btn = $t.IndexOf('ML-BODYTOP-SHOP-BUTTON-START')
    if ($body.Success -and $btn -gt $body.Index -and $btn -lt ($body.Index + 2500) -and
        $t -match 'id="ml-bodytop-shop-button"') {
        $ok++
    }
}

Write-Host ""
Write-Host "=== 업체 바로가기 최상단 삽입 완료 ==="
Write-Host "대상 페이지:" $pages.Count
Write-Host "수정 페이지:" $modified
Write-Host "수정 실패:" $failed
Write-Host "BODY 바로 아래 버튼 확인:" "$ok / $($pages.Count)"
Write-Host ""
Write-Host "버튼 위치: 화면 우측 하단"
Write-Host "버튼 색상: 노란색"
Write-Host "삽입 위치: BODY 시작 바로 아래"
Write-Host "백업 폴더:" $backupRoot
Write-Host "상세페이지 수정: 0"
