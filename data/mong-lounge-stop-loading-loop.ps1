$ErrorActionPreference = "Stop"

$root = "E:\mong-lounge-site"
$stamp = Get-Date -Format "yyyyMMdd-HHmmss"
$backup = "E:\mong-lounge-loading-fix-backup-$stamp"
$badJs = Join-Path $root "assets\js\main-recommend-existing-cards.js"

New-Item -ItemType Directory -Force -Path $backup | Out-Null

$targets = @(
    (Join-Path $root "index.html"),
    (Join-Path $root "vip.html"),
    (Join-Path $root "premium.html")
)

$modified = 0

foreach($p in $targets){
    if(!(Test-Path -LiteralPath $p)){ continue }

    Copy-Item -LiteralPath $p -Destination (Join-Path $backup ([IO.Path]::GetFileName($p))) -Force

    $html = Get-Content -LiteralPath $p -Raw -Encoding UTF8
    $before = $html

    $html = [regex]::Replace(
        $html,
        '(?is)\s*<!--\s*ML-EXISTING-RECOMMEND-PATCH-START\s*-->.*?<!--\s*ML-EXISTING-RECOMMEND-PATCH-END\s*-->\s*',
        "`r`n"
    )

    $html = [regex]::Replace(
        $html,
        '(?is)\s*<script\b[^>]*src=["'']assets/js/main-recommend-existing-cards\.js[^"'']*["''][^>]*>\s*</script>\s*',
        "`r`n"
    )

    if($html -ne $before){
        Set-Content -LiteralPath $p -Value $html -Encoding UTF8
        $modified++
    }
}

if(Test-Path -LiteralPath $badJs){
    Copy-Item -LiteralPath $badJs -Destination (Join-Path $backup "main-recommend-existing-cards.js") -Force
    Rename-Item -LiteralPath $badJs -NewName "main-recommend-existing-cards.js.disabled" -Force
}

Write-Host ""
Write-Host "=== 무한 로딩 긴급 복구 완료 ==="
Write-Host "HTML 수정:" $modified
Write-Host "문제 JS 비활성화:" (Test-Path (Join-Path $root "assets\js\main-recommend-existing-cards.js.disabled"))
Write-Host "엑셀 수정: 0"
Write-Host "상세페이지 수정: 0"
Write-Host "백업 폴더:" $backup
Write-Host ""
Write-Host "이제 브라우저 탭을 닫았다가 5512 주소를 다시 여세요."
