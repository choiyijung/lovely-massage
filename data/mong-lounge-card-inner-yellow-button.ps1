$ErrorActionPreference = "Stop"

$root = "E:\mong-lounge-site"
$massageRoot = Join-Path $root "massage"
$stamp = Get-Date -Format "yyyyMMdd-HHmmss"
$backupRoot = "E:\mong-lounge-card-shop-button-backup-$stamp"

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
            $t -match '<!--\s*ML-TOP-SHOP-START\s*-->' -and $t -match '/shops/'
        } catch { $false }
    }
)

$modified = 0
$failed = 0
$buttonCount = 0
$floatingRemaining = 0
$i = 0

$buttonHtml = '<span class="ml-card-shop-go" style="display:flex!important;align-items:center!important;justify-content:center!important;width:92px!important;height:28px!important;margin:12px 12px 0 auto!important;padding:0!important;border:1px solid #ffc928!important;border-radius:8px!important;background:#ffc928!important;color:#083f35!important;font-size:11px!important;font-weight:900!important;line-height:1!important;text-decoration:none!important;box-shadow:0 3px 8px rgba(0,0,0,.16)!important;">업체 바로가기</span>'

foreach ($f in $pages) {
    $i++
    try {
        $html = Get-Content -LiteralPath $f.FullName -Raw -Encoding UTF8
        $original = $html

        # 1) 지금까지 만든 화면 우측 하단 고정 버튼 계열을 모두 제거
        $html = [regex]::Replace($html,'(?is)\s*<!--\s*ML-BODYTOP-SHOP-BUTTON-START\s*-->.*?<!--\s*ML-BODYTOP-SHOP-BUTTON-END\s*-->\s*',"`r`n")
        $html = [regex]::Replace($html,'(?is)\s*<!--\s*ML-VISIBLE-SHOP-BUTTON-START\s*-->.*?<!--\s*ML-VISIBLE-SHOP-BUTTON-END\s*-->\s*',"`r`n")
        $html = [regex]::Replace($html,'(?is)\s*<!--\s*ML-SHOP-JUMP-START\s*-->.*?<!--\s*ML-SHOP-JUMP-END\s*-->\s*',"`r`n")
        $html = [regex]::Replace($html,'(?is)<style\b[^>]*id=["'']ml-visible-shop-button-style["''][^>]*>.*?</style>',"")
        $html = [regex]::Replace($html,'(?is)<button\b[^>]*id=["'']ml-bodytop-shop-button["''][^>]*>.*?</button>',"")
        $html = [regex]::Replace($html,'(?is)<a\b[^>]*id=["'']ml-visible-shop-button["''][^>]*>.*?</a>',"")
        $html = [regex]::Replace($html,'(?is)<a\b[^>]*class=["''][^"'']*ml-shop-jump-fixed[^"'']*["''][^>]*>.*?</a>',"")

        # 2) 상단 등록업체 블록 안에서만 작업
        $topMatch = [regex]::Match($html,'(?is)<!--\s*ML-TOP-SHOP-START\s*-->(.*?)<!--\s*ML-TOP-SHOP-END\s*-->')
        if ($topMatch.Success) {
            $inner = $topMatch.Groups[1].Value

            # 이전 카드 내부 버튼이 있으면 제거 후 다시 정확히 삽입
            $inner = [regex]::Replace(
                $inner,
                '(?is)<span\b[^>]*class=["''][^"'']*ml-card-shop-go[^"'']*["''][^>]*>.*?</span>',
                ''
            )

            # /shops/ 상세페이지로 연결되는 카드 링크 안에 작은 노란 버튼 삽입
            $inner = [regex]::Replace(
                $inner,
                '(?is)(<a\b[^>]*href=["''][^"'']*/shops/[^"'']*["''][^>]*>)(.*?)(</a>)',
                [System.Text.RegularExpressions.MatchEvaluator]{
                    param($m)
                    $body = $m.Groups[2].Value
                    # 카드 안에 기존 "상세보기" 문구/버튼이 있으면 텍스트만 제거
                    $body = [regex]::Replace($body,'(?is)<[^>]*>\s*상세보기\s*</[^>]+>','')
                    $body = $body -replace '상세보기',''
                    return $m.Groups[1].Value + $body + "`r`n" + $buttonHtml + $m.Groups[3].Value
                }
            )

            $newTop = '<!-- ML-TOP-SHOP-START -->' + $inner + '<!-- ML-TOP-SHOP-END -->'
            $html = $html.Substring(0,$topMatch.Index) + $newTop + $html.Substring($topMatch.Index + $topMatch.Length)
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

# 검증
foreach ($f in $pages) {
    $t = Get-Content -LiteralPath $f.FullName -Raw -Encoding UTF8
    $buttonCount += ([regex]::Matches($t,'class=["''][^"'']*ml-card-shop-go[^"'']*["'']')).Count

    if ($t -match 'ml-bodytop-shop-button|ml-visible-shop-button|ml-shop-jump-fixed|ML-BODYTOP-SHOP-BUTTON|ML-VISIBLE-SHOP-BUTTON|ML-SHOP-JUMP-START') {
        $floatingRemaining++
    }
}

Write-Host ""
Write-Host "=== 카드 내부 업체 바로가기 적용 완료 ==="
Write-Host "대상 페이지:" $pages.Count
Write-Host "수정 페이지:" $modified
Write-Host "수정 실패:" $failed
Write-Host "카드 내부 노란 버튼 총 개수:" $buttonCount
Write-Host "우측 하단 고정 버튼 남은 페이지:" $floatingRemaining
Write-Host ""
Write-Host "버튼 문구: 업체 바로가기"
Write-Host "버튼 크기: 92 x 28"
Write-Host "버튼 위치: 업체 카드 내부 오른쪽 끝"
Write-Host "버튼 색상: 노란색"
Write-Host "상세페이지 수정: 0"
Write-Host "백업 폴더:" $backupRoot
