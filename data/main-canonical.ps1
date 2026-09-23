$ErrorActionPreference = "Stop"

$file = "E:\mong-lounge-site\index.html"
$url = "https://1004테라피.shop/"
$stamp = Get-Date -Format "yyyyMMdd-HHmmss"
$backup = "E:\mong-lounge-main-canonical-backup-$stamp"

if(!(Test-Path -LiteralPath $file)){
    throw "메인 index.html을 찾지 못했습니다: $file"
}

New-Item -ItemType Directory -Force -Path $backup | Out-Null
Copy-Item -LiteralPath $file -Destination (Join-Path $backup "index.html") -Force

$html = Get-Content -LiteralPath $file -Raw -Encoding UTF8

# 기존 canonical 제거
$html = [regex]::Replace(
    $html,
    '(?is)\s*<link\b(?=[^>]*\brel\s*=\s*["'']canonical["''])[^>]*>\s*',
    "`r`n"
)

$tag = '<link rel="canonical" href="' + $url + '">'

if([regex]::IsMatch($html,'(?is)</title>')){
    $html = [regex]::Replace(
        $html,
        '(?is)</title>',
        '</title>' + "`r`n" + $tag,
        1
    )
}
elseif([regex]::IsMatch($html,'(?is)</head>')){
    $html = [regex]::Replace(
        $html,
        '(?is)</head>',
        $tag + "`r`n</head>",
        1
    )
}
else{
    throw "HEAD 영역을 찾지 못했습니다."
}

Set-Content -LiteralPath $file -Value $html -Encoding UTF8

# 검증
$check = Get-Content -LiteralPath $file -Raw -Encoding UTF8
$matches = [regex]::Matches(
    $check,
    '(?is)<link\b(?=[^>]*\brel\s*=\s*["'']canonical["''])(?=[^>]*\bhref\s*=\s*["'']([^"'']+)["''])[^>]*>'
)

$href = if($matches.Count -gt 0){ $matches[0].Groups[1].Value } else { "" }

Write-Host ""
Write-Host "=== 메인 canonical 적용 완료 ==="
Write-Host "canonical 태그 수:" $matches.Count
Write-Host "canonical URL:" $href
Write-Host "정상 여부:" ($matches.Count -eq 1 -and $href -eq $url)
Write-Host "백업 폴더:" $backup
