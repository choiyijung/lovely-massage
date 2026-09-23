$ErrorActionPreference = "Stop"

$root = "E:\mong-lounge-site"
$out  = Join-Path $root "data\other-region-link-audit.csv"

$targets = [ordered]@{
    "대전"       = "E:\mong-lounge-site\massage\other\daejeon"
    "청주"       = "E:\mong-lounge-site\massage\other\cheongju"
    "창원"       = "E:\mong-lounge-site\massage\other\changwon"
    "대구"       = "E:\mong-lounge-site\massage\other\daegu"
    "공주"       = "E:\mong-lounge-site\massage\other\gongju"
    "계룡"       = "E:\mong-lounge-site\massage\other\gyeryong"
    "전주"       = "E:\mong-lounge-site\massage\other\jeonju"
    "광주광역시" = "E:\mong-lounge-site\massage\other\gwangju"
    "익산"       = "E:\mong-lounge-site\massage\other\iksan"
    "김해"       = "E:\mong-lounge-site\massage\other\gimhae"
}

$rows = New-Object System.Collections.Generic.List[object]

foreach ($region in $targets.Keys) {
    $folder = $targets[$region]
    if (!(Test-Path -LiteralPath $folder)) {
        Write-Host "폴더 없음: $region -> $folder"
        continue
    }

    $files = @(Get-ChildItem -LiteralPath $folder -Recurse -File -Filter "index.html")
    foreach ($file in $files) {
        $html = Get-Content -LiteralPath $file.FullName -Raw -Encoding UTF8

        foreach ($m in [regex]::Matches($html, '(?is)<a\b[^>]*>(.*?)</a>')) {
            $text = $m.Groups[1].Value
            $text = [regex]::Replace($text, '(?is)<[^>]+>', ' ')
            $text = [System.Net.WebUtility]::HtmlDecode($text)
            $text = [regex]::Replace($text, '\s+', ' ').Trim()

            if ([string]::IsNullOrWhiteSpace($text)) { continue }

            # 동/읍/면/구 형태의 지역 링크 후보만 수집
            if ($text -match '(동|읍|면|구)$' -and $text.Length -le 30) {
                $rows.Add([pscustomobject]@{
                    지역       = $region
                    링크텍스트 = $text
                    원본페이지 = $file.FullName
                })
            }
        }
    }
}

$final = @(
    $rows |
    Sort-Object 지역, 링크텍스트, 원본페이지 -Unique
)

$final | Export-Csv -LiteralPath $out -NoTypeInformation -Encoding UTF8

Write-Host ""
Write-Host "=== 기타지역 링크 추출 점검 ==="
Write-Host "저장:" $out
Write-Host "전체 후보:" $final.Count
Write-Host ""

foreach ($region in $targets.Keys) {
    $r = @($final | Where-Object 지역 -eq $region)
    $names = @($r | Select-Object -ExpandProperty 링크텍스트 -Unique | Sort-Object)

    Write-Host "$region : 고유 지역명 $($names.Count)개 / 링크 $($r.Count)개"
    if ($names.Count -gt 0) {
        $sample = ($names | Select-Object -First 25) -join ", "
        Write-Host "  예시: $sample"
    }
}

Write-Host ""
Write-Host "사이트 HTML 수정: 0"
