$ErrorActionPreference = "Stop"

$root = "E:\mong-lounge-site"
$massageRoot = Join-Path $root "massage"
$dataDir = Join-Path $root "data"
$stamp = Get-Date -Format "yyyyMMdd-HHmmss"
$backupRoot = "E:\mong-lounge-missing-region-shop-backup-$stamp"
$reportPath = Join-Path $dataDir "missing-region-shop-fill-report.csv"

New-Item -ItemType Directory -Force -Path $backupRoot | Out-Null

function Read-Html([string]$path) {
    Get-Content -LiteralPath $path -Raw -Encoding UTF8
}

function Get-Relative([string]$path) {
    $full = [System.IO.Path]::GetFullPath($path)
    $base = [System.IO.Path]::GetFullPath($root)
    if ($full.StartsWith($base,[System.StringComparison]::OrdinalIgnoreCase)) {
        return $full.Substring($base.Length).TrimStart("\")
    }
    return [System.IO.Path]::GetFileName($full)
}

function Backup-File([string]$path) {
    $rel = Get-Relative $path
    $dst = Join-Path $backupRoot $rel
    New-Item -ItemType Directory -Force -Path (Split-Path -Parent $dst) | Out-Null
    Copy-Item -LiteralPath $path -Destination $dst -Force
}

function Get-ShopBlock([string]$html) {
    $m = [regex]::Match(
        $html,
        '(?is)<!--\s*ML-TOP-SHOP-START\s*-->.*?<!--\s*ML-TOP-SHOP-END\s*-->'
    )
    if ($m.Success) { return $m.Value }
    return $null
}

function Get-ShopAnchors([string]$html) {
    $list = New-Object System.Collections.Generic.List[string]
    $block = Get-ShopBlock $html
    if (!$block) { return $list }

    foreach ($m in [regex]::Matches(
        $block,
        '(?is)<a\b[^>]*href=["''][^"'']*/shops/[^"'']*["''][^>]*>.*?</a>'
    )) {
        $list.Add($m.Value)
    }
    return $list
}

function Get-VendorKey([string]$anchor) {
    $m = [regex]::Match($anchor,'(?is)href=["'']([^"'']*/shops/[^"'']+)["'']')
    if (!$m.Success) { return $anchor }

    $href = $m.Groups[1].Value.Split('?')[0].Split('#')[0].TrimEnd('/')
    $parts = @($href -split '/' | Where-Object { $_ -ne '' })
    if ($parts.Count -eq 0) { return $href }

    if ($parts[-1] -match '^index\.html$' -and $parts.Count -ge 2) {
        return $parts[-2]
    }
    return $parts[-1]
}

function Ensure-InnerButton([string]$anchor) {
    # 기존 상세보기/업체 바로가기 버튼을 정리하고 작은 노란 버튼 1개만 유지
    $a = [regex]::Replace(
        $anchor,
        '(?is)<span\b[^>]*class=["''][^"'']*ml-card-shop-go[^"'']*["''][^>]*>.*?</span>',
        ''
    )
    $a = [regex]::Replace($a,'(?is)<[^>]+>\s*상세보기\s*</[^>]+>','')
    $a = $a -replace '상세보기',''

    $button = '<span class="ml-card-shop-go" style="display:flex!important;align-items:center!important;justify-content:center!important;width:92px!important;height:28px!important;margin:12px 12px 0 auto!important;padding:0!important;border:1px solid #ffc928!important;border-radius:8px!important;background:#ffc928!important;color:#083f35!important;font-size:11px!important;font-weight:900!important;line-height:1!important;text-decoration:none!important;box-shadow:0 3px 8px rgba(0,0,0,.16)!important;">업체 바로가기</span>'

    return [regex]::Replace(
        $a,
        '(?is)</a>\s*$',
        "`r`n$button</a>",
        1
    )
}

function Get-RegionName([string]$html) {
    # title/h1/breadcrumb 등에서 "수원시 출장마사지" 같은 지역명을 우선 추출
    $plain = [System.Net.WebUtility]::HtmlDecode(
        [regex]::Replace($html,'(?is)<script\b.*?</script>|<style\b.*?</style>|<[^>]+>',' ')
    )
    $plain = ($plain -replace '\s+',' ').Trim()

    $patterns = @(
        '([가-힣]{1,12}(?:특별시|광역시|특별자치시|특별자치도|도|시|군|구))\s*출장\s*마사지',
        '([가-힣]{1,12})\s*출장\s*마사지'
    )
    foreach ($p in $patterns) {
        $m = [regex]::Match($plain,$p)
        if ($m.Success) { return $m.Groups[1].Value.Trim() }
    }
    return "지역"
}

function Build-Block([string]$templateBlock, [string[]]$anchors, [string]$regionName) {
    $ms = [regex]::Matches(
        $templateBlock,
        '(?is)<a\b[^>]*href=["''][^"'']*/shops/[^"'']*["''][^>]*>.*?</a>'
    )
    if ($ms.Count -eq 0) { return $null }

    $clean = New-Object System.Collections.Generic.List[string]
    foreach ($a in $anchors) {
        $clean.Add((Ensure-InnerButton $a))
    }
    $joined = $clean -join "`r`n"

    $first = $ms[0]
    $last = $ms[$ms.Count - 1]
    $end = $last.Index + $last.Length

    $newBlock = $templateBlock.Substring(0,$first.Index) +
                $joined +
                $templateBlock.Substring($end)

    # 제목을 현재 상위 지역명으로
    $newBlock = [regex]::Replace(
        $newBlock,
        '(?is)(<h2\b[^>]*>).*?등록\s*업체(</h2>)',
        ('$1' + $regionName + ' 등록 업체$2'),
        1
    )

    # 우측 업체 수 표시가 있으면 실제 숫자로 갱신
    $newBlock = [regex]::Replace(
        $newBlock,
        '\d+\s*개\s*업체',
        ($anchors.Count.ToString() + '개 업체'),
        1
    )

    return $newBlock
}

$allPages = @(
    Get-ChildItem -LiteralPath $massageRoot -Recurse -File -Filter "index.html"
)

$htmlCache = @{}
$shopPages = New-Object System.Collections.Generic.List[object]

foreach ($f in $allPages) {
    try {
        $html = Read-Html $f.FullName
        $htmlCache[$f.FullName] = $html

        if ((Get-ShopAnchors $html).Count -gt 0) {
            $shopPages.Add($f)
        }
    } catch {}
}

$candidates = New-Object System.Collections.Generic.List[object]

foreach ($f in $allPages) {
    if (!$htmlCache.ContainsKey($f.FullName)) { continue }

    $html = $htmlCache[$f.FullName]

    # 이미 등록업체 영역이 있으면 건드리지 않음
    if ((Get-ShopAnchors $html).Count -gt 0) { continue }

    # 지역 서비스 페이지에만 적용
    if ($html -notmatch '출장\s*마사지|출장마사지') { continue }

    $dir = Split-Path -Parent $f.FullName

    # 1차: 현재 디렉터리 아래의 업체 페이지를 찾음
    $sources = @(
        $shopPages | Where-Object {
            $_.FullName -ne $f.FullName -and
            $_.FullName.StartsWith($dir + "\", [System.StringComparison]::OrdinalIgnoreCase)
        }
    )

    $scope = $dir

    # SEO 긴 이름 폴더의 index는 자식이 없으므로,
    # 2차: 바로 위 지역 폴더(예: ...\gyeonggi\suwon\) 전체에서 찾음
    if ($sources.Count -eq 0) {
        $parent = Split-Path -Parent $dir

        # massage 루트까지 올라가지는 않음
        if ($parent -and
            $parent.StartsWith($massageRoot,[System.StringComparison]::OrdinalIgnoreCase) -and
            $parent -ne $massageRoot) {

            $sources = @(
                $shopPages | Where-Object {
                    $_.FullName -ne $f.FullName -and
                    $_.FullName.StartsWith($parent + "\", [System.StringComparison]::OrdinalIgnoreCase)
                }
            )
            $scope = $parent
        }
    }

    if ($sources.Count -eq 0) { continue }

    # 가장 가까운 기존 등록업체 영역을 디자인 템플릿으로 사용
    $sources = @(
        $sources | Sort-Object {
            $_.FullName.Substring($scope.Length).Split('\').Count
        }, FullName
    )

    $templateBlock = Get-ShopBlock $htmlCache[$sources[0].FullName]
    if (!$templateBlock) { continue }

    # 같은 업체가 여러 동/구에 있어도 상위 페이지에는 업체별 1개
    $vendors = [ordered]@{}
    foreach ($s in $sources) {
        foreach ($a in (Get-ShopAnchors $htmlCache[$s.FullName])) {
            $key = Get-VendorKey $a
            if (!$vendors.Contains($key)) {
                $vendors[$key] = $a
            }
        }
    }

    if ($vendors.Count -eq 0) { continue }

    $regionName = Get-RegionName $html
    $newBlock = Build-Block $templateBlock @($vendors.Values) $regionName
    if (!$newBlock) { continue }

    $candidates.Add([pscustomobject]@{
        File        = $f.FullName
        Region      = $regionName
        Scope       = $scope
        VendorCount = $vendors.Count
        Block       = $newBlock
    })
}

$modified = 0
$failed = 0
$report = New-Object System.Collections.Generic.List[object]

foreach ($c in $candidates) {
    try {
        $html = Read-Html $c.File
        $original = $html

        # 서울처럼 상단에 보이도록 main 시작 직후 삽입
        $main = [regex]::Match($html,'(?is)<main\b[^>]*>')
        if ($main.Success) {
            $at = $main.Index + $main.Length
            $html = $html.Substring(0,$at) +
                    "`r`n" + $c.Block + "`r`n" +
                    $html.Substring($at)
        }
        else {
            $body = [regex]::Match($html,'(?is)<body\b[^>]*>')
            if (!$body.Success) { throw "main/body 태그 없음" }

            $at = $body.Index + $body.Length
            $html = $html.Substring(0,$at) +
                    "`r`n" + $c.Block + "`r`n" +
                    $html.Substring($at)
        }

        Backup-File $c.File
        Set-Content -LiteralPath $c.File -Value $html -Encoding UTF8
        $modified++

        $report.Add([pscustomobject]@{
            지역 = $c.Region
            업체수 = $c.VendorCount
            파일 = $c.File
            결과 = "수정"
        })
    }
    catch {
        $failed++
        $report.Add([pscustomobject]@{
            지역 = $c.Region
            업체수 = $c.VendorCount
            파일 = $c.File
            결과 = "실패: $($_.Exception.Message)"
        })
    }
}

$report | Export-Csv -LiteralPath $reportPath -NoTypeInformation -Encoding UTF8

# 최종 검증
$ok = 0
foreach ($c in $candidates) {
    try {
        $t = Read-Html $c.File
        if ((Get-ShopAnchors $t).Count -gt 0 -and $t -match 'ml-card-shop-go') {
            $ok++
        }
    } catch {}
}

Write-Host ""
Write-Host "=== 누락 상위지역 등록업체 영역 생성 완료 ==="
Write-Host "누락 대상:" $candidates.Count
Write-Host "수정 페이지:" $modified
Write-Host "수정 실패:" $failed
Write-Host "등록업체 + 노란버튼 확인:" "$ok / $($candidates.Count)"
Write-Host ""
Write-Host "수원시 포함 SEO 긴 이름 상위 페이지: 자동 탐색"
Write-Host "기존 구/동/읍/면 등록업체 영역: 수정 안 함"
Write-Host "상세페이지 수정: 0"
Write-Host "백업 폴더:" $backupRoot
Write-Host "보고서:" $reportPath
