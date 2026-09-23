$ErrorActionPreference = "Stop"

$root = "E:\mong-lounge-site"
$stamp = Get-Date -Format "yyyyMMdd-HHmmss"
$backup = "E:\mong-lounge-northmyeon-meta-backup-$stamp"

$targets = @(
    @{
        Path = "massage\북면-출장마사지-창원-의창구-생활권-지역정보-이용안내\index.html"
        Description = "창원 의창구 북면의 주거·산업 생활권과 인접 지역 이동 흐름을 기준으로 위치 정보를 정리했습니다. 주변 읍·면·동과 연결되는 범위를 비교할 때 참고하세요."
    },
    @{
        Path = "massage\북면-출장마사지-가평군-생활권-지역정보-이용안내\index.html"
        Description = "가평군 북면의 산간 생활권과 주변 읍·면 이동 범위를 기준으로 지역 정보를 정리했습니다. 현재 위치에서 가까운 생활권과 이동 방향을 확인할 때 참고하세요."
    }
)

New-Item -ItemType Directory -Force -Path $backup | Out-Null

function Set-Meta([string]$html,[string]$attr,[string]$name,[string]$value){
    $encoded=[System.Net.WebUtility]::HtmlEncode($value)
    $replacement="<meta $attr=`"$name`" content=`"$encoded`">"

    $patterns=@(
        "(?is)<meta\b(?=[^>]*\b$attr\s*=\s*[""']$([regex]::Escape($name))[""'])(?=[^>]*\bcontent\s*=\s*[""'][^""']*[""'])[^>]*>",
        "(?is)<meta\b(?=[^>]*\bcontent\s*=\s*[""'][^""']*[""'])(?=[^>]*\b$attr\s*=\s*[""']$([regex]::Escape($name))[""'])[^>]*>"
    )

    foreach($p in $patterns){
        if([regex]::IsMatch($html,$p)){
            return [regex]::Replace($html,$p,$replacement,1)
        }
    }

    if($html -match '(?is)</head>'){
        return [regex]::Replace($html,'(?is)</head>',"    $replacement`r`n</head>",1)
    }

    return $html
}

$modified=0

foreach($t in $targets){
    $full=Join-Path $root $t.Path
    if(!(Test-Path -LiteralPath $full)){
        throw "파일 없음: $full"
    }

    $dest=Join-Path $backup $t.Path
    New-Item -ItemType Directory -Force -Path (Split-Path $dest -Parent) | Out-Null
    Copy-Item -LiteralPath $full -Destination $dest -Force

    $html=Get-Content -LiteralPath $full -Raw -Encoding UTF8
    $html=Set-Meta $html "name" "description" $t.Description
    $html=Set-Meta $html "property" "og:description" $t.Description
    Set-Content -LiteralPath $full -Value $html -Encoding UTF8
    $modified++
}

Write-Host ""
Write-Host "=== 북면 DESCRIPTION 중복 수정 완료 ==="
Write-Host "수정 페이지:" $modified
Write-Host "창원 의창구 북면 / 가평군 북면 설명 분리: 완료"
Write-Host "META = OG DESCRIPTION: 적용"
Write-Host "백업 폴더:" $backup
