$ErrorActionPreference = "Stop"

$root    = "E:\mong-lounge-site"
$dataDir = Join-Path $root "data"
$master  = Join-Path $dataDir "몽라운지_업체자동등록_마스터_v2.xlsx"
$stamp   = Get-Date -Format "yyyyMMdd-HHmmss"
$backup  = "E:\mong-lounge-static-recommend-backup-$stamp"
$temp    = Join-Path $env:TEMP "mong_lounge_static_recommend_$stamp.xlsx"

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

if (!(Test-Path -LiteralPath $master)) { throw "마스터 엑셀 없음: $master" }

# 엑셀 파일이 열려 있어도 읽기는 가능하도록 임시 복사
$inStream = [System.IO.File]::Open($master,[System.IO.FileMode]::Open,[System.IO.FileAccess]::Read,[System.IO.FileShare]::ReadWrite)
try {
    $outStream = [System.IO.File]::Create($temp)
    try { $inStream.CopyTo($outStream) } finally { $outStream.Dispose() }
}
finally { $inStream.Dispose() }

New-Item -ItemType Directory -Force -Path $backup | Out-Null

function Get-ZipText {
    param([System.IO.Compression.ZipArchive]$Zip,[string]$Name)
    $e=$Zip.GetEntry($Name)
    if(!$e){throw "XLSX 내부 파일 누락: $Name"}
    $s=$e.Open()
    try{
        $r=[IO.StreamReader]::new($s,[Text.Encoding]::UTF8,$true)
        try{return $r.ReadToEnd()}finally{$r.Dispose()}
    }finally{$s.Dispose()}
}

function Get-SharedStrings {
    param([System.IO.Compression.ZipArchive]$Zip)
    $list=New-Object System.Collections.Generic.List[string]
    if(!$Zip.GetEntry("xl/sharedStrings.xml")){return $list}
    [xml]$x=Get-ZipText $Zip "xl/sharedStrings.xml"
    foreach($si in $x.SelectNodes("//*[local-name()='si']")){
        $parts=@($si.SelectNodes(".//*[local-name()='t']")|ForEach-Object{$_.InnerText})
        $list.Add(($parts -join ""))
    }
    return $list
}

function Get-CellText {
    param([System.Xml.XmlElement]$Cell,$Shared)
    if(!$Cell){return ""}
    $t=$Cell.GetAttribute("t")
    if($t -eq "inlineStr"){
        return ((@($Cell.SelectNodes(".//*[local-name()='t']")|ForEach-Object{$_.InnerText})) -join "")
    }
    $v=$Cell.SelectSingleNode("./*[local-name()='v']")
    if(!$v){return ""}
    if($t -eq "s"){
        $i=0
        if([int]::TryParse($v.InnerText,[ref]$i) -and $i -ge 0 -and $i -lt $Shared.Count){
            return [string]$Shared[$i]
        }
    }
    return [string]$v.InnerText
}

function Get-SheetPath {
    param([xml]$Wb,[xml]$Rels,[string]$Name)
    $sheet=$Wb.SelectSingleNode("//*[local-name()='sheet' and @name='$Name']")
    if(!$sheet){throw "시트 없음: $Name"}
    $rid=$sheet.GetAttribute("id","http://schemas.openxmlformats.org/officeDocument/2006/relationships")
    $rel=$Rels.SelectSingleNode("//*[local-name()='Relationship' and @Id='$rid']")
    if(!$rel){throw "시트 관계 없음: $Name"}
    $target=$rel.GetAttribute("Target").Replace("\","/")
    if($target.StartsWith("/")){return $target.TrimStart("/")}
    return "xl/"+$target.TrimStart("/")
}

function Read-Recommend {
    param([System.IO.Compression.ZipArchive]$Zip,[xml]$Wb,[xml]$Rels,$Shared,[string]$SheetName)

    $path=Get-SheetPath $Wb $Rels $SheetName
    [xml]$sx=Get-ZipText $Zip $path
    $rows=@()

    foreach($row in $sx.SelectNodes("//*[local-name()='row' and number(@r)>=2]")){
        $r=[int]$row.GetAttribute("r")
        $v=@{}
        foreach($c in @("A","B","C","D","E","F","G","H")){
            $cell=$sx.SelectSingleNode("//*[local-name()='c' and @r='$c$r']")
            $v[$c]=(Get-CellText $cell $Shared).Trim()
        }
        if($v.A.ToUpperInvariant() -eq "N" -or !$v.B){continue}

        $sort=999
        [void][int]::TryParse($v.H,[ref]$sort)

        $rows += [pscustomobject]@{
            shop=$v.B
            region=$v.C
            price=$v.D
            image=$v.E
            link=$v.F
            desc=$v.G
            sort=$sort
        }
    }

    return @($rows|Sort-Object sort,shop|Select-Object -First 4)
}

$zip=[IO.Compression.ZipFile]::OpenRead($temp)
try{
    [xml]$wb=Get-ZipText $zip "xl/workbook.xml"
    [xml]$rels=Get-ZipText $zip "xl/_rels/workbook.xml.rels"
    $shared=Get-SharedStrings $zip
    $vip=@(Read-Recommend $zip $wb $rels $shared "메인VIP추천")
    $pre=@(Read-Recommend $zip $wb $rels $shared "메인프리미엄추천")
}
finally{
    $zip.Dispose()
    Remove-Item -LiteralPath $temp -Force -ErrorAction SilentlyContinue
}

if($vip.Count -lt 4 -or $pre.Count -lt 4){
    throw "추천 시트에 VIP 4개 / 프리미엄 4개가 필요합니다. 현재 VIP=$($vip.Count), 프리미엄=$($pre.Count)"
}

function Normalize-Link([string]$x){
    if(!$x){return "#"}
    if($x -match '^https?://'){return $x}
    return $x.TrimStart("/")
}
function Normalize-Image([string]$x){
    if(!$x){return "assets/images/shops/room-1.jpg"}
    if($x -match '^https?://'){return $x}
    return $x.TrimStart("/")
}

function Update-Anchor {
    param([string]$Anchor,$Item)

    $a=$Anchor

    # href
    $href=Normalize-Link $Item.link
    $a=[regex]::Replace(
        $a,
        '(?is)(<a\b[^>]*\bhref=["''])[^"'']*(["''])',
        ('$1'+$href+'$2'),
        1
    )

    # img src / alt
    $img=Normalize-Image $Item.image
    $a=[regex]::Replace(
        $a,
        '(?is)(<img\b[^>]*\bsrc=["''])[^"'']*(["''])',
        ('$1'+$img+'$2'),
        1
    )
    if($a -match '(?is)<img\b[^>]*\balt=["'']'){
        $a=[regex]::Replace(
            $a,
            '(?is)(<img\b[^>]*\balt=["''])[^"'']*(["''])',
            ('$1'+$Item.shop+'$2'),
            1
        )
    }

    # 업체명
    $a=[regex]::Replace(
        $a,
        '(?is)(<h3\b[^>]*>).*?(</h3>)',
        ('$1'+$Item.shop+'$2'),
        1
    )

    # 지역 설명
    $region=$Item.region
    if(!$region){$region="추천 지역"}
    $a=[regex]::Replace(
        $a,
        '(?is)(<p\b[^>]*>).*?(</p>)',
        ('$1'+$region+' · 추천 업체$2'),
        1
    )

    # 가격
    $price=$Item.price
    if(!$price){$price="가격 문의"}
    $a=[regex]::Replace(
        $a,
        '(?is)(<(?:b|strong)\b[^>]*>)[^<]*원[^<]*(</(?:b|strong)>)',
        ('$1'+$price+'$2'),
        1
    )

    return $a
}

$legacyHref='(?:shop-gangnam|shop-songpa|shop-suwon|shop-bupyeong)\.html'
$anchorPattern='(?is)<a\b(?=[^>]*\bhref=["'']'+$legacyHref+'["''])[^>]*>.*?</a>'

function Update-StaticPage {
    param(
        [string]$Path,
        [string]$Mode
    )

    if(!(Test-Path -LiteralPath $Path)){return [pscustomobject]@{Path=$Path;Found=0;Changed=0}}

    Copy-Item -LiteralPath $Path -Destination (Join-Path $backup ([IO.Path]::GetFileName($Path))) -Force

    $html=Get-Content -LiteralPath $Path -Raw -Encoding UTF8
    $matches=@([regex]::Matches($html,$anchorPattern))
    if($matches.Count -eq 0){return [pscustomobject]@{Path=$Path;Found=0;Changed=0}}

    # 뒤에서부터 교체해야 문자열 위치가 안 틀어짐
    $repls=New-Object System.Collections.Generic.List[object]

    for($i=0;$i -lt $matches.Count;$i++){
        $m=$matches[$i]

        if($Mode -eq "VIP"){
            $item=$vip[$i % 4]
        }
        elseif($Mode -eq "PREMIUM"){
            $item=$pre[$i % 4]
        }
        else {
            # 메인: 첫 4개 VIP, 다음 4개 프리미엄
            if($matches.Count -ge 8){
                if($i -lt 4){$item=$vip[$i]}else{$item=$pre[($i-4)%4]}
            }
            else {
                # 4개뿐이면 해당 카드 앞의 마지막 제목으로 판단
                $before=$html.Substring(0,$m.Index)
                $vipPos=$before.LastIndexOf("VIP 추천")
                $prePos=$before.LastIndexOf("프리미엄 추천")
                if($prePos -gt $vipPos){$item=$pre[$i%4]}else{$item=$vip[$i%4]}
            }
        }

        $newAnchor=Update-Anchor $m.Value $item
        $repls.Add([pscustomobject]@{Index=$m.Index;Length=$m.Length;Text=$newAnchor})
    }

    foreach($r in ($repls|Sort-Object Index -Descending)){
        $html=$html.Substring(0,$r.Index)+$r.Text+$html.Substring($r.Index+$r.Length)
    }

    Set-Content -LiteralPath $Path -Value $html -Encoding UTF8
    return [pscustomobject]@{Path=$Path;Found=$matches.Count;Changed=$matches.Count}
}

$results=@()
$results += Update-StaticPage (Join-Path $root "vip.html") "VIP"
$results += Update-StaticPage (Join-Path $root "premium.html") "PREMIUM"
$results += Update-StaticPage (Join-Path $root "index.html") "MAIN"

# 추천 관련 동적 JS는 메인에서 모두 제거: 정적 카드만 사용
$index=Join-Path $root "index.html"
$html=Get-Content -LiteralPath $index -Raw -Encoding UTF8

$patterns=@(
    '(?is)\s*<!--\s*ML-MAIN-RECOMMEND-AUTO-START\s*-->.*?<!--\s*ML-MAIN-RECOMMEND-AUTO-END\s*-->\s*',
    '(?is)\s*<!--\s*ML-MAIN-RECOMMEND-CLEANUP-START\s*-->.*?<!--\s*ML-MAIN-RECOMMEND-CLEANUP-END\s*-->\s*',
    '(?is)\s*<!--\s*ML-MAIN-RECOMMEND-FINAL-START\s*-->.*?<!--\s*ML-MAIN-RECOMMEND-FINAL-END\s*-->\s*',
    '(?is)\s*<!--\s*ML-EXISTING-RECOMMEND-PATCH-START\s*-->.*?<!--\s*ML-EXISTING-RECOMMEND-PATCH-END\s*-->\s*',
    '(?is)\s*<script\b[^>]*src=["'']assets/js/main-recommend-(?:auto|cleanup|final|existing-cards)\.js[^"'']*["''][^>]*>\s*</script>\s*'
)
foreach($p in $patterns){$html=[regex]::Replace($html,$p,"`r`n")}
Set-Content -LiteralPath $index -Value $html -Encoding UTF8

# 최종 검사
$old='강남 골드케어|송파 밸런스룸|광교 리셋테라피|부평 온기케어'
$checkFiles=@(
    (Join-Path $root "index.html"),
    (Join-Path $root "vip.html"),
    (Join-Path $root "premium.html")
)

Write-Host ""
Write-Host "=== 추천업체 정적 원본 직접 교체 완료 ==="
foreach($r in $results){
    Write-Host ([IO.Path]::GetFileName($r.Path)) "| 기존 카드 발견:" $r.Found "| 교체:" $r.Changed
}
Write-Host ""
Write-Host "VIP:"
$vip|ForEach-Object{Write-Host " -" $_.shop "|" $_.region "|" $_.price}
Write-Host "프리미엄:"
$pre|ForEach-Object{Write-Host " -" $_.shop "|" $_.region "|" $_.price}
Write-Host ""
foreach($f in $checkFiles){
    $t=Get-Content -LiteralPath $f -Raw -Encoding UTF8
    Write-Host ([IO.Path]::GetFileName($f)) "예전 임시업체 남음:" ($t -match $old)
}
$idx=Get-Content -LiteralPath (Join-Path $root "index.html") -Raw -Encoding UTF8
Write-Host "메인 추천용 동적 JS 연결 남음:" ($idx -match 'main-recommend-(?:auto|cleanup|final|existing-cards)\.js')
Write-Host "백업 폴더:" $backup
Write-Host ""
Write-Host "이제 5512 주소에서 Ctrl+F5 하세요."
