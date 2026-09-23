$ErrorActionPreference = "Stop"

$root = "E:\mong-lounge-site"
$port = 5500
$prefix = "http://localhost:$port/"

if (!(Test-Path -LiteralPath $root)) {
    throw "사이트 폴더를 찾을 수 없습니다: $root"
}

$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add($prefix)

try {
    $listener.Start()
} catch {
    Write-Host "포트 $port 를 열지 못했습니다."
    Write-Host "이미 다른 서버가 실행 중일 수 있습니다."
    throw
}

Write-Host ""
Write-Host "=== 몽라운지 로컬 사이트 실행 ==="
Write-Host "사이트 주소: $prefix"
Write-Host "사이트 폴더: $root"
Write-Host "종료: Ctrl+C"
Write-Host ""

Start-Process $prefix

$mime = @{
    ".html"="text/html; charset=utf-8"
    ".htm" ="text/html; charset=utf-8"
    ".css" ="text/css; charset=utf-8"
    ".js"  ="application/javascript; charset=utf-8"
    ".json"="application/json; charset=utf-8"
    ".xml" ="application/xml; charset=utf-8"
    ".txt" ="text/plain; charset=utf-8"
    ".png" ="image/png"
    ".jpg" ="image/jpeg"
    ".jpeg"="image/jpeg"
    ".webp"="image/webp"
    ".gif" ="image/gif"
    ".svg" ="image/svg+xml"
    ".ico" ="image/x-icon"
    ".woff"="font/woff"
    ".woff2"="font/woff2"
}

while ($listener.IsListening) {
    try {
        $context = $listener.GetContext()
        $req = $context.Request
        $res = $context.Response

        $urlPath = [uri]::UnescapeDataString($req.Url.AbsolutePath).TrimStart("/")
        $urlPath = $urlPath -replace "/", "\"

        $candidate = Join-Path $root $urlPath

        if (Test-Path -LiteralPath $candidate -PathType Container) {
            $candidate = Join-Path $candidate "index.html"
        }

        if (!(Test-Path -LiteralPath $candidate -PathType Leaf)) {
            # 확장자 없는 경로라면 /index.html 시도
            $candidate2 = Join-Path (Join-Path $root $urlPath) "index.html"
            if (Test-Path -LiteralPath $candidate2 -PathType Leaf) {
                $candidate = $candidate2
            }
        }

        $full = [System.IO.Path]::GetFullPath($candidate)
        $rootFull = [System.IO.Path]::GetFullPath($root)

        if (!$full.StartsWith($rootFull, [System.StringComparison]::OrdinalIgnoreCase)) {
            $res.StatusCode = 403
            $bytes = [System.Text.Encoding]::UTF8.GetBytes("403 Forbidden")
        }
        elseif (!(Test-Path -LiteralPath $full -PathType Leaf)) {
            $res.StatusCode = 404
            $bytes = [System.Text.Encoding]::UTF8.GetBytes("404 Not Found")
        }
        else {
            $ext = [System.IO.Path]::GetExtension($full).ToLowerInvariant()
            if ($mime.ContainsKey($ext)) {
                $res.ContentType = $mime[$ext]
            } else {
                $res.ContentType = "application/octet-stream"
            }
            $bytes = [System.IO.File]::ReadAllBytes($full)
            $res.StatusCode = 200
        }

        $res.ContentLength64 = $bytes.Length
        $res.OutputStream.Write($bytes,0,$bytes.Length)
        $res.OutputStream.Close()
    }
    catch [System.Net.HttpListenerException] {
        break
    }
    catch {
        try {
            if ($context -and $context.Response) {
                $context.Response.StatusCode = 500
                $context.Response.Close()
            }
        } catch {}
    }
}

if ($listener.IsListening) {
    $listener.Stop()
}
$listener.Close()
