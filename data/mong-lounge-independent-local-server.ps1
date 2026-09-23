param(
    [string]$Root = "E:\mong-lounge-site",
    [int]$Port = 5512
)

$ErrorActionPreference = "Stop"

if (!(Test-Path -LiteralPath $Root)) {
    throw "사이트 폴더를 찾을 수 없습니다: $Root"
}

$Root = [System.IO.Path]::GetFullPath($Root)

$mime = @{
    ".html" = "text/html; charset=utf-8"
    ".htm"  = "text/html; charset=utf-8"
    ".css"  = "text/css; charset=utf-8"
    ".js"   = "application/javascript; charset=utf-8"
    ".json" = "application/json; charset=utf-8"
    ".svg"  = "image/svg+xml"
    ".png"  = "image/png"
    ".jpg"  = "image/jpeg"
    ".jpeg" = "image/jpeg"
    ".gif"  = "image/gif"
    ".webp" = "image/webp"
    ".ico"  = "image/x-icon"
    ".xml"  = "application/xml; charset=utf-8"
    ".txt"  = "text/plain; charset=utf-8"
    ".csv"  = "text/csv; charset=utf-8"
    ".woff" = "font/woff"
    ".woff2"= "font/woff2"
}

$listener = [System.Net.Sockets.TcpListener]::new([System.Net.IPAddress]::Loopback, $Port)
$listener.Start()

Write-Host ""
Write-Host "=== 몽라운지 독립 로컬 서버 실행 ==="
Write-Host "사이트 폴더: $Root"
Write-Host "주소: http://127.0.0.1:$Port/"
Write-Host "종료: 이 창에서 Ctrl+C"
Write-Host ""

function Send-Response {
    param(
        [System.Net.Sockets.NetworkStream]$Stream,
        [int]$Status,
        [string]$StatusText,
        [byte[]]$Body,
        [string]$ContentType = "text/plain; charset=utf-8"
    )

    $header = "HTTP/1.1 $Status $StatusText`r`n" +
              "Content-Type: $ContentType`r`n" +
              "Content-Length: $($Body.Length)`r`n" +
              "Cache-Control: no-store, no-cache, must-revalidate, max-age=0`r`n" +
              "Pragma: no-cache`r`n" +
              "Expires: 0`r`n" +
              "Connection: close`r`n`r`n"

    $headerBytes = [System.Text.Encoding]::ASCII.GetBytes($header)
    $Stream.Write($headerBytes,0,$headerBytes.Length)
    if ($Body.Length -gt 0) {
        $Stream.Write($Body,0,$Body.Length)
    }
    $Stream.Flush()
}

try {
    while ($true) {
        $client = $listener.AcceptTcpClient()

        try {
            $stream = $client.GetStream()
            $reader = [System.IO.StreamReader]::new(
                $stream,
                [System.Text.Encoding]::ASCII,
                $false,
                8192,
                $true
            )

            $requestLine = $reader.ReadLine()
            if ([string]::IsNullOrWhiteSpace($requestLine)) {
                $client.Close()
                continue
            }

            while ($true) {
                $line = $reader.ReadLine()
                if ($null -eq $line -or $line -eq "") { break }
            }

            $parts = $requestLine -split " "
            if ($parts.Count -lt 2 -or $parts[0] -ne "GET") {
                $body = [System.Text.Encoding]::UTF8.GetBytes("405 Method Not Allowed")
                Send-Response $stream 405 "Method Not Allowed" $body
                $client.Close()
                continue
            }

            $urlPath = $parts[1].Split("?")[0]
            $urlPath = [System.Uri]::UnescapeDataString($urlPath)

            if ([string]::IsNullOrWhiteSpace($urlPath) -or $urlPath -eq "/") {
                $urlPath = "/index.html"
            }
            elseif ($urlPath.EndsWith("/")) {
                $urlPath += "index.html"
            }

            $relative = $urlPath.TrimStart("/").Replace("/","\")
            $filePath = [System.IO.Path]::GetFullPath((Join-Path $Root $relative))

            if (!$filePath.StartsWith($Root,[System.StringComparison]::OrdinalIgnoreCase)) {
                $body = [System.Text.Encoding]::UTF8.GetBytes("403 Forbidden")
                Send-Response $stream 403 "Forbidden" $body
                $client.Close()
                continue
            }

            if (Test-Path -LiteralPath $filePath -PathType Container) {
                $filePath = Join-Path $filePath "index.html"
            }

            if (!(Test-Path -LiteralPath $filePath -PathType Leaf)) {
                $body = [System.Text.Encoding]::UTF8.GetBytes("404 Not Found")
                Send-Response $stream 404 "Not Found" $body
                $client.Close()
                continue
            }

            $bytes = [System.IO.File]::ReadAllBytes($filePath)
            $ext = [System.IO.Path]::GetExtension($filePath).ToLowerInvariant()
            $contentType = if ($mime.ContainsKey($ext)) { $mime[$ext] } else { "application/octet-stream" }

            Send-Response $stream 200 "OK" $bytes $contentType
        }
        catch {
            try {
                if ($stream) {
                    $body = [System.Text.Encoding]::UTF8.GetBytes("500 Internal Server Error")
                    Send-Response $stream 500 "Internal Server Error" $body
                }
            } catch {}
        }
        finally {
            if ($reader) { $reader.Dispose() }
            if ($client) { $client.Close() }
        }
    }
}
finally {
    $listener.Stop()
}
