param(
    [string]$Root = "C:\Users\vmichalovska\Documents\VK_medijos_stebesena",
    [int]$Port = 8901
)
Add-Type -AssemblyName System.Net.HttpListener -ErrorAction SilentlyContinue
$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://localhost:$Port/")
$listener.Start()
Write-Host "Serving $Root at http://localhost:$Port/"

$mime = @{
    ".html" = "text/html; charset=utf-8"; ".json" = "application/json; charset=utf-8"
    ".js" = "application/javascript"; ".css" = "text/css"; ".md" = "text/plain; charset=utf-8"
    ".xlsx" = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
}

while ($listener.IsListening) {
    $ctx = $listener.GetContext()
    $path = $ctx.Request.Url.LocalPath
    if ($path -eq "/") { $path = "/index.html" }
    $fsPath = Join-Path $Root ($path.TrimStart("/") -replace "/", "\")
    if (Test-Path $fsPath -PathType Leaf) {
        $ext = [System.IO.Path]::GetExtension($fsPath)
        $ctx.Response.ContentType = if ($mime.ContainsKey($ext)) { $mime[$ext] } else { "application/octet-stream" }
        $bytes = [System.IO.File]::ReadAllBytes($fsPath)
        $ctx.Response.OutputStream.Write($bytes, 0, $bytes.Length)
    } else {
        $ctx.Response.StatusCode = 404
        $msg = [System.Text.Encoding]::UTF8.GetBytes("Not found: $path")
        $ctx.Response.OutputStream.Write($msg, 0, $msg.Length)
    }
    $ctx.Response.OutputStream.Close()
}
