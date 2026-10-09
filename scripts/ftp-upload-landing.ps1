# Upload landing/ nach FTP (ms365.schule Root). Liest .ftp-credentials.local
# Usage: powershell -File scripts/ftp-upload-landing.ps1

$ErrorActionPreference = 'Stop'
$root = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
Set-Location $root

$credPath = Join-Path $root '.ftp-credentials.local'
if (-not (Test-Path $credPath)) { throw "Missing $credPath" }

$map = @{}
Get-Content $credPath | Where-Object { $_ -match '=' -and $_ -notmatch '^\s*#' } | ForEach-Object {
    $k, $v = $_.Split('=', 2)
    $map[$k.Trim()] = $v.Trim()
}

$hostName = $map['FTP_HOST']
$user = $map['FTP_USER']
$pass = $map['FTP_PASS']
$remoteDir = ($map['FTP_LANDING_REMOTE_DIR'] -replace '\\', '/').TrimEnd('/')
if (-not $remoteDir) { $remoteDir = '/' }
$useSsl = ($map['FTP_SECURE'] -eq 'true')

$localRoot = Join-Path $root 'landing'
if (-not (Test-Path (Join-Path $localRoot 'index.html'))) {
    throw "landing/index.html fehlt"
}

$files = Get-ChildItem -Path $localRoot -Recurse -File
Write-Host "Upload $($files.Count) Dateien (landing/) -> ftp://${hostName}${remoteDir}/ (ssl=$useSsl)"

function New-FtpRequest([string]$uri, [string]$method) {
    $req = [System.Net.FtpWebRequest]::Create($uri)
    $req.Credentials = New-Object System.Net.NetworkCredential($user, $pass)
    $req.Method = $method
    $req.UseBinary = $true
    $req.UsePassive = $true
    $req.KeepAlive = $false
    $req.EnableSsl = $useSsl
    return $req
}

function Ensure-RemoteDir([string]$dirPath) {
    $parts = $dirPath.Trim('/').Split('/') | Where-Object { $_ }
    $cur = ''
    foreach ($p in $parts) {
        $cur = "$cur/$p"
        $uri = "ftp://${hostName}${cur}"
        try {
            $req = New-FtpRequest $uri ([System.Net.WebRequestMethods+Ftp]::MakeDirectory)
            $req.GetResponse().Close()
        } catch { }
    }
}

foreach ($f in $files) {
    $rel = $f.FullName.Substring($localRoot.Length).TrimStart('\', '/').Replace('\', '/')
    $remotePath = "$remoteDir/$rel".Replace('//', '/')
    $remoteDirOnly = ($remotePath -replace '/[^/]+$', '')
    Ensure-RemoteDir $remoteDirOnly
    $uri = "ftp://${hostName}${remotePath}"
    $req = New-FtpRequest $uri ([System.Net.WebRequestMethods+Ftp]::UploadFile)
    $bytes = [System.IO.File]::ReadAllBytes($f.FullName)
    $req.ContentLength = $bytes.Length
    $stream = $req.GetRequestStream()
    $stream.Write($bytes, 0, $bytes.Length)
    $stream.Close()
    $req.GetResponse().Close()
    Write-Host "  $rel"
}

Write-Host "Fertig. Prüfen: https://ms365.schule/ (Footer: Homepage zuletzt aktualisiert)"
