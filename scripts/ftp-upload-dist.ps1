# Uploads dist/ to FTP remote dir. Reads .ftp-credentials.local (never logs password).
# Usage: powershell -File scripts/ftp-upload-dist.ps1

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
$remoteDir = ($map['FTP_REMOTE_DIR'] -replace '\\', '/').TrimEnd('/')
if (-not $remoteDir) { $remoteDir = '/app' }
$useSsl = ($map['FTP_SECURE'] -eq 'true')

$localRoot = Join-Path $root 'dist'
if (-not (Test-Path (Join-Path $localRoot 'index.html'))) {
    throw "dist/index.html fehlt - zuerst builden (VITE_BASE=/)."
}

$files = Get-ChildItem -Path $localRoot -Recurse -File
Write-Host "Upload $($files.Count) Dateien -> ftp://${hostName}${remoteDir}/ (ssl=$useSsl)"

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
            $resp = $req.GetResponse()
            $resp.Close()
        } catch {
            # existiert schon / kein Recht – weiter
        }
    }
}

Ensure-RemoteDir $remoteDir

# Alle benötigten Unterordner einmalig anlegen
$dirs = $files | ForEach-Object {
    $rel = $_.FullName.Substring($localRoot.Length).TrimStart('\', '/')
    $parent = Split-Path $rel -Parent
    if ($parent) { ($parent -replace '\\', '/') }
} | Where-Object { $_ } | Sort-Object -Unique

foreach ($d in $dirs) {
    Ensure-RemoteDir "$remoteDir/$d"
}

$i = 0
$ok = 0
$fail = 0
foreach ($f in $files) {
    $i++
    $rel = $f.FullName.Substring($localRoot.Length).TrimStart('\', '/')
    $relUnix = $rel -replace '\\', '/'
    $uri = "ftp://${hostName}${remoteDir}/$relUnix"
    try {
        $req = New-FtpRequest $uri ([System.Net.WebRequestMethods+Ftp]::UploadFile)
        $bytes = [System.IO.File]::ReadAllBytes($f.FullName)
        $req.ContentLength = $bytes.Length
        $stream = $req.GetRequestStream()
        $stream.Write($bytes, 0, $bytes.Length)
        $stream.Close()
        $resp = $req.GetResponse()
        $resp.Close()
        $ok++
        if (($i % 25) -eq 0 -or $i -eq $files.Count) {
            Write-Host "  $i / $($files.Count) …"
        }
    } catch {
        $fail++
        Write-Warning "FAIL $relUnix : $($_.Exception.Message)"
    }
}

Write-Host "Fertig: ok=$ok fail=$fail"
if ($fail -gt 0) { exit 1 }
