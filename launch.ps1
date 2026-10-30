param([switch]$NoBrowser)
$ErrorActionPreference = "SilentlyContinue"
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$port = 8770
$py = (Get-Command python -ErrorAction SilentlyContinue).Source
if (-not $py) { $py = (Get-Command py -ErrorAction SilentlyContinue).Source }
if (-not $py) {
  Write-Host "Python 3 was not found. Install it from https://www.python.org/downloads/."
  exit 1
}

$up = $false
try {
  $r = Invoke-WebRequest -Uri "http://127.0.0.1:$port/api/books" -UseBasicParsing -TimeoutSec 2
  if ($r.StatusCode -eq 200) { $up = $true }
} catch { $up = $false }

# server is up, but code may have changed after it started -> restart it
if ($up) {
  $stampPath = Join-Path $here "data\.server-stamp.json"
  $stale = $false
  if (Test-Path $stampPath) {
    try {
      $stamp = Get-Content $stampPath -Raw | ConvertFrom-Json
      foreach ($f in @("server.py", "index.html")) {
        $fp = Join-Path $here $f
        if (Test-Path $fp) {
          if ((Get-Item $fp).LastWriteTime -gt ([DateTimeOffset]::FromUnixTimeSeconds([int]$stamp.started_ts)).LocalDateTime) { $stale = $true }
        }
      }
      if ($stale -and $stamp.pid) {
        Stop-Process -Id ([int]$stamp.pid) -Force -ErrorAction SilentlyContinue
        Start-Sleep -Milliseconds 600
        $up = $false
      }
    } catch { }
  }
}

if (-not $up) {
  Start-Process -FilePath $py -ArgumentList @("`"$here\server.py`"") -WorkingDirectory $here -WindowStyle Hidden
  for ($i = 0; $i -lt 30; $i++) {
    Start-Sleep -Milliseconds 400
    try {
      $r = Invoke-WebRequest -Uri "http://127.0.0.1:$port/api/books" -UseBasicParsing -TimeoutSec 2
      if ($r.StatusCode -eq 200) { break }
    } catch { }
  }
}

# Open it the same way the reader does: a chrome-less app window (no tabs, no address bar).
# KEEP THIS BLOCK ASCII-ONLY: this file has no BOM, PowerShell 5.1 would garble non-ASCII text.
if (-not $NoBrowser) {
  $url  = "http://127.0.0.1:$port/"
  $edge = (Get-Command msedge.exe -ErrorAction SilentlyContinue).Source
  if (-not $edge) { $edge = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" }
  if (Test-Path $edge) {
    Start-Process -FilePath $edge -ArgumentList @("--app=$url", "--window-size=1440,940", "--new-window")
  } else {
    Start-Process $url
  }
}
