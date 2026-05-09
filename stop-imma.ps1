$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $root

$pidFile = Join-Path $root ".imma-server.pid"
if (Test-Path $pidFile) {
  $serverPid = Get-Content $pidFile -ErrorAction SilentlyContinue | Select-Object -First 1
  if ($serverPid) {
    Stop-Process -Id ([int]$serverPid) -Force -ErrorAction SilentlyContinue
    Remove-Item $pidFile -Force -ErrorAction SilentlyContinue
    Write-Host "Stopped IMMA server process $serverPid."
    exit 0
  }
}

$connections = Get-NetTCPConnection -LocalPort 3100 -State Listen -ErrorAction SilentlyContinue
if ($connections) {
  $connections | ForEach-Object {
    Stop-Process -Id $_.OwningProcess -Force -ErrorAction SilentlyContinue
    Write-Host "Stopped process $($_.OwningProcess) listening on port 3100."
  }
} else {
  Write-Host "No IMMA server process found."
}
