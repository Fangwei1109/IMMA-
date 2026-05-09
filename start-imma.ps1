$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $root

function Import-EnvFile {
  param([string]$Path)

  if (-not (Test-Path $Path)) {
    return
  }

  Get-Content $Path | ForEach-Object {
    $line = $_.Trim()
    if (-not $line -or $line.StartsWith("#")) {
      return
    }

    $separator = $line.IndexOf("=")
    if ($separator -lt 1) {
      return
    }

    $name = $line.Substring(0, $separator).Trim()
    $value = $line.Substring($separator + 1).Trim().Trim('"').Trim("'")
    if ($name) {
      [Environment]::SetEnvironmentVariable($name, $value, "Process")
    }
  }
}

function Test-ImmaHealth {
  param([string]$Url)

  try {
    $response = Invoke-WebRequest -UseBasicParsing -Uri $Url -TimeoutSec 2
    return $response.StatusCode -ge 200 -and $response.StatusCode -lt 300
  } catch {
    return $false
  }
}

if (-not (Test-Path ".env") -and (Test-Path ".env.example")) {
  Copy-Item ".env.example" ".env"
  Write-Host "Created .env from .env.example. Add API keys there when needed."
}

Import-EnvFile (Join-Path $root ".env")

$port = if ($env:PORT) { $env:PORT } else { "3100" }
$baseUrl = "http://localhost:$port"
$healthUrl = "$baseUrl/api/health"

if (-not (Get-Command node -ErrorAction SilentlyContinue)) {
  throw "Node.js is not installed or not available in PATH. Install Node.js LTS first."
}

if (-not (Test-Path "node_modules")) {
  Write-Host "Installing dependencies..."
  npm install
}

if (Test-ImmaHealth $healthUrl) {
  Write-Host "IMMA is already running at $baseUrl"
  Start-Process $baseUrl
  exit 0
}

Write-Host "Starting IMMA..."
$process = Start-Process -FilePath "node" -ArgumentList "src/index.js" -WorkingDirectory $root -WindowStyle Hidden -PassThru
Set-Content -Path ".imma-server.pid" -Value $process.Id -Encoding ASCII

$ready = $false
for ($i = 0; $i -lt 30; $i++) {
  Start-Sleep -Seconds 1
  if (Test-ImmaHealth $healthUrl) {
    $ready = $true
    break
  }
}

if (-not $ready) {
  throw "IMMA started but did not become ready at $baseUrl. Check the server process or run npm start manually."
}

Write-Host "IMMA is running at $baseUrl"
Start-Process $baseUrl
