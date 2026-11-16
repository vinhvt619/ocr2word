param(
    [switch]$SkipNpmInstall,
    [switch]$SkipLogin
)

# ASCII-only on purpose: Windows PowerShell 5.1 can parse this file without BOM.
$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$frontendRoot = Join-Path $projectRoot "frontend"
$envFile = Join-Path $frontendRoot ".env.local"
$vercelVersion = "59.4.0"

function Assert-Node24 {
    $node = Get-Command node -ErrorAction SilentlyContinue
    if (-not $node) {
        throw "Node.js is missing. Install Node.js 24 LTS first."
    }
    $version = (& node --version).Trim()
    if ($version -notmatch '^v24\.') {
        throw "Local parity requires Node.js 24.x, but found $version. Use 'nvm install 24' then 'nvm use 24'."
    }
}

function Assert-Docker {
    if (-not (Get-Command docker -ErrorAction SilentlyContinue)) {
        throw "Docker Desktop is missing. Install and start Docker Desktop before local parity testing."
    }
    & docker info *> $null
    if ($LASTEXITCODE -ne 0) {
        throw "Docker is installed but its daemon is not running. Start Docker Desktop."
    }
}

function Invoke-Vercel {
    param([Parameter(Mandatory = $true)][string[]]$Arguments)
    $globalCli = Get-Command vercel -ErrorAction SilentlyContinue
    if ($globalCli) {
        & $globalCli.Source @Arguments
    } else {
        & npx.cmd --yes "vercel@$vercelVersion" @Arguments
    }
}

function Read-EnvKeys {
    param([Parameter(Mandatory = $true)][string]$Path)
    $keys = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::Ordinal)
    foreach ($line in Get-Content -LiteralPath $Path) {
        if ($line -match '^\s*([A-Za-z_][A-Za-z0-9_]*)=') {
            [void]$keys.Add($Matches[1])
        }
    }
    return ,$keys
}

Assert-Node24
Assert-Docker

if (-not $SkipNpmInstall) {
    Write-Host "Installing locked frontend dependencies..." -ForegroundColor Cyan
    & npm.cmd ci --prefix $frontendRoot
    if ($LASTEXITCODE -ne 0) { throw "npm ci failed." }
}

Write-Host "Checking Vercel CLI authentication..." -ForegroundColor Cyan
Invoke-Vercel -Arguments @("--version")
if ($LASTEXITCODE -ne 0) { throw "Vercel CLI could not start." }

Invoke-Vercel -Arguments @("whoami")
if ($LASTEXITCODE -ne 0) {
    if ($SkipLogin) {
        throw "Vercel CLI is not authenticated. Run this script again without -SkipLogin."
    }
    Invoke-Vercel -Arguments @("login")
    if ($LASTEXITCODE -ne 0) { throw "Vercel login failed." }
}

$projectLink = Join-Path $projectRoot ".vercel\project.json"
$repoLink = Join-Path $projectRoot ".vercel\repo.json"
if (-not (Test-Path -LiteralPath $projectLink) -and -not (Test-Path -LiteralPath $repoLink)) {
    Write-Host "Link this folder to the Vercel project that will host OCR2Word." -ForegroundColor Cyan
    Push-Location $projectRoot
    try {
        Invoke-Vercel -Arguments @("link")
        if ($LASTEXITCODE -ne 0) { throw "Vercel project linking failed." }
    } finally {
        Pop-Location
    }
}

Push-Location $projectRoot
try {
    Write-Host "Pulling Development project settings and fresh OIDC credentials..." -ForegroundColor Cyan
    Invoke-Vercel -Arguments @("pull", "--yes", "--environment=development")
    if ($LASTEXITCODE -ne 0) { throw "vercel pull failed." }
    Invoke-Vercel -Arguments @("env", "pull", $envFile, "--yes", "--environment=development")
    if ($LASTEXITCODE -ne 0) { throw "vercel env pull failed." }
} finally {
    Pop-Location
}

if (-not (Test-Path -LiteralPath $envFile)) {
    throw ".env.local was not created."
}

$keys = Read-EnvKeys -Path $envFile
$missing = New-Object 'System.Collections.Generic.List[string]'
foreach ($key in @(
    "OCR2WORD_INTERNAL_SECRET",
    "OCR2WORD_QUEUE_REGION",
    "VERCEL_OIDC_TOKEN"
)) {
    if (-not $keys.Contains($key)) { $missing.Add($key) }
}
$hasUpstashRedis = $keys.Contains("UPSTASH_REDIS_REST_URL") -and $keys.Contains("UPSTASH_REDIS_REST_TOKEN")
$hasVercelKvRedis = $keys.Contains("KV_REST_API_URL") -and $keys.Contains("KV_REST_API_TOKEN")
if (-not $hasUpstashRedis -and -not $hasVercelKvRedis) {
    $missing.Add("UPSTASH_REDIS_REST_URL/TOKEN or KV_REST_API_URL/TOKEN")
}
$hasBlobToken = $keys.Contains("BLOB_READ_WRITE_TOKEN")
$hasBlobOidc = $keys.Contains("BLOB_STORE_ID") -and $keys.Contains("VERCEL_OIDC_TOKEN")
if (-not $hasBlobToken -and -not $hasBlobOidc) {
    $missing.Add("BLOB_READ_WRITE_TOKEN or BLOB_STORE_ID")
}

if ($missing.Count -gt 0) {
    $names = $missing -join ", "
    throw "Development environment is incomplete: $names. Connect Private Blob and Upstash Redis."
}

if (Test-Path -LiteralPath $projectLink) {
    try {
        $project = Get-Content -LiteralPath $projectLink -Raw | ConvertFrom-Json
        $framework = $project.settings.framework
        if ($framework -and $framework -ne "services") {
            throw "Vercel Framework Preset is '$framework'. Change it to 'Services' in Project Settings before deploying."
        }
    } catch {
        if ($_.Exception.Message -like "Vercel Framework Preset*") { throw }
        Write-Warning "Could not confirm Framework Preset from .vercel/project.json. Confirm it is set to Services in the dashboard."
    }
}

Write-Host "Local parity setup is ready." -ForegroundColor Green
Write-Host "Next: .\run-vercel-local.ps1 -BuildFirst" -ForegroundColor Green
