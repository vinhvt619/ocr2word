param(
    [ValidateRange(1024, 65535)][int]$Port = 3000,
    [switch]$BuildFirst
)

# ASCII-only on purpose: Windows PowerShell 5.1 can parse this file without BOM.
$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$frontendRoot = Join-Path $projectRoot "frontend"
$vercelVersion = "59.4.0"

function Invoke-Vercel {
    param([Parameter(Mandatory = $true)][string[]]$Arguments)
    $globalCli = Get-Command vercel -ErrorAction SilentlyContinue
    if ($globalCli) {
        & $globalCli.Source @Arguments
    } else {
        & npx.cmd --yes "vercel@$vercelVersion" @Arguments
    }
}

function Read-DotEnvValue {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Name
    )
    $prefix = "$Name="
    $line = Get-Content -LiteralPath $Path | Where-Object { $_.StartsWith($prefix) } | Select-Object -First 1
    if (-not $line) { return $null }
    $value = $line.Substring($prefix.Length)
    if ($value.Length -ge 2 -and $value.StartsWith('"') -and $value.EndsWith('"')) {
        $value = $value.Substring(1, $value.Length - 2)
    }
    return $value
}

function Stop-StaleProjectWorkerContainers {
    $projectConfig = Join-Path $projectRoot ".vercel\project.json"
    if (-not (Test-Path -LiteralPath $projectConfig)) { return }
    $projectId = (Get-Content -LiteralPath $projectConfig -Raw | ConvertFrom-Json).projectId
    if ([string]::IsNullOrWhiteSpace($projectId)) { return }
    $containers = @(& docker ps --filter "ancestor=vercel-dev/worker:dev" --format "{{.Names}}")
    foreach ($container in $containers) {
        if ([string]::IsNullOrWhiteSpace($container)) { continue }
        $containerEnv = @(& docker inspect --format "{{range .Config.Env}}{{println .}}{{end}}" $container)
        if ($containerEnv -contains "VERCEL_PROJECT_ID=$projectId") {
            & docker stop $container *> $null
        }
    }
}

function Stop-StaleProjectWebProcesses {
    if ($env:OS -ne "Windows_NT") { return }
    $nextRuntimeMarker = Join-Path $frontendRoot "node_modules\next\dist"
    $staleProcesses = Get-CimInstance Win32_Process -ErrorAction SilentlyContinue | Where-Object {
        $_.CommandLine -and $_.CommandLine.Contains($nextRuntimeMarker, [StringComparison]::OrdinalIgnoreCase)
    }
    foreach ($staleProcess in $staleProcesses) {
        Stop-Process -Id $staleProcess.ProcessId -Force -ErrorAction SilentlyContinue
    }
}

function Clear-NextDevCache {
    $frontendAbsolute = [IO.Path]::GetFullPath($frontendRoot).TrimEnd('\', '/') + [IO.Path]::DirectorySeparatorChar
    $nextDevCache = [IO.Path]::GetFullPath((Join-Path $frontendRoot ".next\dev"))
    if (-not $nextDevCache.StartsWith($frontendAbsolute, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Refusing to clear a Next.js cache outside the frontend directory."
    }
    if (-not (Test-Path -LiteralPath $nextDevCache)) { return }
    # Next.js creates package junctions in this cache. Remove the links first
    # so recursive cleanup never traverses into frontend/node_modules.
    $links = Get-ChildItem -LiteralPath $nextDevCache -Recurse -Force -Attributes ReparsePoint -ErrorAction SilentlyContinue |
        Sort-Object { $_.FullName.Length } -Descending
    foreach ($link in $links) {
        if ($link.FullName.StartsWith($nextDevCache, [StringComparison]::OrdinalIgnoreCase)) {
            Remove-Item -LiteralPath $link.FullName -Force -ErrorAction SilentlyContinue
        }
    }
    Remove-Item -LiteralPath $nextDevCache -Recurse -Force
}

$nodeVersion = (& node --version).Trim()
if ($nodeVersion -notmatch '^v24\.') {
    throw "Local parity requires Node.js 24.x, but found $nodeVersion. Use 'nvm install 24' then 'nvm use 24'."
}
if (-not (Get-Command docker -ErrorAction SilentlyContinue)) {
    throw "Docker Desktop is missing."
}
& docker info *> $null
if ($LASTEXITCODE -ne 0) { throw "Docker Desktop is not running." }

$projectLink = Join-Path $projectRoot ".vercel\project.json"
$repoLink = Join-Path $projectRoot ".vercel\repo.json"
if (-not (Test-Path -LiteralPath $projectLink) -and -not (Test-Path -LiteralPath $repoLink)) {
    throw "Project is not linked. Run .\setup-vercel-local.ps1 first."
}
$frontendEnvFile = Join-Path $frontendRoot ".env.local"
if (-not (Test-Path -LiteralPath $frontendEnvFile)) {
    Write-Host "frontend/.env.local is missing; it will be created from Vercel Development." -ForegroundColor Yellow
}

& npm.cmd run sync-static --prefix $frontendRoot
if ($LASTEXITCODE -ne 0) { throw "Static asset sync failed." }

$forwardedEnvNames = @("OCR2WORD_INTERNAL_SECRET", "GEMINI_API_KEY", "OCR2WORD_AI_MODEL")
$previousForwardedValues = @{}
Push-Location $projectRoot
try {
    Write-Host "Refreshing Development environment and OIDC token..." -ForegroundColor Cyan
    Invoke-Vercel -Arguments @("pull", "--yes", "--environment=development")
    if ($LASTEXITCODE -ne 0) { throw "Could not refresh Vercel Development project settings." }
    Invoke-Vercel -Arguments @("env", "pull", "frontend/.env.local", "--yes", "--environment=development")
    if ($LASTEXITCODE -ne 0) { throw "Could not refresh Vercel Development environment." }
    foreach ($name in $forwardedEnvNames) {
        $previousForwardedValues[$name] = [Environment]::GetEnvironmentVariable($name, "Process")
        $value = Read-DotEnvValue -Path $frontendEnvFile -Name $name
        if ([string]::IsNullOrWhiteSpace($value)) {
            throw "frontend/.env.local is missing required worker variable $name."
        }
        [Environment]::SetEnvironmentVariable($name, $value, "Process")
    }

    if ($BuildFirst) {
        Write-Host "Running TypeScript parity check..." -ForegroundColor Cyan
        & npm.cmd run typecheck --prefix $frontendRoot
        if ($LASTEXITCODE -ne 0) { throw "TypeScript check failed." }
        if ($env:OS -eq "Windows_NT") {
            # Vercel's Windows packager creates function symlinks and fails with
            # EPERM unless Windows Developer Mode or elevated privileges are enabled.
            # Validate the same two deployable artifacts without requiring a
            # machine-wide security setting; vercel dev below still verifies the
            # complete Services graph and bindings end to end.
            Write-Host "Running the Next.js production build for Windows parity..." -ForegroundColor Cyan
            & npm.cmd run build --prefix $frontendRoot
            if ($LASTEXITCODE -ne 0) { throw "Next.js production build failed." }
            Write-Host "Building the Vercel worker container for Windows parity..." -ForegroundColor Cyan
            & docker build --file (Join-Path $projectRoot "Dockerfile.vercel") --tag "ocr2word-worker:local-parity" $projectRoot
            if ($LASTEXITCODE -ne 0) { throw "Vercel worker container build failed." }
        } else {
            Write-Host "Running the Vercel production build locally..." -ForegroundColor Cyan
            Invoke-Vercel -Arguments @("build")
            if ($LASTEXITCODE -ne 0) { throw "Vercel build failed." }
        }
    }

    $env:PORT = "$Port"
    $env:OCR2WORD_LOCAL_PARITY = "1"
    $env:VERCEL_QUEUE_DEBUG = "1"
    Clear-NextDevCache
    Write-Host "Starting the same Services graph used by Vercel..." -ForegroundColor Green
    Write-Host "Open http://127.0.0.1:$Port" -ForegroundColor Green
    Write-Host "This mode uses the real Development Blob, Queue, Redis and Google Gemini API." -ForegroundColor Yellow
    $devAttempt = 1
    while ($devAttempt -le 2) {
        $devStartedAt = Get-Date
        Invoke-Vercel -Arguments @("dev", "--listen", "127.0.0.1:$Port")
        $devExitCode = $LASTEXITCODE
        $devElapsedSeconds = ((Get-Date) - $devStartedAt).TotalSeconds
        if ($devExitCode -eq 0) { break }
        if ($devAttempt -eq 1 -and $devElapsedSeconds -lt 60) {
            Write-Host "Vercel CLI failed during its first Windows watcher startup; cleaning the project worker and retrying once..." -ForegroundColor Yellow
            Stop-StaleProjectWorkerContainers
            Stop-StaleProjectWebProcesses
            Clear-NextDevCache
            Start-Sleep -Seconds 2
            $devAttempt += 1
            continue
        }
        throw "vercel dev stopped with an error."
    }
} finally {
    Remove-Item Env:OCR2WORD_LOCAL_PARITY -ErrorAction SilentlyContinue
    Remove-Item Env:VERCEL_QUEUE_DEBUG -ErrorAction SilentlyContinue
    foreach ($name in $forwardedEnvNames) {
        [Environment]::SetEnvironmentVariable($name, $previousForwardedValues[$name], "Process")
    }
    Pop-Location
}
