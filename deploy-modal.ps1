param(
    [switch]$BudgetConfirmed,
    [switch]$SkipTests
)

# ASCII-only on purpose: Windows PowerShell 5.1 can parse this file without BOM.
$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$pythonExe = Join-Path $projectRoot ".venv\Scripts\python.exe"
$modalExe = Join-Path $projectRoot ".venv\Scripts\modal.exe"
$urlFile = Join-Path $projectRoot ".modal-url"

if (-not $BudgetConfirmed) {
    throw "Safety stop: set a hard Modal Workspace Budget (recommended USD 20) first, then rerun with -BudgetConfirmed."
}
if (-not (Test-Path -LiteralPath $pythonExe) -or -not (Test-Path -LiteralPath $modalExe)) {
    throw "Modal is not installed. Run .\setup-modal.ps1 first."
}

$env:PYTHONUTF8 = "1"
& $modalExe token info *> $null
if ($LASTEXITCODE -ne 0) { throw "Modal CLI is not authenticated. Run .\setup-modal.ps1 first." }

$secretJson = (& $modalExe secret list --json | Out-String)
if ($LASTEXITCODE -ne 0) { throw "Could not list Modal Secrets." }
if ($secretJson -notmatch 'ocr2word-secrets') {
    throw "Modal Secret 'ocr2word-secrets' is missing. Run .\setup-modal.ps1 first."
}

Push-Location $projectRoot
try {
    if (-not $SkipTests) {
        Write-Host "Running deployment checks..." -ForegroundColor Cyan
        & $pythonExe -m py_compile modal_app.py app\services\serverless_jobs.py
        if ($LASTEXITCODE -ne 0) { throw "Python syntax checks failed." }
        & $pythonExe -m pytest tests\test_serverless_jobs.py -q --basetemp artifacts\pytest-modal-deploy
        if ($LASTEXITCODE -ne 0) { throw "Modal deployment tests failed." }
        if (Get-Command node -ErrorAction SilentlyContinue) {
            & node --check app\static\app.js
            if ($LASTEXITCODE -ne 0) { throw "Browser JavaScript syntax check failed." }
        }
    }

    if (Get-Command npm.cmd -ErrorAction SilentlyContinue) {
        & npm.cmd run sync-static --prefix frontend
        if ($LASTEXITCODE -ne 0) { throw "Could not synchronize the web interface." }
    }

    Write-Host "Deploying OCR2Word to Modal..." -ForegroundColor Cyan
    & $modalExe deploy modal_app.py --strategy rolling 2>&1 | Tee-Object -Variable deployOutput
    if ($LASTEXITCODE -ne 0) { throw "Modal deployment failed." }
} finally {
    Pop-Location
}

$allOutput = $deployOutput | Out-String
$match = [Regex]::Match($allOutput, 'https://[A-Za-z0-9.-]+\.modal\.run(?:/[^\s]*)?')
if ($match.Success) {
    $url = $match.Value.TrimEnd('.', ',', ')')
    [IO.File]::WriteAllText($urlFile, $url, (New-Object Text.UTF8Encoding($false)))
    Write-Host "OCR2Word URL: $url" -ForegroundColor Green
    Write-Host "Verify it with: .\verify-modal.ps1" -ForegroundColor Green
} else {
    Write-Warning "Deployment succeeded, but the URL could not be extracted. Open it with: .\.venv\Scripts\modal.exe dashboard"
}
