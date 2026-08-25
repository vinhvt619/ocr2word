param(
    [switch]$LocalOnly
)

# ASCII-only on purpose: Windows PowerShell 5.1 can parse this file safely.
$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$localEnvFile = Join-Path $projectRoot ".env"
$vercelVersion = "59.4.0"
$model = "gemini-3.7-flash"

function Invoke-Vercel {
    param([Parameter(Mandatory = $true)][string[]]$Arguments)
    $globalCli = Get-Command vercel -ErrorAction SilentlyContinue
    if ($globalCli) {
        & $globalCli.Source @Arguments
    } else {
        & npx.cmd --yes "vercel@$vercelVersion" @Arguments
    }
}

function Invoke-VercelWithInput {
    param(
        [Parameter(Mandatory = $true)][string]$InputValue,
        [Parameter(Mandatory = $true)][string[]]$Arguments
    )
    $globalCli = Get-Command vercel -ErrorAction SilentlyContinue
    if ($globalCli) {
        $InputValue | & $globalCli.Source @Arguments
    } else {
        $InputValue | & npx.cmd --yes "vercel@$vercelVersion" @Arguments
    }
}

function Save-LocalGeminiConfig {
    param([Parameter(Mandatory = $true)][string]$ApiKey)
    $values = [ordered]@{
        GEMINI_API_KEY = $ApiKey
        OCR2WORD_AI_MODEL = $model
    }
    $lines = New-Object 'System.Collections.Generic.List[string]'
    if (Test-Path -LiteralPath $localEnvFile) {
        foreach ($line in Get-Content -LiteralPath $localEnvFile) {
            if ($line -notmatch '^\s*(GEMINI_API_KEY|OCR2WORD_AI_MODEL)=') {
                $lines.Add($line)
            }
        }
    }
    foreach ($entry in $values.GetEnumerator()) {
        $lines.Add("$($entry.Key)=$($entry.Value)")
    }
    [System.IO.File]::WriteAllLines($localEnvFile, $lines, (New-Object System.Text.UTF8Encoding($false)))
}

$secureKey = Read-Host "Paste the Gemini API key (input is hidden)" -AsSecureString
$pointer = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureKey)
try {
    $apiKey = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($pointer)
} finally {
    [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($pointer)
}

if ([string]::IsNullOrWhiteSpace($apiKey) -or $apiKey.Length -lt 20 -or $apiKey -match '\s') {
    throw "The Gemini API key does not look valid. Create a key at https://aistudio.google.com/apikey."
}

Save-LocalGeminiConfig -ApiKey $apiKey
Write-Host "Saved the local Gemini configuration in the ignored .env file." -ForegroundColor Green

if (-not $LocalOnly) {
    $projectLink = Join-Path $projectRoot ".vercel\project.json"
    $repoLink = Join-Path $projectRoot ".vercel\repo.json"
    if (-not (Test-Path -LiteralPath $projectLink) -and -not (Test-Path -LiteralPath $repoLink)) {
        throw "This folder is not linked to Vercel. Run .\setup-vercel-local.ps1 or use -LocalOnly."
    }
    Push-Location $projectRoot
    try {
        Write-Host "Saving GEMINI_API_KEY to Vercel Production and Preview as a sensitive secret..." -ForegroundColor Cyan
        Invoke-VercelWithInput -InputValue $apiKey -Arguments @(
            "env", "add", "GEMINI_API_KEY", "production,preview",
            "--force", "--sensitive", "--visibility=secret", "--yes"
        )
        if ($LASTEXITCODE -ne 0) { throw "Could not save GEMINI_API_KEY to Vercel Production/Preview." }

        # Vercel does not allow Sensitive variables in Development because
        # `vercel env pull` must be able to read them for local development.
        # Regular Vercel variables are still encrypted at rest.
        Write-Host "Saving GEMINI_API_KEY to Vercel Development..." -ForegroundColor Cyan
        Invoke-VercelWithInput -InputValue $apiKey -Arguments @(
            "env", "add", "GEMINI_API_KEY", "development",
            "--force", "--no-sensitive", "--yes"
        )
        if ($LASTEXITCODE -ne 0) { throw "Could not save GEMINI_API_KEY to Vercel Development." }

        Invoke-Vercel -Arguments @(
            "env", "add", "OCR2WORD_AI_MODEL", "production,preview,development",
            "--value", $model, "--force", "--no-sensitive", "--visibility=config", "--yes"
        )
        if ($LASTEXITCODE -ne 0) { throw "Could not set OCR2WORD_AI_MODEL on Vercel." }

        Invoke-Vercel -Arguments @("pull", "--yes", "--environment=development")
        if ($LASTEXITCODE -ne 0) { throw "Could not refresh Vercel Development project settings." }
        Invoke-Vercel -Arguments @("env", "pull", "frontend/.env.local", "--yes", "--environment=development")
        if ($LASTEXITCODE -ne 0) { throw "Could not refresh the Development environment." }
    } finally {
        Pop-Location
    }
    Write-Host "Gemini is configured for Development, Preview and Production." -ForegroundColor Green
}

$apiKey = $null
Write-Host "Next: start the app and enable 'Ho tro voi AI mien phi'." -ForegroundColor Green
