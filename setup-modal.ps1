param(
    [switch]$SkipInstall,
    [switch]$SkipLogin,
    [switch]$SkipSecret,
    [switch]$PublicSite
)

# ASCII-only on purpose: Windows PowerShell 5.1 can parse this file without BOM.
$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$pythonExe = Join-Path $projectRoot ".venv\Scripts\python.exe"
$modalExe = Join-Path $projectRoot ".venv\Scripts\modal.exe"
$requirementsFile = Join-Path $projectRoot "requirements-modal.txt"
$localEnvFile = Join-Path $projectRoot ".env"
$secretName = "ocr2word-secrets"

function ConvertFrom-SecureValue {
    param([Parameter(Mandatory = $true)][Security.SecureString]$Value)
    $pointer = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($Value)
    try {
        return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($pointer)
    } finally {
        [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($pointer)
    }
}

function Read-DotEnvValue {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Name
    )
    if (-not (Test-Path -LiteralPath $Path)) { return $null }
    foreach ($line in Get-Content -LiteralPath $Path) {
        if ($line -match "^\s*$([Regex]::Escape($Name))\s*=\s*(.*)\s*$") {
            $value = $Matches[1].Trim()
            if ($value.Length -ge 2) {
                if (($value[0] -eq '"' -and $value[$value.Length - 1] -eq '"') -or
                    ($value[0] -eq "'" -and $value[$value.Length - 1] -eq "'")) {
                    $value = $value.Substring(1, $value.Length - 2)
                }
            }
            return $value
        }
    }
    return $null
}

if (-not (Test-Path -LiteralPath $pythonExe)) {
    throw "Python virtual environment is missing. Run .\setup.ps1 first."
}

$env:PYTHONUTF8 = "1"
$env:PIP_PROGRESS_BAR = "off"
$env:PIP_DISABLE_PIP_VERSION_CHECK = "1"

if (-not $SkipInstall) {
    Write-Host "Installing Modal deployment dependencies..." -ForegroundColor Cyan
    & $pythonExe -m pip install --disable-pip-version-check --progress-bar off -r $requirementsFile
    if ($LASTEXITCODE -ne 0) { throw "Modal dependencies could not be installed." }
}

if (-not (Test-Path -LiteralPath $modalExe)) {
    throw "Modal CLI is missing. Run this script again without -SkipInstall."
}

& $modalExe token info *> $null
if ($LASTEXITCODE -ne 0) {
    if ($SkipLogin) {
        throw "Modal CLI is not authenticated. Run .\setup-modal.ps1 without -SkipLogin."
    }
    Write-Host "Opening the browser to sign in or create a Modal account..." -ForegroundColor Cyan
    & $modalExe setup
    if ($LASTEXITCODE -ne 0) { throw "Modal login did not complete." }
    & $modalExe token info *> $null
    if ($LASTEXITCODE -ne 0) { throw "Modal authentication could not be verified." }
}

if ($SkipSecret) {
    Write-Host "Modal CLI is installed and authenticated. Secret setup was skipped." -ForegroundColor Green
    exit 0
}

$secretValues = [ordered]@{}
if (-not $PublicSite) {
    $username = (Read-Host "Family website username [ocr2word]").Trim()
    if ([string]::IsNullOrWhiteSpace($username)) { $username = "ocr2word" }
    if ($username -notmatch '^[A-Za-z0-9._-]{3,64}$') {
        throw "Username must be 3-64 characters using letters, numbers, dot, dash or underscore."
    }

    $securePassword = Read-Host "Family website password (hidden, at least 12 characters)" -AsSecureString
    $securePasswordAgain = Read-Host "Repeat the family website password" -AsSecureString
    $password = ConvertFrom-SecureValue -Value $securePassword
    $passwordAgain = ConvertFrom-SecureValue -Value $securePasswordAgain
    if ($password.Length -lt 12) { throw "Website password must contain at least 12 characters." }
    if ($password -cne $passwordAgain) { throw "The two website passwords do not match." }

    $secretValues.OCR2WORD_BASIC_USER = $username
    $secretValues.OCR2WORD_BASIC_PASSWORD = $password
}
$geminiKey = Read-DotEnvValue -Path $localEnvFile -Name "GEMINI_API_KEY"
$geminiModel = Read-DotEnvValue -Path $localEnvFile -Name "OCR2WORD_AI_MODEL"
if (-not [string]::IsNullOrWhiteSpace($geminiKey)) {
    $secretValues.GEMINI_API_KEY = $geminiKey
    $secretValues.OCR2WORD_AI_MODEL = if ([string]::IsNullOrWhiteSpace($geminiModel)) {
        "gemini-3.7-flash"
    } else {
        $geminiModel
    }
}
if ([string]::IsNullOrWhiteSpace($geminiKey)) {
    throw "GEMINI_API_KEY is missing from .env."
}

$tempJson = Join-Path ([IO.Path]::GetTempPath()) ("ocr2word-modal-secret-{0}.json" -f [Guid]::NewGuid().ToString("N"))
try {
    $json = $secretValues | ConvertTo-Json -Compress
    [IO.File]::WriteAllText($tempJson, $json, (New-Object Text.UTF8Encoding($false)))
    Write-Host "Saving the private website credentials to Modal Secret..." -ForegroundColor Cyan
    & $modalExe secret create $secretName --from-json $tempJson --force
    if ($LASTEXITCODE -ne 0) { throw "Modal Secret could not be saved." }
} finally {
    if (Test-Path -LiteralPath $tempJson) {
        Remove-Item -LiteralPath $tempJson -Force
    }
    $password = $null
    $passwordAgain = $null
    $geminiKey = $null
    $json = $null
    $secretValues = $null
}

Write-Host "Modal setup is ready." -ForegroundColor Green
Write-Host "Before deployment, set a hard Workspace Budget below the monthly free credit in the Modal dashboard." -ForegroundColor Yellow
Write-Host "Recommended cap: USD 20. Then run: .\deploy-modal.ps1 -BudgetConfirmed" -ForegroundColor Yellow
