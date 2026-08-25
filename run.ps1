param(
    [int]$Port = 8000
)

# Keep this file encoded as UTF-8 with BOM for Windows PowerShell 5.1.

$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$venvPython = Join-Path $projectRoot ".venv\Scripts\python.exe"

if (-not (Test-Path -LiteralPath $venvPython)) {
    throw "Chưa có môi trường chạy. Hãy chạy .\setup.ps1 trước."
}

Set-Location -LiteralPath $projectRoot
$localEnvFile = Join-Path $projectRoot ".env"
if (Test-Path -LiteralPath $localEnvFile) {
    foreach ($line in Get-Content -LiteralPath $localEnvFile) {
        if ($line -match '^\s*(GEMINI_API_KEY|OCR2WORD_AI_MODEL)=(.*)$') {
            [Environment]::SetEnvironmentVariable($Matches[1], $Matches[2], "Process")
        }
    }
}
Write-Host "OCR2Word đang chạy tại http://127.0.0.1:$Port" -ForegroundColor Green
Write-Host "Nhấn Ctrl+C để dừng."
& $venvPython -m uvicorn app.main:app --host 127.0.0.1 --port $Port
