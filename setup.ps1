param(
    [switch]$InstallTesseract,
    [ValidateSet("fast", "best")]
    [string]$OcrModel = "best"
)

# Keep this file encoded as UTF-8 with BOM for Windows PowerShell 5.1.

$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location -LiteralPath $projectRoot
$env:PYTHONUTF8 = "1"
$env:PIP_PROGRESS_BAR = "off"

function Find-Python {
    if ($env:OCR2WORD_PYTHON -and (Test-Path -LiteralPath $env:OCR2WORD_PYTHON)) {
        return @{ Command = $env:OCR2WORD_PYTHON; Arguments = @() }
    }
    $pyLauncher = Get-Command py -ErrorAction SilentlyContinue
    if ($pyLauncher) {
        return @{ Command = $pyLauncher.Source; Arguments = @("-3") }
    }
    $python = Get-Command python -ErrorAction SilentlyContinue
    if ($python) {
        return @{ Command = $python.Source; Arguments = @() }
    }
    throw "Không tìm thấy Python 3.11+. Hãy cài Python rồi chạy lại, hoặc đặt OCR2WORD_PYTHON tới python.exe."
}

$pythonInfo = Find-Python
$venvPython = Join-Path $projectRoot ".venv\Scripts\python.exe"
if (-not (Test-Path -LiteralPath $venvPython)) {
    Write-Host "Đang tạo môi trường Python..." -ForegroundColor Cyan
    & $pythonInfo.Command @($pythonInfo.Arguments) -m venv ".venv"
}

Write-Host "Đang cài thư viện..." -ForegroundColor Cyan
& $venvPython -m pip install --upgrade pip
& $venvPython -m pip install -r "requirements.txt"

if ($InstallTesseract) {
    $tesseract = Get-Command tesseract -ErrorAction SilentlyContinue
    if (-not $tesseract) {
        $winget = Get-Command winget -ErrorAction SilentlyContinue
        if (-not $winget) {
            throw "Không tìm thấy winget. Hãy cài Tesseract thủ công rồi chạy scripts\download_languages.ps1."
        }
        Write-Host "Đang cài Tesseract OCR..." -ForegroundColor Cyan
        & $winget.Source install --id UB-Mannheim.TesseractOCR -e --accept-source-agreements --accept-package-agreements
    }
    & (Join-Path $projectRoot "scripts\download_languages.ps1") -Model $OcrModel
    & $venvPython -X utf8 -c "from app.services.ocr import get_ocr_engine_info; info = get_ocr_engine_info(); print(info.message); raise SystemExit(0 if info.available else 1)"
    if ($LASTEXITCODE -ne 0) {
        throw "Tesseract đã cài nhưng ứng dụng chưa thể sử dụng dữ liệu ngôn ngữ OCR."
    }
}

Write-Host ""
Write-Host "Cài đặt hoàn tất." -ForegroundColor Green
Write-Host "Chạy ứng dụng bằng: .\run.ps1"
if (-not $InstallTesseract) {
    Write-Host "Để OCR PDF scan chất lượng cao, chạy lại: .\setup.ps1 -InstallTesseract -OcrModel best" -ForegroundColor Yellow
}
Write-Host "Để bật Gemini AI, chạy: .\setup-gemini.ps1" -ForegroundColor Yellow
