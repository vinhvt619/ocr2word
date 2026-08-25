param(
    [ValidateSet("fast", "best")]
    [string]$Model = "best"
)

# Keep this file encoded as UTF-8 with BOM for Windows PowerShell 5.1.

$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$target = if ($env:OCR2WORD_TESSDATA_DIR) {
    $env:OCR2WORD_TESSDATA_DIR
} elseif ($env:LOCALAPPDATA) {
    Join-Path $env:LOCALAPPDATA "OCR2Word\tessdata"
} else {
    Join-Path $projectRoot ".tessdata"
}
New-Item -ItemType Directory -Force -Path $target | Out-Null

$repository = if ($Model -eq "best") { "tessdata_best" } else { "tessdata_fast" }
$languages = @("vie", "eng", "osd")
foreach ($language in $languages) {
    $destination = Join-Path $target "$language.traineddata"
    $url = "https://raw.githubusercontent.com/tesseract-ocr/$repository/main/$language.traineddata"
    Write-Host "Đang tải $language ($Model)..." -ForegroundColor Cyan
    Invoke-WebRequest -Uri $url -OutFile $destination
    if ((Get-Item -LiteralPath $destination).Length -lt 100000) {
        throw "File dữ liệu $language tải về không hợp lệ."
    }
}

Write-Host "Đã lưu dữ liệu OCR tại $target" -ForegroundColor Green
