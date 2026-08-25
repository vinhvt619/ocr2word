param(
    [Parameter(Mandatory = $true)][string]$FilePath,
    [string]$BaseUrl = "http://127.0.0.1:3000",
    [ValidateRange(60, 3600)][int]$TimeoutSeconds = 1800,
    [switch]$AiAssist,
    [ValidateSet("arn1", "bom1", "cdg1", "cle1", "cpt1", "dub1", "dxb1", "fra1", "gru1", "hkg1", "hnd1", "iad1", "icn1", "kix1", "lhr1", "pdx1", "sfo1", "sin1", "syd1", "yul1")]
    [string]$ExpectedQueueRegion = "sin1",
    [string]$OutputDirectory = "output\vercel-parity"
)

# ASCII-only on purpose: Windows PowerShell 5.1 can parse this file without BOM.
$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$source = Get-Item -LiteralPath $FilePath -ErrorAction Stop
$base = $BaseUrl.TrimEnd('/')

function Assert-True {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw $Message }
}

function Invoke-JsonRequest {
    param(
        [Parameter(Mandatory = $true)][string]$Method,
        [Parameter(Mandatory = $true)][string]$Uri,
        [object]$Body = $null,
        [int]$TimeoutSec = 120
    )
    $parameters = @{
        Method = $Method
        Uri = $Uri
        TimeoutSec = $TimeoutSec
        ErrorAction = "Stop"
    }
    if ($null -ne $Body) {
        $parameters.ContentType = "application/json"
        $parameters.Body = ($Body | ConvertTo-Json -Depth 10 -Compress)
    }
    return Invoke-RestMethod @parameters
}

Write-Host "Checking the production-like service graph..." -ForegroundColor Cyan
$status = Invoke-JsonRequest -Method "GET" -Uri "$base/api/status" -TimeoutSec 30
Assert-True ($status.deployment.mode -eq "vercel") "The controller is not using the Vercel job flow."
Assert-True ([bool]$status.deployment.local_parity) "The server is not running in local parity mode. Use .\run-vercel-local.ps1."
Assert-True ([bool]$status.deployment.available) $status.deployment.message
Assert-True ($status.runtime.controller.node -match '^24\.') "Controller must run Node.js 24.x."
Assert-True ($status.runtime.worker.system -eq "Linux") "OCR worker must run in Linux."
Assert-True ([bool]$status.runtime.worker.container) "OCR worker is not running from the Vercel container image."
Assert-True ($status.runtime.queue.region -eq $ExpectedQueueRegion) "Queue region must be $ExpectedQueueRegion for this parity profile."
$languages = @($status.ocr.languages)
Assert-True ($languages -contains "vie") "Tesseract Vietnamese model is missing."
Assert-True ($languages -contains "eng") "Tesseract English model is missing."
if ($AiAssist) {
    Assert-True ([bool]$status.ai.available) $status.ai.message
    Assert-True ($status.ai.provider -match "Google Gemini") "AI is not using Google Gemini."
}

$options = @{
    language = "vie+eng"
    ocr_mode = "auto"
    dpi = 300
    page_breaks = $true
    preserve_images = $true
    preserve_layout = $true
    detect_tables = $true
    ai_assist = [bool]$AiAssist
}
$createBody = @{
    files = @(@{ name = $source.Name; size = $source.Length })
    options = $options
}

Write-Host "Creating a real Vercel Blob/Redis job..." -ForegroundColor Cyan
$created = Invoke-JsonRequest -Method "POST" -Uri "$base/api/jobs" -Body $createBody
Assert-True (-not [string]::IsNullOrWhiteSpace($created.job_id)) "Job creation returned no job_id."
Assert-True (@($created.uploads).Count -eq 1) "Job creation returned no signed upload URL."
$upload = @($created.uploads)[0]

Write-Host "Uploading through the real signed Private Blob URL..." -ForegroundColor Cyan
Invoke-WebRequest -Method Put -Uri $upload.url -InFile $source.FullName -ContentType $upload.content_type -UseBasicParsing -TimeoutSec 3600 | Out-Null

Write-Host "Publishing the job to the real Vercel Queue..." -ForegroundColor Cyan
[void](Invoke-JsonRequest -Method "POST" -Uri "$base/api/jobs/$($created.job_id)/start" -Body @{})

$deadline = (Get-Date).AddSeconds($TimeoutSeconds)
$lastPercent = -1
$samples = New-Object 'System.Collections.Generic.List[object]'
$job = $null
while ((Get-Date) -lt $deadline) {
    Start-Sleep -Seconds 2
    $job = Invoke-JsonRequest -Method "GET" -Uri "$base/api/jobs/$($created.job_id)" -TimeoutSec 30
    $percent = [int]$job.percent
    Assert-True ($percent -ge $lastPercent) "Progress moved backwards from $lastPercent to $percent."
    if ($percent -ne $lastPercent) {
        Write-Host ("[{0,3}%] {1}" -f $percent, $job.message)
        $samples.Add([pscustomobject]@{
            at = (Get-Date).ToString("o")
            percent = $percent
            status = $job.status
            message = $job.message
        })
        $lastPercent = $percent
    }
    if ($job.status -eq "error") {
        throw "Conversion failed: $($job.error)"
    }
    if ($job.status -eq "complete") { break }
}

Assert-True ($null -ne $job -and $job.status -eq "complete") "Timed out after $TimeoutSeconds seconds."
Assert-True ($job.percent -eq 100) "Completed job did not report 100 percent."
Assert-True (-not [string]::IsNullOrWhiteSpace($job.download_url)) "Completed job returned no signed download URL."

$outputRoot = $OutputDirectory
if (-not [System.IO.Path]::IsPathRooted($outputRoot)) {
    $outputRoot = Join-Path $projectRoot $outputRoot
}
[void](New-Item -ItemType Directory -Path $outputRoot -Force)
$safeName = [System.IO.Path]::GetFileName([string]$job.download_name)
$resultPath = Join-Path $outputRoot $safeName

Write-Host "Downloading the generated document through its signed URL..." -ForegroundColor Cyan
Invoke-WebRequest -Method Get -Uri $job.download_url -OutFile $resultPath -UseBasicParsing -TimeoutSec 600
Assert-True ((Get-Item -LiteralPath $resultPath).Length -gt 0) "Downloaded result is empty."

Add-Type -AssemblyName System.IO.Compression.FileSystem
$archive = [System.IO.Compression.ZipFile]::OpenRead($resultPath)
try {
    $documentXml = $archive.Entries | Where-Object { $_.FullName -eq "word/document.xml" }
    Assert-True ($null -ne $documentXml) "Result is not a valid editable DOCX package."
} finally {
    $archive.Dispose()
}

$reportPath = Join-Path $outputRoot "last-parity-report.json"
$report = [ordered]@{
    verified_at = (Get-Date).ToString("o")
    base_url = $base
    source_file = $source.Name
    source_bytes = $source.Length
    job_id = $created.job_id
    result_file = $resultPath
    ai_assist = [bool]$AiAssist
    controller_node = $status.runtime.controller.node
    worker_python = $status.runtime.worker.python
    worker_system = $status.runtime.worker.system
    queue_region = $status.runtime.queue.region
    progress = $samples
    conversion_report = $job.report
}
$report | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $reportPath -Encoding UTF8

Write-Host "PASS: the Vercel-like flow completed end to end." -ForegroundColor Green
Write-Host "DOCX: $resultPath" -ForegroundColor Green
Write-Host "Report: $reportPath" -ForegroundColor Green
