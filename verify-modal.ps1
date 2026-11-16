param(
    [string]$Url,
    [string]$Username = "ocr2word",
    [string]$TestFile,
    [switch]$PublicSite
)

# ASCII-only on purpose: Windows PowerShell 5.1 can parse this file without BOM.
$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$urlFile = Join-Path $projectRoot ".modal-url"

if ([string]::IsNullOrWhiteSpace($Url) -and (Test-Path -LiteralPath $urlFile)) {
    $Url = (Get-Content -LiteralPath $urlFile -Raw).Trim()
}
if ([string]::IsNullOrWhiteSpace($Url)) {
    throw "Modal URL is missing. Pass -Url or deploy first."
}
$Url = $Url.TrimEnd('/')

$health = Invoke-RestMethod -Uri "$Url/healthz" -Method Get -TimeoutSec 180
if (-not $health.ok) { throw "The public Modal health endpoint did not pass." }
Write-Host "Public health check passed." -ForegroundColor Green

$headers = @{}
$token = $null
if (-not $PublicSite) {
    $securePassword = Read-Host "Family website password" -AsSecureString
    $pointer = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePassword)
    try {
        $password = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($pointer)
    } finally {
        [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($pointer)
    }
    $token = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes("${Username}:${password}"))
    $headers.Authorization = "Basic $token"
}
$status = Invoke-RestMethod -Uri "$Url/api/status" -Method Get -Headers $headers -TimeoutSec 180
if ($status.deployment.mode -ne "modal" -or -not $status.ocr.available) {
    throw "The protected OCR2Word status endpoint did not pass."
}
Write-Host "Protected status check passed: OCR2Word $($status.version)." -ForegroundColor Green

if ([string]::IsNullOrWhiteSpace($TestFile)) {
    Write-Host "Website verification completed." -ForegroundColor Green
    exit 0
}

$source = (Resolve-Path -LiteralPath $TestFile).Path
if ((Get-Item -LiteralPath $source).Length -gt 80MB) {
    throw "Test file is larger than the 80 MB Modal limit."
}

Add-Type -AssemblyName System.Net.Http
$client = [Net.Http.HttpClient]::new()
$client.Timeout = [TimeSpan]::FromMinutes(10)
if (-not [string]::IsNullOrWhiteSpace($token)) {
    $client.DefaultRequestHeaders.Authorization = [Net.Http.Headers.AuthenticationHeaderValue]::new("Basic", $token)
}
$multipart = [Net.Http.MultipartFormDataContent]::new()
try {
    $fileBytes = [IO.File]::ReadAllBytes($source)
    $fileContent = [Net.Http.ByteArrayContent]::new([byte[]]$fileBytes)
    $fileContent.Headers.ContentType = [Net.Http.Headers.MediaTypeHeaderValue]::new("application/octet-stream")
    $multipart.Add($fileContent, "files", [IO.Path]::GetFileName($source))
    $options = [ordered]@{
        language = "vie+eng"
        ocr_mode = "auto"
        dpi = 300
        page_breaks = $true
        preserve_images = $true
        preserve_layout = $true
        detect_tables = $true
    } | ConvertTo-Json -Compress
    $optionsContent = [Net.Http.StringContent]::new($options, [Text.Encoding]::UTF8)
    $multipart.Add($optionsContent, "options")

    Write-Host "Uploading the test file and starting a real Modal job..." -ForegroundColor Cyan
    $createResponse = $client.PostAsync("$Url/api/modal/jobs", $multipart).GetAwaiter().GetResult()
    $createBody = $createResponse.Content.ReadAsStringAsync().GetAwaiter().GetResult()
    if (-not $createResponse.IsSuccessStatusCode) {
        throw "Modal job creation failed: $createBody"
    }
    $created = $createBody | ConvertFrom-Json
    $jobId = $created.job_id
    $deadline = [DateTime]::UtcNow.AddMinutes(40)
    $lastPercent = -1
    do {
        Start-Sleep -Milliseconds 1000
        $job = Invoke-RestMethod -Uri "$Url/api/modal/jobs/$jobId" -Method Get -Headers $headers -TimeoutSec 180
        if ([int]$job.percent -ne $lastPercent) {
            Write-Host ("{0}% - {1}" -f $job.percent, $job.message)
            $lastPercent = [int]$job.percent
        }
        if ($job.status -eq "error") { throw $job.error }
        if ($job.status -eq "complete") { break }
    } while ([DateTime]::UtcNow -lt $deadline)
    if ($job.status -ne "complete") { throw "Modal test job exceeded 40 minutes." }

    $resultBytes = $client.GetByteArrayAsync($job.download_url).GetAwaiter().GetResult()
    $outputRoot = Join-Path $projectRoot "output\modal-web"
    [void][IO.Directory]::CreateDirectory($outputRoot)
    $outputName = if ([string]::IsNullOrWhiteSpace($job.download_name)) {
        "ocr2word-modal-result.docx"
    } else {
        [IO.Path]::GetFileName([string]$job.download_name)
    }
    $target = Join-Path $outputRoot $outputName
    [IO.File]::WriteAllBytes($target, $resultBytes)
    Write-Host "Real conversion passed. Result: $target" -ForegroundColor Green
} finally {
    $multipart.Dispose()
    $client.Dispose()
    $password = $null
    $token = $null
}
