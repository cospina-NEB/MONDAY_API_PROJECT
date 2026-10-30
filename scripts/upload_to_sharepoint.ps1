# upload_to_sharepoint.ps1
# Uploads the generated CSV and HTML reports to SharePoint via Microsoft Graph API.
# Reads credentials from env vars (set by .env locally, or GitHub Secrets in CI).
#
# Defaults upload today's Coral User Report (CSV/HTML/XLSX). Other reports pass
# -ReportsDir/-Files/-SubFolder. -Download fetches files from the same SharePoint
# folder instead of uploading (used by the RCM export to restore its watermark).
#
#   .\scripts\upload_to_sharepoint.ps1
#   .\scripts\upload_to_sharepoint.ps1 -ReportsDir Reports\monday_rcm -SubFolder rcm -Download rcm_state.json
#   .\scripts\upload_to_sharepoint.ps1 -ReportsDir Reports\monday_rcm -SubFolder rcm -Files a.csv,rcm_state.json

param(
    [string]$ReportsDir = "Reports\monday",   # relative to project root
    [string[]]$Files,                         # names inside ReportsDir; default = today's Coral User Report
    [string]$SubFolder,                       # appended to SHAREPOINT_FOLDER
    [string[]]$Download                       # download these from SharePoint instead of uploading
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# Load .env if running locally
$ScriptDir   = Split-Path -Parent $MyInvocation.MyCommand.Path
$ProjectRoot = Split-Path -Parent $ScriptDir
$EnvFile     = Join-Path $ProjectRoot ".env"
if (Test-Path $EnvFile) {
    Get-Content $EnvFile | ForEach-Object {
        if ($_ -match '^\s*([^#][^=]+)=(.*)$') {
            [System.Environment]::SetEnvironmentVariable($Matches[1].Trim(), $Matches[2].Trim(), "Process")
        }
    }
}

# Validate required env vars
$Required = @('SHAREPOINT_TENANT_ID','SHAREPOINT_CLIENT_ID','SHAREPOINT_CLIENT_SECRET','SHAREPOINT_SITE_URL','SHAREPOINT_FOLDER')
foreach ($Var in $Required) {
    if (-not [System.Environment]::GetEnvironmentVariable($Var)) { throw "Missing required env var: $Var" }
}

$TenantId     = $env:SHAREPOINT_TENANT_ID
$ClientId     = $env:SHAREPOINT_CLIENT_ID
$ClientSecret = $env:SHAREPOINT_CLIENT_SECRET
$SiteUrl      = $env:SHAREPOINT_SITE_URL.TrimEnd('/')
$FolderPath   = $env:SHAREPOINT_FOLDER.Trim('/')
if ($SubFolder) { $FolderPath = "$FolderPath/$($SubFolder.Trim('/'))" }

# 1. Authenticate via OAuth2 client credentials
Write-Host "Authenticating with Microsoft Graph..."
$TokenResponse = Invoke-RestMethod `
    -Uri         "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" `
    -Method      Post `
    -ContentType "application/x-www-form-urlencoded" `
    -Body        @{
        grant_type    = "client_credentials"
        client_id     = $ClientId
        client_secret = $ClientSecret
        scope         = "https://graph.microsoft.com/.default"
    }
$AccessToken = $TokenResponse.access_token
$AuthHeader  = @{ Authorization = "Bearer $AccessToken" }

# 2. Resolve SharePoint site ID
$SiteUri  = [System.Uri]$SiteUrl
$Hostname = $SiteUri.Host
$SitePath = $SiteUri.AbsolutePath.TrimEnd('/')

Write-Host "Resolving SharePoint site: $Hostname$SitePath"
$SiteInfo = Invoke-RestMethod `
    -Uri     "https://graph.microsoft.com/v1.0/sites/${Hostname}:${SitePath}" `
    -Headers $AuthHeader
$SiteId = $SiteInfo.id
Write-Host "  Site ID: $SiteId"

$ReportsDir = Join-Path $ProjectRoot $ReportsDir
if (-not (Test-Path $ReportsDir)) { New-Item -ItemType Directory -Path $ReportsDir | Out-Null }

# Encode each path segment individually, then rejoin with /
$EncodedFolder = (($FolderPath -split '/') | ForEach-Object { [System.Uri]::EscapeDataString($_) }) -join '/'
function Get-DriveItemUri {
    param([string]$FileName)
    $EncodedName = [System.Uri]::EscapeDataString($FileName)
    return "https://graph.microsoft.com/v1.0/sites/$SiteId/drive/root:/$EncodedFolder/${EncodedName}:/content"
}

# 3a. Download mode
if ($Download) {
    foreach ($File in $Download) {
        $Target = Join-Path $ReportsDir $File
        Write-Host "Downloading $File..."
        try {
            Invoke-RestMethod -Uri (Get-DriveItemUri $File) -Headers $AuthHeader -OutFile $Target
            Write-Host "  -> $Target"
        } catch {
            if ($_.Exception.Response -and [int]$_.Exception.Response.StatusCode -eq 404) {
                Write-Warning "$File not found in SharePoint (first run?), skipping"
            } else { throw }
        }
    }
    return
}

# 3b. Upload (default: today's CSV, HTML, and Excel from Reports\monday)
if (-not $Files) {
    $Date  = Get-Date -Format "yyyy-MM-dd"
    $Files = @("coral_user_report_$Date.csv", "coral_user_report_$Date.html", "coral_user_report_$Date.xlsx")
}
$Uploaded = 0

foreach ($File in $Files) {
    $FullPath = Join-Path $ReportsDir $File
    if (-not (Test-Path $FullPath)) {
        Write-Warning "File not found, skipping: $File"
        continue
    }

    $FileName  = Split-Path -Leaf $FullPath
    $UploadUri = Get-DriveItemUri $FileName

    Write-Host "Uploading $FileName..."
    $FileBytes = [System.IO.File]::ReadAllBytes($FullPath)
    $Result    = Invoke-RestMethod `
        -Uri     $UploadUri `
        -Method  Put `
        -Headers ($AuthHeader + @{ "Content-Type" = "application/octet-stream" }) `
        -Body    $FileBytes
    Write-Host "  -> $($Result.webUrl)"
    $Uploaded++
}

Write-Host ""
if ($Uploaded -eq 0) {
    Write-Warning "No files were uploaded. Were the report files generated first?"
} else {
    Write-Host "Done! $Uploaded file(s) uploaded to SharePoint."
}
