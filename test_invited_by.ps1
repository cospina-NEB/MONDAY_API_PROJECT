# test_invited_by.ps1
# Runs a quick API check to see what invited_by actually returns.

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$EnvFile   = Join-Path $ScriptDir ".env"
if (Test-Path $EnvFile) {
    Get-Content $EnvFile | ForEach-Object {
        if ($_ -match '^\s*([^#][^=]+)=(.*)$') {
            [System.Environment]::SetEnvironmentVariable($Matches[1].Trim(), $Matches[2].Trim(), "Process")
        }
    }
}

$Token = $env:MONDAY_API_TOKEN
if (-not $Token) { throw "Set MONDAY_API_TOKEN in .env first" }

$Query = @"
query {
  users(limit: 20, kind: all) {
    id
    name
    email
    invited_by {
      id
      name
      email
    }
  }
}
"@

$Body = @{ query = $Query } | ConvertTo-Json -Compress

$Response = Invoke-RestMethod `
    -Uri     "https://api.monday.com/v2" `
    -Method  Post `
    -Headers @{
        "Content-Type" = "application/json"
        "Authorization" = $Token
        "API-Version"  = "2024-07"
    } `
    -Body $Body

if ($Response.PSObject.Properties['errors']) {
    Write-Host "API ERRORS:" -ForegroundColor Red
    $Response.errors | ConvertTo-Json | Write-Host
    exit 1
}

Write-Host ""
Write-Host "invited_by results for first 20 users:" -ForegroundColor Cyan
Write-Host ("-" * 60)

foreach ($u in $Response.data.users) {
    $inviter = if ($u.invited_by) { "$($u.invited_by.name) <$($u.invited_by.email)>" } else { "(null)" }
    Write-Host "$($u.name.PadRight(30)) invited_by: $inviter"
}
