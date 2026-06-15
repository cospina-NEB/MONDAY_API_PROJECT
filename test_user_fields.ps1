# test_user_fields.ps1
# Introspects the Monday.com API to list all fields available on the User type.

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
  __type(name: "User") {
    fields {
      name
    }
  }
}
"@

$Body = @{ query = $Query } | ConvertTo-Json -Compress

$Response = Invoke-RestMethod `
    -Uri     "https://api.monday.com/v2" `
    -Method  Post `
    -Headers @{
        "Content-Type"  = "application/json"
        "Authorization" = $Token
        "API-Version"   = "2024-07"
    } `
    -Body $Body

if ($Response.PSObject.Properties['errors']) {
    Write-Host "API ERRORS:" -ForegroundColor Red
    $Response.errors | ConvertTo-Json | Write-Host
    exit 1
}

Write-Host ""
Write-Host "Fields available on the User type:" -ForegroundColor Cyan
Write-Host ("-" * 60)

$Response.data.__type.fields | Sort-Object name | ForEach-Object {
    Write-Host $_.name
}
