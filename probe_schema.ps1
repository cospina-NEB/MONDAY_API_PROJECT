# ============================================================
# probe_schema.ps1  –  Discover invitation-related API surface
# Uses Invoke-WebRequest + manual JSON parse to avoid the
# Invoke-RestMethod deserialization bug on deep introspection.
# ============================================================

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

function Invoke-GQL {
    param([string]$Query, [string]$ApiVersion = "2024-07")
    $Body    = @{ query = $Query } | ConvertTo-Json -Compress -Depth 10
    $headers = @{
        "Content-Type"  = "application/json"
        "Authorization" = $Token
        "API-Version"   = $ApiVersion
    }
    $raw  = Invoke-WebRequest -Uri "https://api.monday.com/v2" -Method Post -Headers $headers -Body $Body -UseBasicParsing
    return $raw.Content | ConvertFrom-Json
}

# ── 1. User type fields – current version (2024-07) ──────────
Write-Host "`n=== User type fields (API v2024-07) ===" -ForegroundColor Cyan
$q = 'query { __type(name: "User") { fields { name type { name kind ofType { name kind } } } } }'
try {
    $r = Invoke-GQL -Query $q -ApiVersion "2024-07"
    $r.data.__type.fields | Sort-Object name | ForEach-Object {
        $t = if ($_.type.name) { $_.type.name } else { $_.type.ofType.name }
        Write-Host ("  {0,-40} {1}" -f $_.name, $t)
    }
} catch { Write-Warning "v2024-07 introspection failed: $_" }

# ── 2. User type fields – latest stable version ───────────────
Write-Host "`n=== User type fields (API v2025-01) ===" -ForegroundColor Cyan
try {
    $r2 = Invoke-GQL -Query $q -ApiVersion "2025-01"
    $r2.data.__type.fields | Sort-Object name | ForEach-Object {
        $t = if ($_.type.name) { $_.type.name } else { $_.type.ofType.name }
        Write-Host ("  {0,-40} {1}" -f $_.name, $t)
    }
} catch { Write-Warning "v2025-01 introspection failed: $_" }

# ── 3. Probe user_connections ─────────────────────────────────
Write-Host "`n=== user_connections sample (first 3) ===" -ForegroundColor Cyan
$ucQuery = @'
query {
  user_connections(limit: 3) {
    id
    name
    email
  }
}
'@
try {
    $r3 = Invoke-GQL -Query $ucQuery -ApiVersion "2024-07"
    if ($r3.errors) {
        Write-Host "  Error: $($r3.errors[0].message)"
    } else {
        $r3.data | ConvertTo-Json -Depth 5
    }
} catch { Write-Warning "user_connections failed: $_" }

# ── 4. Check InviteUsersResult.invited_users sub-type ─────────
Write-Host "`n=== InviteUsersResult.invited_users type ===" -ForegroundColor Cyan
$iurQuery = 'query { __type(name: "InviteUsersResult") { fields { name type { name kind ofType { name kind fields { name } } } } } }'
try {
    $r4 = Invoke-GQL -Query $iurQuery -ApiVersion "2024-07"
    $r4.data.__type.fields | ForEach-Object {
        Write-Host "  $($_.name)  ->  type: $($_.type.name)$($_.type.ofType.name)"
    }
} catch { Write-Warning "InviteUsersResult introspection failed: $_" }

# ── 5. Single user sample – see raw fields returned ──────────
Write-Host "`n=== Raw fields on a single user (me) ===" -ForegroundColor Cyan
$meQuery = @'
query {
  me {
    id
    name
    email
    created_at
    last_activity
    is_admin
    is_guest
    is_view_only
    enabled
    join_date
    title
    url
    account { id name }
  }
}
'@
try {
    $r5 = Invoke-GQL -Query $meQuery -ApiVersion "2024-07"
    $r5.data.me | ConvertTo-Json -Depth 3
} catch { Write-Warning "me query failed: $_" }

Write-Host "`nDone." -ForegroundColor Green
