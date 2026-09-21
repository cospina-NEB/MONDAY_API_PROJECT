# ============================================================
# sharefile_get_refresh_token.ps1
# One-time interactive setup for sharefile_user_report.ps1.
#
# ShareFile's OAuth2 password grant has no way to satisfy an MFA
# challenge, so any account with MFA enabled always gets back
# "invalid_grant: invalid username or password" from it. The
# Authorization Code flow instead has you log in (and clear MFA)
# in a real browser, then exchanges the resulting code for an
# access token + a long-lived refresh token. This script runs
# that flow once and saves the refresh token to .env, so the
# report script can silently mint new access tokens on every run
# via the refresh_token grant, without ever needing MFA again
# (until the refresh token itself is revoked or expires).
#
# Usage:
#   Add to .env:
#     SHAREFILE_SUBDOMAIN=yourcompany
#     SHAREFILE_CLIENT_ID=...
#     SHAREFILE_CLIENT_SECRET=...
#     SHAREFILE_REDIRECT_URI=https://secure.sharefile.com/oauth/oauthcomplete.aspx
#   The redirect URI must exactly match what's registered on this
#   API app in ShareFile's Admin > API/App Management. For API apps
#   without their own web server, this is ShareFile's own built-in
#   landing page (as above) - after login it displays the resulting
#   URL/code on screen rather than redirecting to your machine, so
#   this script has you paste that back in rather than running a
#   local listener.
#
#   .\scripts\sharefile_get_refresh_token.ps1
#
# It opens your default browser to ShareFile's login page. After
# you log in (and clear MFA), copy the full URL of the page you
# land on (or just the "code" value from it) and paste it back
# when prompted. It then exchanges the code for tokens and
# writes/updates SHAREFILE_REFRESH_TOKEN in .env.
# ============================================================

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ── Load .env ─────────────────────────────────────────────────
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

$Subdomain    = $env:SHAREFILE_SUBDOMAIN
$ClientId     = $env:SHAREFILE_CLIENT_ID
$ClientSecret = $env:SHAREFILE_CLIENT_SECRET
$RedirectUri  = $env:SHAREFILE_REDIRECT_URI

foreach ($pair in @(
    @("SHAREFILE_SUBDOMAIN", $Subdomain),
    @("SHAREFILE_CLIENT_ID", $ClientId),
    @("SHAREFILE_CLIENT_SECRET", $ClientSecret),
    @("SHAREFILE_REDIRECT_URI", $RedirectUri)
)) {
    if (-not $pair[1]) { throw "Set $($pair[0]) in .env first" }
}

# ── Step 1: send the user to ShareFile's login/consent page ──
$AuthUrl = "https://$Subdomain.sharefile.com/oauth/authorize" +
    "?response_type=code" +
    "&client_id=$([uri]::EscapeDataString($ClientId))" +
    "&redirect_uri=$([uri]::EscapeDataString($RedirectUri))"

Write-Host "Opening browser to log in to ShareFile (complete MFA there if prompted)..."
Write-Host "   $AuthUrl"
Start-Process $AuthUrl

# ── Step 2: get the authorization code back from the user ────
# The redirect URI is ShareFile's own landing page (not a local
# server we control), so after login it shows the code on screen
# / in the resulting URL rather than redirecting to this machine.
Write-Host ""
Write-Host "After you finish logging in, ShareFile will show a page with a 'code' value"
Write-Host "(either displayed directly, or as '?code=...' in the browser's address bar)."
$Pasted = Read-Host "Paste that code, or the full URL/page text containing it"

$Code = $Pasted.Trim()
if ($Code -match 'code=([^&\s"]+)') {
    $Code = $Matches[1]
}
$Code = [uri]::UnescapeDataString($Code)

if (-not $Code) {
    throw "No authorization code provided"
}

Write-Host "   Authorization code received."

# ── Step 3: exchange the code for an access token + refresh token ──
Write-Host "Exchanging code for tokens..."

$TokenUrl  = "https://$Subdomain.sharefile.com/oauth/token"
$TokenBody = @{
    grant_type    = "authorization_code"
    code          = $Code
    client_id     = $ClientId
    client_secret = $ClientSecret
    redirect_uri  = $RedirectUri
}

$TokenResponse = Invoke-RestMethod -Uri $TokenUrl -Method Post -Body $TokenBody -ContentType "application/x-www-form-urlencoded"

if (-not $TokenResponse.refresh_token) {
    throw "ShareFile token exchange did not return a refresh_token. Response: $($TokenResponse | ConvertTo-Json -Compress)"
}

$RefreshToken = $TokenResponse.refresh_token
Write-Host "   Refresh token received."

# ── Step 4: save/update SHAREFILE_REFRESH_TOKEN in .env ───────
$EnvLines = if (Test-Path $EnvFile) { Get-Content $EnvFile } else { @() }
$Found = $false
$NewLines = $EnvLines | ForEach-Object {
    if ($_ -match '^\s*SHAREFILE_REFRESH_TOKEN\s*=') {
        $Found = $true
        "SHAREFILE_REFRESH_TOKEN=$RefreshToken"
    } else {
        $_
    }
}
if (-not $Found) {
    $NewLines += "SHAREFILE_REFRESH_TOKEN=$RefreshToken"
}
Set-Content -Path $EnvFile -Value $NewLines -Encoding UTF8

Write-Host ""
Write-Host "Done! SHAREFILE_REFRESH_TOKEN saved to .env."
Write-Host "   You can now run .\scripts\sharefile_user_report.ps1"
