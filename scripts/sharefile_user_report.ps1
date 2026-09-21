# ============================================================
# sharefile_user_report.ps1
# Generates a User report for ShareFile (Citrix), for AES-315.
# Replicates the manually-exported ShareFile "UserList.xlsx"
# (Admin > Users > Export) via the ShareFile REST API v3.
#
# Usage:
#   Add to .env:
#     SHAREFILE_SUBDOMAIN=yourcompany
#     SHAREFILE_CLIENT_ID=...
#     SHAREFILE_CLIENT_SECRET=...
#     SHAREFILE_REFRESH_TOKEN=...   (from a one-time run of
#                                     scripts\sharefile_get_refresh_token.ps1 -
#                                     see that script for why: the password
#                                     grant can't satisfy MFA)
#   .\scripts\sharefile_user_report.ps1
#
#   First run against a real account: add -DumpRawSample to write
#   the raw JSON of one Employee and one Client to Reports\, so the
#   field-name mapping below (see Get-Field candidate lists) can be
#   verified/corrected against your account's actual schema. The
#   ShareFile API's public docs do not publish a full AccountUser
#   field list, so Get-Field tries several likely names per column
#   and falls back to blank + a one-time warning if none match.
#
# Output: sharefile_user_report_YYYY-MM-DD.csv / .html / .xlsx
# ============================================================

param(
    [switch]$DumpRawSample
)

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

# ── Config ────────────────────────────────────────────────────
$Subdomain     = $env:SHAREFILE_SUBDOMAIN
$ClientId      = $env:SHAREFILE_CLIENT_ID
$ClientSecret  = $env:SHAREFILE_CLIENT_SECRET
$RefreshToken  = $env:SHAREFILE_REFRESH_TOKEN

foreach ($pair in @(
    @("SHAREFILE_SUBDOMAIN", $Subdomain),
    @("SHAREFILE_CLIENT_ID", $ClientId),
    @("SHAREFILE_CLIENT_SECRET", $ClientSecret),
    @("SHAREFILE_REFRESH_TOKEN", $RefreshToken)
)) {
    if (-not $pair[1]) { throw "Set $($pair[0]) in .env first - run .\scripts\sharefile_get_refresh_token.ps1 to obtain SHAREFILE_REFRESH_TOKEN" }
}

$Date       = Get-Date -Format "yyyy-MM-dd"
$ReportsDir = Join-Path $ProjectRoot "Reports\sharefile"
if (-not (Test-Path $ReportsDir)) { New-Item -ItemType Directory -Path $ReportsDir | Out-Null }
$OutputFile = Join-Path $ReportsDir "sharefile_user_report_$Date.csv"

# ── Step 0: OAuth2 refresh-token grant ───────────────────────
# Uses the refresh token from a one-time interactive login (see
# scripts\sharefile_get_refresh_token.ps1) instead of a password
# grant, since password grant can't satisfy MFA.
Write-Host "Authenticating to ShareFile..."

$TokenUrl  = "https://$Subdomain.sharefile.com/oauth/token"
$TokenBody = @{
    grant_type    = "refresh_token"
    client_id     = $ClientId
    client_secret = $ClientSecret
    refresh_token = $RefreshToken
}

$TokenResponse = Invoke-RestMethod -Uri $TokenUrl -Method Post -Body $TokenBody -ContentType "application/x-www-form-urlencoded"

if (-not $TokenResponse.access_token) {
    throw "ShareFile authentication failed - no access_token returned. Response: $($TokenResponse | ConvertTo-Json -Compress)"
}

$AccessToken = $TokenResponse.access_token
$ApiCp       = if ($TokenResponse.apicp) { $TokenResponse.apicp } else { "sf-api.com" }
$ApiSubdomain = if ($TokenResponse.subdomain) { $TokenResponse.subdomain } else { $Subdomain }
$ApiBase     = "https://$ApiSubdomain.$ApiCp/sf/v3"

Write-Host "   Authenticated as account: $ApiSubdomain"

# ShareFile may rotate the refresh token on use; persist the new
# one so the next run doesn't fail with a stale token.
if ($TokenResponse.refresh_token -and $TokenResponse.refresh_token -ne $RefreshToken) {
    $EnvLines = Get-Content $EnvFile
    $NewEnvLines = $EnvLines | ForEach-Object {
        if ($_ -match '^\s*SHAREFILE_REFRESH_TOKEN\s*=') {
            "SHAREFILE_REFRESH_TOKEN=$($TokenResponse.refresh_token)"
        } else {
            $_
        }
    }
    Set-Content -Path $EnvFile -Value $NewEnvLines -Encoding UTF8
}

# ── Helper: call the ShareFile REST API ──────────────────────
function Invoke-SFApi {
    param([string]$Path)

    return Invoke-RestMethod -Uri "$ApiBase$Path" -Method Get -Headers @{
        "Authorization" = "Bearer $AccessToken"
    }
}

# ── Helper: pull the item array out of an OData feed response ─
function Get-ODataItems {
    param($Response)

    if ($null -eq $Response) { return @() }
    if ($Response.PSObject.Properties['value']) { return @($Response.value) }
    if ($Response.PSObject.Properties['Items']) { return @($Response.Items) }
    return @($Response)
}

# ── Helper: read the first matching property from an object ──
# Field names below are the most likely candidates for the ShareFile
# AccountUser entity; the public API docs don't publish a full field
# list, so run with -DumpRawSample against a real account and adjust
# these lists if a column comes back consistently blank.
$script:MissingFields = [System.Collections.Generic.HashSet[string]]::new()
function Get-Field {
    param($Obj, [string[]]$Names)
    foreach ($n in $Names) {
        if ($Obj.PSObject.Properties[$n]) {
            $v = $Obj.$n
            # $v -ne "" coerces "" to $false when $v is boolean, so an
            # explicit $false (e.g. IsDisabled: false) would otherwise
            # look empty and fall through as missing.
            if ($v -is [bool]) { return $v }
            if ($null -ne $v -and $v -ne "") { return $v }
        }
    }
    $script:MissingFields.Add($Names[0]) | Out-Null
    return $null
}

# ── Helper: escape a value for CSV ───────────────────────────
function ConvertTo-CsvField {
    param([string]$Value)
    $escaped = $Value -replace '"', '""'
    return "`"$escaped`""
}

# ── Helper: UTC -> account-local time ─────────────────────────
# ShareFile's API returns timestamps in UTC (PowerShell parses them
# as DateTime with Kind=Utc), but the account's own UI/exports show
# local time. There's no timezone field exposed via the API to read
# this from account preferences; confirmed empirically against a
# real UserList.xlsx export that this account is on US Eastern time
# (offsets matched EDT/EST exactly, DST-correct, across 34 users).
$script:EasternTz = [System.TimeZoneInfo]::FindSystemTimeZoneById("Eastern Standard Time")
# ShareFile represents "never logged in" as a 1900-01-01 sentinel rather
# than null (observed as 1900-01-01T05:00:00Z - account-local midnight,
# not UTC midnight) - treat any 1900 date as blank rather than a real one.
function ConvertTo-EasternString {
    param($Value)
    if ($null -eq $Value -or $Value -eq "") { return $Value }
    if ($Value -isnot [datetime]) { return $Value }
    if ($Value.Year -eq 1900) { return "" }
    $utc = if ($Value.Kind -eq [System.DateTimeKind]::Utc) { $Value } else { $Value.ToUniversalTime() }
    $local = [System.TimeZoneInfo]::ConvertTimeFromUtc($utc, $script:EasternTz)
    return $local.ToString("MM/dd/yyyy HH:mm:ss")
}

# ── Step 1: Fetch Employees and Clients ──────────────────────
# ShareFile splits account users into two feeds; the split itself
# tells us "UserType" (Employee vs Client) without needing to guess
# at a field for it.
$AllUsers = [System.Collections.Generic.List[object]]::new()

foreach ($Segment in @(
    @{ Path = "/Accounts/Employees"; Type = "Employee" },
    @{ Path = "/Accounts/Clients";   Type = "Client"   }
)) {
    Write-Host "Fetching $($Segment.Type)s..."
    $Page     = 0
    $PageSize = 100
    do {
        $Query    = "$($Segment.Path)?`$top=$PageSize&`$skip=$($Page * $PageSize)"
        $Response = Invoke-SFApi -Path $Query
        $Items    = Get-ODataItems -Response $Response
        $Count    = $Items.Count
        Write-Host "    Page $($Page + 1): $Count $($Segment.Type.ToLower())(s) returned"

        if ($DumpRawSample -and $Count -gt 0 -and $Page -eq 0) {
            $DebugDir = Join-Path $ReportsDir "debug"
            if (-not (Test-Path $DebugDir)) { New-Item -ItemType Directory -Path $DebugDir | Out-Null }
            $SampleFile = Join-Path $DebugDir "sharefile_raw_sample_$($Segment.Type).json"
            $Items[0] | ConvertTo-Json -Depth 10 | Set-Content -Path $SampleFile -Encoding UTF8
            Write-Host "    Raw sample written to: $SampleFile"
        }

        foreach ($Item in $Items) {
            $AllUsers.Add(@{ Data = $Item; Type = $Segment.Type })
        }

        $Page++
        Start-Sleep -Milliseconds 150
    } while ($Count -eq $PageSize)
}

Write-Host "Total users fetched: $($AllUsers.Count)"

# ── Step 1b: Fetch each user's secondary email ────────────────
# Not present on the Contact objects returned by Accounts/Employees
# / Accounts/Clients, and OData $select can't project it from that
# feed either (confirmed against a real account) - it only shows up
# on the full Users(id) object, as the non-primary entry in
# EmailAddresses. Requires one extra call per user.
Write-Host "Fetching secondary emails..."
$SecondaryEmailById = @{}
foreach ($Entry in $AllUsers) {
    $Id = $Entry.Data.Id
    try {
        $FullUser = Invoke-SFApi -Path "/Users($Id)?`$select=Id,EmailAddresses"
        $Secondary = @($FullUser.EmailAddresses | Where-Object { -not $_.IsPrimary } | Select-Object -ExpandProperty Email)
        if ($Secondary.Count -gt 0) {
            $SecondaryEmailById[$Id] = ($Secondary -join "; ")
        }
    } catch {
        Write-Warning "Could not fetch secondary email for user $Id`: $_"
    }
    Start-Sleep -Milliseconds 100
}

# ── Step 2: Write CSV (same columns as the manual export) ────
$Header = "Email,FirstName,LastName,Company,UserType,CreationDate,LastLoginDate,UserDisabled,SharedAddressBook,SecondaryEmail"
Set-Content -Path $OutputFile -Value $Header -Encoding UTF8

foreach ($Entry in $AllUsers) {
    $U = $Entry.Data

    $Email       = Get-Field $U @('Email', 'Username')
    $FirstName   = Get-Field $U @('FirstName')
    $LastName    = Get-Field $U @('LastName')
    $Company     = Get-Field $U @('Company')
    $Created     = ConvertTo-EasternString (Get-Field $U @('CreatedDate', 'CreationDate', 'DateCreated', 'Created'))
    $LastLogin   = ConvertTo-EasternString (Get-Field $U @('LastAnyLogin', 'LastLoginDate', 'LastLogin', 'LastAccess'))
    $IsDisabled  = Get-Field $U @('IsDisabled', 'Disabled', 'UserDisabled', 'IsInactive')

    # SharedAddressBook has no backing field anywhere in the API (checked
    # the full Users(id) property list and the Contacts feed) - it was
    # "Yes" for all 34 users in the manual reference export with zero
    # variance, so it's treated as a constant rather than a real per-user
    # value. If ShareFile ever exposes a real field for this, replace
    # the hardcode below.
    $SharedAB    = "Yes"
    $SecondEmail = if ($SecondaryEmailById.ContainsKey($U.Id)) { $SecondaryEmailById[$U.Id] } else { "" }

    $DisabledStr = if ($IsDisabled -eq $true -or "$IsDisabled" -ieq "true") { "Yes" } else { "No" }

    $Row = @(
        $Email,
        $FirstName,
        $LastName,
        $Company,
        $Entry.Type,
        $Created,
        $LastLogin,
        $DisabledStr,
        $SharedAB,
        $SecondEmail
    ) | ForEach-Object { ConvertTo-CsvField "$_" }

    Add-Content -Path $OutputFile -Value ($Row -join ",") -Encoding UTF8
}

$RowCount = $AllUsers.Count
Write-Host ""
Write-Host "Done! Report saved to: $OutputFile"
Write-Host "   Rows written: $RowCount"

if ($script:MissingFields.Count -gt 0) {
    Write-Warning "Some columns never matched a known field name on any user: $($script:MissingFields -join ', ')"
    Write-Warning "Re-run with -DumpRawSample and check the written sample JSON to find the correct field name(s), then update Get-Field candidate lists in this script."
}

# ── Step 3: Generate HTML executive report ────────────────────
Write-Host "Generating HTML executive report..."

$HtmlFile   = Join-Path $ReportsDir "sharefile_user_report_$Date.html"
$ReportData = Import-Csv -Path $OutputFile -Encoding UTF8

# Group by company
$CompanyGroups = $ReportData | Group-Object -Property "Company"

# Inactive = never logged in OR last login older than 30 days
$Cutoff        = (Get-Date).AddDays(-30)
$InactiveUsers = $ReportData | Where-Object {
    $la = $_."LastLoginDate"
    if (-not $la) { return $true }
    try { [datetime]::Parse($la) -lt $Cutoff } catch { $false }
} | Sort-Object -Property "Company", "LastLoginDate"

# External / Client users (ShareFile's guest-equivalent)
$ClientUsers = $ReportData | Where-Object { $_."UserType" -eq "Client" } | Sort-Object -Property "Company"

# New hires — created within the last 30 days
$NewHireCutoff = (Get-Date).AddDays(-30)
$NewHires      = $ReportData | Where-Object {
    $c = $_."CreationDate"
    if (-not $c) { return $false }
    try { [datetime]::Parse($c) -gt $NewHireCutoff } catch { $false }
} | Sort-Object -Property "CreationDate" -Descending

$NewHireCompanies = @($NewHires | Select-Object -ExpandProperty "Company" -Unique | Sort-Object)
$CoOptions = ($NewHireCompanies | ForEach-Object {
    $esc = $_ -replace '"', '&quot;'
    "<option value=`"$esc`">$_</option>"
}) -join "`n"

$NewHireTypeValues = @($NewHires | Select-Object -ExpandProperty "UserType" -Unique | Sort-Object)
$TypeOptions = ($NewHireTypeValues | ForEach-Object {
    $esc = $_ -replace '"', '&quot;'
    "<option value=`"$esc`">$_</option>"
}) -join "`n"

function Format-HtmlDate { param([string]$d) if ($d -and $d.Length -ge 10) { $d.Substring(0,10) } else { $d } }

# ── Company breakdown rows ──
$CoRows = foreach ($grp in ($CompanyGroups | Sort-Object Name)) {
    $u        = $grp.Group
    $total    = $u.Count
    $employees = @($u | Where-Object { $_."UserType" -eq "Employee" }).Count
    $clients   = @($u | Where-Object { $_."UserType" -eq "Client"   }).Count
    $disabled  = @($u | Where-Object { $_."UserDisabled" -eq "Yes" }).Count
    $active    = $total - $disabled
    $disClass  = if ($disabled -gt 0) { " class='warn'" } else { "" }
    $coName    = if ($grp.Name) { $grp.Name } else { "(none)" }
    @"
    <tr>
      <td>$coName</td>
      <td class="num">$total</td>
      <td class="num">$employees</td>
      <td class="num">$clients</td>
      <td class="num active">$active</td>
      <td class="num"$disClass>$disabled</td>
    </tr>
"@
}

# ── Inactive user rows ──
$InactiveRows = foreach ($u in $InactiveUsers) {
    $laCell = if (-not $u."LastLoginDate") {
        "<span class='never'>Never logged in</span>"
    } else {
        Format-HtmlDate $u."LastLoginDate"
    }
    @"
    <tr>
      <td>$($u.FirstName) $($u.LastName)</td>
      <td>$($u.Email)</td>
      <td>$($u.Company)</td>
      <td>$($u.UserType)</td>
      <td>$laCell</td>
    </tr>
"@
}

# ── External / Client user rows ──
$ClientRows = foreach ($u in $ClientUsers) {
    $stClass = if ($u.UserDisabled -eq "Yes") { "warn" } else { "active" }
    $stText  = if ($u.UserDisabled -eq "Yes") { "Disabled" } else { "Active" }
    @"
    <tr>
      <td>$($u.FirstName) $($u.LastName)</td>
      <td>$($u.Email)</td>
      <td>$($u.Company)</td>
      <td class="$stClass">$stText</td>
      <td>$(Format-HtmlDate $u.CreationDate)</td>
    </tr>
"@
}

# ── New hire rows ──
$NewHireRows = foreach ($u in $NewHires) {
    $coAttr   = ($u.Company  -replace '"', '&quot;')
    $typeAttr = ($u.UserType -replace '"', '&quot;')
    $stAttr   = ($u.UserDisabled -replace '"', '&quot;')
    @"
    <tr class="new-hire" data-company="$coAttr" data-usertype="$typeAttr" data-disabled="$stAttr">
      <td>$($u.FirstName) $($u.LastName)</td>
      <td>$($u.Email)</td>
      <td>$($u.Company)</td>
      <td>$($u.UserType)</td>
      <td>$(Format-HtmlDate $u.CreationDate)</td>
      <td>$($u.UserDisabled)</td>
    </tr>
"@
}

# ── Helper: render a table section or an empty-state message ──
function New-HtmlSection {
    param([string]$Title, [string]$Head, [string[]]$Rows, [string]$EmptyMsg)
    $rowCount = ($Rows | Measure-Object).Count
    $inner = if ($rowCount -gt 0) {
        "<table><thead><tr>$Head</tr></thead><tbody>$($Rows -join '')</tbody></table>"
    } else {
        "<p class='empty'>$EmptyMsg</p>"
    }
    "<section><h2>$Title</h2>$inner</section>"
}

$CoSection = New-HtmlSection `
    -Title    "Company Breakdown" `
    -Head     "<th class='sortable' onclick='sortTable(this)'>Company<span class='sort-icon'>&#8645;</span></th><th class='sortable' onclick='sortTable(this)' data-type='num'>Total<span class='sort-icon'>&#8645;</span></th><th class='sortable' onclick='sortTable(this)' data-type='num'>Employees<span class='sort-icon'>&#8645;</span></th><th class='sortable' onclick='sortTable(this)' data-type='num'>Clients<span class='sort-icon'>&#8645;</span></th><th class='sortable' onclick='sortTable(this)' data-type='num'>Active<span class='sort-icon'>&#8645;</span></th><th class='sortable' onclick='sortTable(this)' data-type='num'>Disabled<span class='sort-icon'>&#8645;</span></th>" `
    -Rows     $CoRows `
    -EmptyMsg "No company data."

$InactiveSection = New-HtmlSection `
    -Title    "Inactive Users (30+ days without login)" `
    -Head     "<th class='sortable' onclick='sortTable(this)'>Name<span class='sort-icon'>&#8645;</span></th><th class='sortable' onclick='sortTable(this)'>Email<span class='sort-icon'>&#8645;</span></th><th class='sortable' onclick='sortTable(this)'>Company<span class='sort-icon'>&#8645;</span></th><th class='sortable' onclick='sortTable(this)'>Type<span class='sort-icon'>&#8645;</span></th><th class='sortable' onclick='sortTable(this)'>Last Login<span class='sort-icon'>&#8645;</span></th>" `
    -Rows     $InactiveRows `
    -EmptyMsg "No inactive users found."

$ClientSection = New-HtmlSection `
    -Title    "External Users (Clients)" `
    -Head     "<th class='sortable' onclick='sortTable(this)'>Name<span class='sort-icon'>&#8645;</span></th><th class='sortable' onclick='sortTable(this)'>Email<span class='sort-icon'>&#8645;</span></th><th class='sortable' onclick='sortTable(this)'>Company<span class='sort-icon'>&#8645;</span></th><th class='sortable' onclick='sortTable(this)'>Status<span class='sort-icon'>&#8645;</span></th><th class='sortable' onclick='sortTable(this)'>Created<span class='sort-icon'>&#8645;</span></th>" `
    -Rows     $ClientRows `
    -EmptyMsg "No client users found."

$NewHireInner = if ($NewHireRows.Count -gt 0) {
    @"
<div class="filter-bar">
  <label for="co-filter">Company</label>
  <select id="co-filter" onchange="filterNewHires()">
    <option value="">All Companies</option>
    $CoOptions
  </select>
  <label for="ty-filter">User Type</label>
  <select id="ty-filter" onchange="filterNewHires()">
    <option value="">All</option>
    $TypeOptions
  </select>
  <label for="ds-filter">Disabled</label>
  <select id="ds-filter" onchange="filterNewHires()">
    <option value="">All</option>
    <option value="Yes">Yes</option>
    <option value="No">No</option>
  </select>
  <span class="filter-count" id="hire-count">$($NewHires.Count) hire(s)</span>
</div>
<table id="new-hire-table">
  <thead><tr><th class="sortable" onclick="sortTable(this)">Name<span class="sort-icon">&#8645;</span></th><th class="sortable" onclick="sortTable(this)">Email<span class="sort-icon">&#8645;</span></th><th class="sortable" onclick="sortTable(this)">Company<span class="sort-icon">&#8645;</span></th><th class="sortable" onclick="sortTable(this)">Type<span class="sort-icon">&#8645;</span></th><th class="sortable" onclick="sortTable(this)">Created<span class="sort-icon">&#8645;</span></th><th class="sortable" onclick="sortTable(this)">Disabled<span class="sort-icon">&#8645;</span></th></tr></thead>
  <tbody>$($NewHireRows -join '')</tbody>
</table>
"@
} else {
    "<p class='empty'>No new users in the last 30 days.</p>"
}
$NewHireSection = "<section><h2>New Users &mdash; Last 30 Days</h2>$NewHireInner</section>"

$Html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>ShareFile User Report - $Date</title>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: 'Segoe UI', Arial, sans-serif; background: #f4f5f7; color: #172b4d; }
  header { background: #0052cc; color: #fff; padding: 24px 40px; }
  header h1 { font-size: 22px; font-weight: 600; }
  header p  { font-size: 13px; opacity: 0.75; margin-top: 5px; }
  main { padding: 32px 40px; max-width: 1200px; }
  section { background: #fff; border-radius: 6px; box-shadow: 0 1px 3px rgba(0,0,0,.12); margin-bottom: 28px; overflow: hidden; }
  section h2 { font-size: 12px; font-weight: 700; text-transform: uppercase; letter-spacing: .06em;
               color: #5e6c84; padding: 14px 20px; border-bottom: 1px solid #ebecf0; }
  table { width: 100%; border-collapse: collapse; font-size: 13px; }
  th { background: #f4f5f7; color: #5e6c84; font-weight: 700; text-transform: uppercase;
       font-size: 11px; letter-spacing: .05em; padding: 10px 16px; text-align: left; }
  td { padding: 9px 16px; border-top: 1px solid #ebecf0; vertical-align: middle; }
  tr:hover td { background: #fafbfc; }
  td.num { text-align: center; }
  .active { color: #006644; font-weight: 600; }
  .warn   { color: #bf2600; font-weight: 600; }
  .never  { color: #bf2600; }
  a { color: #0052cc; text-decoration: none; }
  a:hover { text-decoration: underline; }
  .empty { color: #7a869a; font-style: italic; padding: 16px 20px; font-size: 13px; }
  .new-hire td { background: #e8f0fe; }
  tr.new-hire:hover td { background: #d2e3fc; }
  .filter-bar { padding: 12px 20px; border-bottom: 1px solid #ebecf0; display: flex; align-items: center; gap: 12px; }
  .filter-bar label { font-size: 12px; font-weight: 700; color: #5e6c84; text-transform: uppercase; letter-spacing: .05em; }
  .filter-bar select { font-size: 13px; padding: 5px 10px; border: 1px solid #dfe1e6; border-radius: 4px; color: #172b4d; cursor: pointer; }
  .filter-count { font-size: 12px; color: #5e6c84; }
  th.sortable { cursor: pointer; user-select: none; white-space: nowrap; }
  th.sortable:hover { background: #e8eaed; color: #172b4d; }
  th.sortable[data-sort="asc"],th.sortable[data-sort="desc"] { color: #0052cc; }
  .sort-icon { display: inline-block; margin-left: 5px; font-size: 10px; color: #b3bac5; vertical-align: middle; transition: color .15s; }
  th.sortable:hover .sort-icon { color: #5e6c84; }
  th.sortable[data-sort="asc"] .sort-icon,th.sortable[data-sort="desc"] .sort-icon { color: #0052cc; }
</style>
</head>
<body>
<header>
  <h1>ShareFile User Report</h1>
  <p>Generated $Date &nbsp;&bull;&nbsp; $RowCount users across $($CompanyGroups.Count) companies &nbsp;&bull;&nbsp; Source: $OutputFile</p>
</header>
<main>
$NewHireSection
$CoSection
$InactiveSection
$ClientSection
</main>
<script>
function filterNewHires() {
  var co = document.getElementById('co-filter').value;
  var ty = document.getElementById('ty-filter').value;
  var ds = document.getElementById('ds-filter').value;
  var rows = document.querySelectorAll('#new-hire-table tbody tr');
  var visible = 0;
  rows.forEach(function(row) {
    var show = (!co || row.dataset.company === co) &&
               (!ty || row.dataset.usertype === ty) &&
               (!ds || row.dataset.disabled === ds);
    row.style.display = show ? '' : 'none';
    if (show) visible++;
  });
  document.getElementById('hire-count').textContent = visible + ' hire(s)';
}
function sortTable(th) {
  var table   = th.closest('table');
  var tbody   = table.querySelector('tbody');
  var colIdx  = th.cellIndex;
  var isNum   = th.dataset.type === 'num';
  var current = th.dataset.sort || 'none';
  var next    = current === 'none' ? 'asc' : current === 'asc' ? 'desc' : 'none';
  table.querySelectorAll('th.sortable').forEach(function(h) {
    delete h.dataset.sort;
    h.querySelector('.sort-icon').textContent = '⇅';
  });
  if (next === 'none') return;
  th.dataset.sort = next;
  th.querySelector('.sort-icon').textContent = next === 'asc' ? '↑' : '↓';
  var rows = Array.from(tbody.querySelectorAll('tr'));
  rows.sort(function(a, b) {
    var aVal = a.cells[colIdx] ? a.cells[colIdx].textContent.trim() : '';
    var bVal = b.cells[colIdx] ? b.cells[colIdx].textContent.trim() : '';
    var cmp;
    if (isNum) {
      var aNum = parseFloat(aVal.replace(/[^0-9.\-]/g, ''));
      var bNum = parseFloat(bVal.replace(/[^0-9.\-]/g, ''));
      cmp = (isNaN(aNum) ? -Infinity : aNum) - (isNaN(bNum) ? -Infinity : bNum);
    } else {
      cmp = aVal.toLowerCase().localeCompare(bVal.toLowerCase());
    }
    return next === 'asc' ? cmp : -cmp;
  });
  var frag = document.createDocumentFragment();
  rows.forEach(function(r) { frag.appendChild(r); });
  tbody.appendChild(frag);
}
</script>
</body>
</html>
"@

Set-Content -Path $HtmlFile -Value $Html -Encoding UTF8
Write-Host "   HTML report saved to: $HtmlFile"

# ── Step 4: Generate Excel report ─────────────────────────────
Write-Host "Generating Excel report..."
$ExcelFile = Join-Path $ReportsDir "sharefile_user_report_$Date.xlsx"

try {
    if (Test-Path $ExcelFile) { Remove-Item $ExcelFile -Force -ErrorAction Stop }

    $ExcelPackage = $ReportData | Export-Excel -Path $ExcelFile -WorksheetName "Users" `
        -AutoFilter -FreezeTopRow -BoldTopRow -AutoSize -PassThru

    $Sheet   = $ExcelPackage.Workbook.Worksheets["Users"]
    $LastCol = $Sheet.Dimension.End.Column
    $LastRow = $Sheet.Dimension.End.Row

    # Locate the "CreationDate" column
    $CreatedIdx = 1
    for ($c = 1; $c -le $LastCol; $c++) {
        if ($Sheet.Cells[1, $c].Value -eq "CreationDate") { $CreatedIdx = $c; break }
    }

    # Blue fill (#e8f0fe) for any row whose CreationDate is within the last 30 days
    $BlueColor = [System.Drawing.Color]::FromArgb(232, 240, 254)
    $XlCutoff  = (Get-Date).AddDays(-30)
    for ($r = 2; $r -le $LastRow; $r++) {
        $cVal = $Sheet.Cells[$r, $CreatedIdx].Value
        if ($cVal) {
            try {
                if ([datetime]::Parse($cVal) -gt $XlCutoff) {
                    $range = $Sheet.Cells[$r, 1, $r, $LastCol]
                    $range.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                    $range.Style.Fill.BackgroundColor.SetColor($BlueColor)
                }
            } catch {}
        }
    }

    Close-ExcelPackage $ExcelPackage
    Write-Host "   Excel report saved to: $ExcelFile"
} catch {
    Write-Warning "Excel report skipped - close '$ExcelFile' in Excel and re-run to generate it."
}
