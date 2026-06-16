# ============================================================
# monday_user_report.ps1
# Generates a Workspace -> User mapping report for Monday.com
# Replicates the Coral User Report (manual XLSX) via API
#
# Usage:
#   Set-Content .env "MONDAY_API_TOKEN=your_token_here"
#   .\monday_user_report.ps1
#
# Output: coral_user_report_YYYY-MM-DD.csv
# ============================================================

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ── Load .env ─────────────────────────────────────────────────
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$EnvFile   = Join-Path $ScriptDir ".env"

if (Test-Path $EnvFile) {
    Get-Content $EnvFile | ForEach-Object {
        if ($_ -match '^\s*([^#][^=]+)=(.*)$') {
            [System.Environment]::SetEnvironmentVariable($Matches[1].Trim(), $Matches[2].Trim(), "Process")
        }
    }
}

# ── Config ────────────────────────────────────────────────────
$ApiUrl     = "https://api.monday.com/v2"
$Token      = $env:MONDAY_API_TOKEN
if (-not $Token) { throw "Set MONDAY_API_TOKEN env var first" }

$Date       = Get-Date -Format "yyyy-MM-dd"
$ReportsDir = Join-Path $ScriptDir "Reports"
if (-not (Test-Path $ReportsDir)) { New-Item -ItemType Directory -Path $ReportsDir | Out-Null }
$OutputFile = Join-Path $ReportsDir "coral_user_report_$Date.csv"

# ── Helper: run a GraphQL query ───────────────────────────────
function Invoke-GQL {
    param([string]$Query)

    $Body = @{ query = $Query } | ConvertTo-Json -Compress

    $Response = Invoke-RestMethod `
        -Uri     $ApiUrl `
        -Method  Post `
        -Headers @{
            "Content-Type" = "application/json"
            "Authorization" = $Token
            "API-Version"  = "2024-07"
        } `
        -Body $Body

    return $Response
}

# ── Helper: escape a value for CSV ───────────────────────────
function ConvertTo-CsvField {
    param([string]$Value)
    $escaped = $Value -replace '"', '""'
    return "`"$escaped`""
}

# ── Step 1: Get all workspaces ────────────────────────────────
Write-Host "Fetching workspaces..."

$WorkspacesQuery = @"
query {
  workspaces(limit: 100) {
    id
    name
    kind
  }
}
"@

$WorkspacesResult = Invoke-GQL -Query $WorkspacesQuery

if ($WorkspacesResult.PSObject.Properties['errors']) {
    Write-Error "API Error fetching workspaces: $($WorkspacesResult.errors | ConvertTo-Json)"
    exit 1
}

# ── Step 2: Write CSV header ──────────────────────────────────
$Header = "Workspace,Name,Email,User Role,Status,Teams,Joined,Last Active,Invited By,2FA,Workspace URL"
Set-Content -Path $OutputFile -Value $Header -Encoding UTF8

# ── Step 3: For each workspace, fetch members ─────────────────
foreach ($Workspace in $WorkspacesResult.data.workspaces) {
    $WsId   = $Workspace.id
    $WsName = $Workspace.name
    $WsKind = $Workspace.kind   # "open" or "closed"
    $WsUrl  = "https://coral.monday.com/workspaces/$WsId"

    Write-Host "  -> Processing workspace: $WsName (id: $WsId, kind: $WsKind)"

    $AllMembers = [System.Collections.Generic.List[object]]::new()
    $Page       = 1
    $PageSize   = 100

    if ($WsKind -eq "open") {
        # Open workspaces include all account users implicitly — users_subscribers
        # only returns explicit subscribers, so we use the top-level users query instead.
        do {
            $MembersQuery = @"
query {
  users(limit: $PageSize, page: $Page, kind: all) {
    id
    name
    email
    is_admin
    is_guest
    enabled
    created_at
    last_activity
    is_view_only
    teams {
      name
    }
  }
}
"@
            $MembersResult = Invoke-GQL -Query $MembersQuery
            if ($MembersResult.PSObject.Properties['errors']) {
                Write-Warning "API Error fetching users (page $Page)`: $($MembersResult.errors | ConvertTo-Json)"
                break
            }

            $PageMembers = $MembersResult.data.users
            $PageCount   = if ($PageMembers) { @($PageMembers).Count } else { 0 }
            Write-Host "    Page $Page`: $PageCount user(s) returned"

            if ($PageCount -gt 0) {
                @($PageMembers) | ForEach-Object { $AllMembers.Add($_) }
            }

            $Page++
            Start-Sleep -Milliseconds 150
        } while ($PageCount -eq $PageSize)
    } else {
        # Closed workspaces have explicit membership — users_subscribers is correct.
        do {
            $MembersQuery = @"
query {
  workspaces(ids: [$WsId]) {
    members: users_subscribers(limit: $PageSize, page: $Page) {
      id
      name
      email
      is_admin
      is_guest
      enabled
      created_at
      last_activity
      is_view_only
      teams {
        name
      }
    }
  }
}
"@
            $MembersResult = Invoke-GQL -Query $MembersQuery
            if ($MembersResult.PSObject.Properties['errors']) {
                Write-Warning "API Error fetching members for workspace $WsName (page $Page)`: $($MembersResult.errors | ConvertTo-Json)"
                break
            }

            $PageMembers = $MembersResult.data.workspaces[0].members
            $PageCount   = if ($PageMembers) { @($PageMembers).Count } else { 0 }
            Write-Host "    Page $Page`: $PageCount member(s) returned"

            if ($PageCount -gt 0) {
                @($PageMembers) | ForEach-Object { $AllMembers.Add($_) }
            }

            $Page++
            Start-Sleep -Milliseconds 150
        } while ($PageCount -eq $PageSize)
    }

    $Members = $AllMembers

    if ($Members.Count -eq 0) {
        Write-Host "    (no members found)"
        continue
    }

    Write-Host "    Total members in '$WsName'`: $($Members.Count)"

    foreach ($Member in $Members) {
        # Determine role
        $Role = if     ($Member.is_admin)     { "Admin"  }
                elseif ($Member.is_guest)     { "Guest"  }
                elseif ($Member.is_view_only) { "Viewer" }
                else                          { "Member" }

        # Invited By
        $InvitedBy = if ($Member.invited_by) { $Member.invited_by.name } else { "N/A" }

        # Products (not available via workspace users_subscribers)
        #$Products = "N/A"

        # Status
        $Status = if ($Member.enabled) { "Active" } else { "Inactive" }

        # Teams
        $TeamNames = @(if ($Member.teams) { $Member.teams | ForEach-Object { $_.name } })
        $Teams     = if ($TeamNames.Count -gt 0) { $TeamNames -join "; " } else { "No Teams" }

        # Dates
        $Joined     = if ($Member.created_at)     { $Member.created_at }     else { "" }
        $LastActive = if ($Member.last_activity)  { $Member.last_activity }  else { "Never logged in" }

        # Build CSV row
        $Row = @(
            $WsName,
            $Member.name,
            $Member.email,
            $Role,
            $Status,
            $Teams,
            $Joined,
            $LastActive,
            $InvitedBy,
            "Disabled",
            $WsUrl
        ) | ForEach-Object { ConvertTo-CsvField $_ }

        Add-Content -Path $OutputFile -Value ($Row -join ",") -Encoding UTF8
    }

    # Rate limiting
    Start-Sleep -Milliseconds 300
}

# ── Step 4: Remove "Deleted member" rows ─────────────────────
Write-Host "Removing deleted-member records..."

$AllLines    = Get-Content $OutputFile -Encoding UTF8
$HeaderLine  = $AllLines[0]
$DataLines   = $AllLines | Select-Object -Skip 1

$Cleaned     = $DataLines | Where-Object { $_ -notmatch '^"Deleted member",' -and $_ -notmatch ',"Deleted member",' }
$RemovedCount = $DataLines.Count - $Cleaned.Count

@($HeaderLine) + $Cleaned | Set-Content -Path $OutputFile -Encoding UTF8

$RowCount = $Cleaned.Count
Write-Host ""
Write-Host "Done! Report saved to: $OutputFile"
Write-Host "   Rows written: $RowCount"
Write-Host "   Deleted-member rows removed: $RemovedCount"

# ── Step 5: Generate HTML executive report ────────────────────
Write-Host "Generating HTML executive report..."

$HtmlFile   = Join-Path $ReportsDir "coral_user_report_$Date.html"
$ReportData = Import-Csv -Path $OutputFile -Encoding UTF8

# Group by workspace
$WorkspaceGroups = $ReportData | Group-Object -Property "Workspace"

# Inactive = never logged in OR last activity older than 30 days
$Cutoff        = (Get-Date).AddDays(-30)
$InactiveUsers = $ReportData | Where-Object {
    $la = $_."Last Active"
    if ($la -eq "Never logged in") { return $true }
    try { [datetime]::Parse($la) -lt $Cutoff } catch { $false }
} | Sort-Object -Property "Workspace", "Last Active"

# Guest / external users
$GuestUsers = $ReportData | Where-Object { $_."User Role" -eq "Guest" } | Sort-Object -Property "Workspace"

# New hires — joined within the last 30 days
$NewHireCutoff     = (Get-Date).AddDays(-30)
$NewHires          = $ReportData | Where-Object {
    $j = $_."Joined"
    if (-not $j) { return $false }
    try { [datetime]::Parse($j) -gt $NewHireCutoff } catch { $false }
} | Sort-Object -Property "Joined" -Descending

$NewHireWorkspaces = @($NewHires | Select-Object -ExpandProperty "Workspace" -Unique | Sort-Object)
$WsOptions = ($NewHireWorkspaces | ForEach-Object {
    $esc = $_ -replace '"', '&quot;'
    "<option value=`"$esc`">$_</option>"
}) -join "`n"

function Format-HtmlDate { param([string]$d) if ($d.Length -ge 10) { $d.Substring(0,10) } else { $d } }

# ── Workspace breakdown rows ──
$WsRows = foreach ($grp in ($WorkspaceGroups | Sort-Object Name)) {
    $u        = $grp.Group
    $wsUrl    = $u[0]."Workspace URL"
    $total    = $u.Count
    $admins   = @($u | Where-Object { $_."User Role" -eq "Admin"   }).Count
    $members  = @($u | Where-Object { $_."User Role" -eq "Member"  }).Count
    $viewers  = @($u | Where-Object { $_."User Role" -eq "Viewer"  }).Count
    $guests   = @($u | Where-Object { $_."User Role" -eq "Guest"   }).Count
    $active   = @($u | Where-Object { $_."Status"    -eq "Active"  }).Count
    $inactive = @($u | Where-Object { $_."Status"    -eq "Inactive" }).Count
    $inClass  = if ($inactive -gt 0) { " class='warn'" } else { "" }
    @"
    <tr>
      <td><a href="$wsUrl" target="_blank">$($grp.Name)</a></td>
      <td class="num">$total</td>
      <td class="num">$admins</td>
      <td class="num">$members</td>
      <td class="num">$viewers</td>
      <td class="num">$guests</td>
      <td class="num active">$active</td>
      <td class="num"$inClass>$inactive</td>
    </tr>
"@
}

# ── Inactive user rows ──
$InactiveRows = foreach ($u in $InactiveUsers) {
    $laCell = if ($u."Last Active" -eq "Never logged in") {
        "<span class='never'>Never logged in</span>"
    } else {
        Format-HtmlDate $u."Last Active"
    }
    @"
    <tr>
      <td>$($u.Name)</td>
      <td>$($u.Email)</td>
      <td>$($u.Workspace)</td>
      <td>$($u."User Role")</td>
      <td>$laCell</td>
    </tr>
"@
}

# ── Guest user rows ──
$GuestRows = foreach ($u in $GuestUsers) {
    $stClass = if ($u.Status -eq "Active") { "active" } else { "warn" }
    @"
    <tr>
      <td>$($u.Name)</td>
      <td>$($u.Email)</td>
      <td>$($u.Workspace)</td>
      <td class="$stClass">$($u.Status)</td>
      <td>$(Format-HtmlDate $u.Joined)</td>
    </tr>
"@
}

# ── New hire rows ──
$NewHireRows = foreach ($u in $NewHires) {
    $wsAttr = ($u.Workspace -replace '"', '&quot;')
    $stAttr = ($u.Status    -replace '"', '&quot;')
    @"
    <tr class="new-hire" data-workspace="$wsAttr" data-status="$stAttr">
      <td>$($u.Name)</td>
      <td>$($u.Email)</td>
      <td>$($u.Workspace)</td>
      <td>$($u."User Role")</td>
      <td>$($u.Status)</td>
      <td>$(Format-HtmlDate $u.Joined)</td>
      <td>$($u."Invited By")</td>
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

$WsSection = New-HtmlSection `
    -Title    "Workspace Breakdown" `
    -Head     "<th>Workspace</th><th>Total</th><th>Admins</th><th>Members</th><th>Viewers</th><th>Guests</th><th>Active</th><th>Inactive</th>" `
    -Rows     $WsRows `
    -EmptyMsg "No workspace data."

$InactiveSection = New-HtmlSection `
    -Title    "Inactive Users (30+ days without login)" `
    -Head     "<th>Name</th><th>Email</th><th>Workspace</th><th>Role</th><th>Last Active</th>" `
    -Rows     $InactiveRows `
    -EmptyMsg "No inactive users found."

$GuestSection = New-HtmlSection `
    -Title    "Guest / External Users" `
    -Head     "<th>Name</th><th>Email</th><th>Workspace</th><th>Status</th><th>Joined</th>" `
    -Rows     $GuestRows `
    -EmptyMsg "No guest users found."

$NewHireInner = if ($NewHireRows.Count -gt 0) {
    @"
<div class="filter-bar">
  <label for="ws-filter">Workspace</label>
  <select id="ws-filter" onchange="filterNewHires()">
    <option value="">All Workspaces</option>
    $WsOptions
  </select>
  <label for="st-filter">Status</label>
  <select id="st-filter" onchange="filterNewHires()">
    <option value="">All</option>
    <option value="Active">Active</option>
    <option value="Inactive">Inactive</option>
  </select>
  <span class="filter-count" id="hire-count">$($NewHires.Count) hire(s)</span>
</div>
<table id="new-hire-table">
  <thead><tr><th>Name</th><th>Email</th><th>Workspace</th><th>Role</th><th>Status</th><th>Joined</th><th>Invited By</th></tr></thead>
  <tbody>$($NewHireRows -join '')</tbody>
</table>
"@
} else {
    "<p class='empty'>No new hires in the last 30 days.</p>"
}
$NewHireSection = "<section><h2>New Hires &mdash; Last 30 Days</h2>$NewHireInner</section>"

$Html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>Coral Monday.com User Report - $Date</title>
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
</style>
</head>
<body>
<header>
  <h1>Coral &mdash; Monday.com User Report</h1>
  <p>Generated $Date &nbsp;&bull;&nbsp; $RowCount users across $($WorkspaceGroups.Count) workspaces &nbsp;&bull;&nbsp; Source: $OutputFile</p>
</header>
<main>
$NewHireSection
$WsSection
$InactiveSection
$GuestSection
</main>
<script>
function filterNewHires() {
  var ws = document.getElementById('ws-filter').value;
  var st = document.getElementById('st-filter').value;
  var rows = document.querySelectorAll('#new-hire-table tbody tr');
  var visible = 0;
  rows.forEach(function(row) {
    var show = (!ws || row.dataset.workspace === ws) &&
               (!st || row.dataset.status === st);
    row.style.display = show ? '' : 'none';
    if (show) visible++;
  });
  document.getElementById('hire-count').textContent = visible + ' hire(s)';
}
</script>
</body>
</html>
"@

Set-Content -Path $HtmlFile -Value $Html -Encoding UTF8
Write-Host "   HTML report saved to: $HtmlFile"

# ── Step 6: Generate Excel report ─────────────────────────────
Write-Host "Generating Excel report..."
$ExcelFile = Join-Path $ReportsDir "coral_user_report_$Date.xlsx"

try {
    if (Test-Path $ExcelFile) { Remove-Item $ExcelFile -Force -ErrorAction Stop }

    $ExcelPackage = $ReportData | Export-Excel -Path $ExcelFile -WorksheetName "Users" `
        -AutoFilter -FreezeTopRow -BoldTopRow -AutoSize -PassThru

    $Sheet   = $ExcelPackage.Workbook.Worksheets["Users"]
    $LastCol = $Sheet.Dimension.End.Column
    $LastRow = $Sheet.Dimension.End.Row

    # Locate the "Joined" column
    $JoinedIdx = 1
    for ($c = 1; $c -le $LastCol; $c++) {
        if ($Sheet.Cells[1, $c].Value -eq "Joined") { $JoinedIdx = $c; break }
    }

    # Blue fill (#e8f0fe) for any row whose Joined date is within the last 30 days
    $BlueColor = [System.Drawing.Color]::FromArgb(232, 240, 254)
    $XlCutoff  = (Get-Date).AddDays(-30)
    for ($r = 2; $r -le $LastRow; $r++) {
        $jVal = $Sheet.Cells[$r, $JoinedIdx].Value
        if ($jVal) {
            try {
                if ([datetime]::Parse($jVal) -gt $XlCutoff) {
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
