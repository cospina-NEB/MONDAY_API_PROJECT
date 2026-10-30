# ============================================================
# monday_rcm_export.ps1
# Trace-level, READ-ONLY export of the CORAL -- RCM workspace
# for Coral's data warehouse / Power BI.
#
# Only GraphQL `query` operations are issued — nothing in monday
# is created or changed.
#
# Usage:
#   .\scripts\monday_rcm_export.ps1 [-InitialLookbackDays 30]
#
# Output (Reports/monday_rcm/, timestamped yyyy-MM-ddTHHmm):
#   rcm_items_<ts>.csv          full snapshot, one row per item
#   rcm_column_values_<ts>.csv  full snapshot, one row per item x column
#   rcm_activity_<ts>.csv       incremental, activity-log events since last run
#   rcm_users_<ts>.csv          full snapshot, users with job titles
#   rcm_state.json              activity-log watermark for the next run
# ============================================================

param(
    # Activity-log window for the very first run (no rcm_state.json yet)
    [int]$InitialLookbackDays = 30
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
$ApiUrl  = "https://api.monday.com/v2"
$Token   = $env:MONDAY_API_TOKEN
if (-not $Token) { throw "Set MONDAY_API_TOKEN env var first" }

$WorkspaceId   = $env:MONDAY_RCM_WORKSPACE_ID
$WorkspaceName = if ($env:MONDAY_RCM_WORKSPACE_NAME) { $env:MONDAY_RCM_WORKSPACE_NAME } else { "CORAL -- RCM" }

$RunStart   = [DateTime]::UtcNow
$Stamp      = $RunStart.ToString("yyyy-MM-ddTHHmm")
$ReportsDir = Join-Path $ProjectRoot "Reports\monday_rcm"
if (-not (Test-Path $ReportsDir)) { New-Item -ItemType Directory -Path $ReportsDir | Out-Null }
$StateFile  = Join-Path $ReportsDir "rcm_state.json"

$ItemsFile    = Join-Path $ReportsDir "rcm_items_$Stamp.csv"
$ValuesFile   = Join-Path $ReportsDir "rcm_column_values_$Stamp.csv"
$ActivityFile = Join-Path $ReportsDir "rcm_activity_$Stamp.csv"
$UsersFile    = Join-Path $ReportsDir "rcm_users_$Stamp.csv"

# ── Helper: run a GraphQL query ───────────────────────────────
# Unlike the user report, any API error is fatal here: a partial
# trace export is worse than a failed run that retries next time.
# Rate-limit / complexity errors are retried with backoff.
function Invoke-GQL {
    param([string]$Query)

    $Body = @{ query = $Query } | ConvertTo-Json -Compress

    for ($Attempt = 1; $Attempt -le 4; $Attempt++) {
        try {
            $Response = Invoke-RestMethod `
                -Uri     $ApiUrl `
                -Method  Post `
                -Headers @{
                    "Content-Type" = "application/json"
                    "Authorization" = $Token
                    "API-Version"  = "2026-07"
                } `
                -Body $Body
        } catch {
            if ($Attempt -lt 4 -and "$_" -match '429|rate limit|complexity') {
                Write-Warning "Rate limited (attempt $Attempt), waiting 30s..."
                Start-Sleep -Seconds 30
                continue
            }
            throw
        }

        if ($Response.PSObject.Properties['errors'] -and $Response.errors) {
            $ErrJson = $Response.errors | ConvertTo-Json -Depth 10 -Compress
            if ($Attempt -lt 4 -and $ErrJson -match 'Complexity|rate limit|RATE_LIMIT') {
                Write-Warning "Complexity budget exhausted (attempt $Attempt), waiting 30s..."
                Start-Sleep -Seconds 30
                continue
            }
            throw "API Error: $ErrJson"
        }
        return $Response
    }
}

# ── Helper: escape a value for CSV ───────────────────────────
function ConvertTo-CsvField {
    param([string]$Value)
    $escaped = $Value -replace '"', '""'
    return "`"$escaped`""
}

# ── Helper: escape a string for a GraphQL string literal ─────
function ConvertTo-GqlString {
    param([string]$Value)
    return '"' + ($Value -replace '\\', '\\' -replace '"', '\"') + '"'
}

# ── Helper: null-safe property read (StrictMode-friendly) ────
function Get-Prop {
    param($Object, [string]$Name)
    if ($null -eq $Object) { return $null }
    if ($Object -is [System.Collections.IDictionary]) { return $Object[$Name] }
    $p = $Object.PSObject.Properties[$Name]
    if ($p) { return $p.Value } else { return $null }
}

# ── Helper: monday activity-log timestamp -> ISO UTC ─────────
# activity_logs.created_at is a 17-digit string in 100ns units
# since the Unix epoch (e.g. "17306281234567890").
function ConvertFrom-MondayLogTimestamp {
    param([string]$Value)
    if ($Value -notmatch '^\d{17}$') { return $Value }
    $Ms = [long]$Value / 10000
    return [DateTimeOffset]::FromUnixTimeMilliseconds($Ms).UtcDateTime.ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
}

# ── Helper: activity-log value object -> readable text ───────
# Values in `data` vary by column type; pick the human-readable
# part where the shape is known, else fall back to compact JSON.
function Get-LogValueText {
    param($Value)
    if ($null -eq $Value) { return "" }
    if ($Value -is [string] -or $Value -is [ValueType]) { return "$Value" }

    $Label = Get-Prop $Value 'label'
    if ($Label) {
        $LabelText = Get-Prop $Label 'text'
        if ($null -ne $LabelText) { return "$LabelText" }
    }
    foreach ($Key in 'text', 'value', 'name') {
        $v = Get-Prop $Value $Key
        if ($null -ne $v -and ($v -is [string] -or $v -is [ValueType])) { return "$v" }
    }
    $Date = Get-Prop $Value 'date'
    if ($Date) {
        $Time = Get-Prop $Value 'time'
        if ($Time) { return "$Date $Time" } else { return "$Date" }
    }
    $People = Get-Prop $Value 'personsAndTeams'
    if ($People) {
        return (@($People) | ForEach-Object { "$(Get-Prop $_ 'kind'):$(Get-Prop $_ 'id')" }) -join "; "
    }
    return ($Value | ConvertTo-Json -Depth 10 -Compress)
}

# ── Helper: parse an activity-log `data` JSON string ─────────
function ConvertFrom-ActivityData {
    param([string]$Data)
    $Result = [ordered]@{ item_id = ""; column_id = ""; column_title = ""; previous_value_text = ""; value_text = "" }
    if ([string]::IsNullOrWhiteSpace($Data)) { return $Result }
    try { $Parsed = $Data | ConvertFrom-Json } catch { return $Result }
    if ($null -eq $Parsed -or $Parsed -is [string] -or $Parsed -is [ValueType]) { return $Result }

    $ItemId = Get-Prop $Parsed 'pulse_id'
    if ($null -eq $ItemId) { $ItemId = Get-Prop $Parsed 'item_id' }
    if ($null -ne $ItemId) { $Result.item_id = "$ItemId" }

    $ColId = Get-Prop $Parsed 'column_id'
    if ($null -ne $ColId) { $Result.column_id = "$ColId" }
    $ColTitle = Get-Prop $Parsed 'column_title'
    if ($null -ne $ColTitle) { $Result.column_title = "$ColTitle" }

    $Result.previous_value_text = Get-LogValueText (Get-Prop $Parsed 'previous_value')
    $Result.value_text          = Get-LogValueText (Get-Prop $Parsed 'value')
    return $Result
}

# ── Row builders ──────────────────────────────────────────────
function Build-ColumnValueRow {
    param([string]$BoardId, [string]$ItemId, $ColumnValue, [hashtable]$ColumnTitles)
    $ColId = "$(Get-Prop $ColumnValue 'id')"
    $Title = if ($ColumnTitles -and $ColumnTitles.ContainsKey($ColId)) { $ColumnTitles[$ColId] } else { "" }
    $Text  = Get-Prop $ColumnValue 'text'
    # Mirror / board-relation / dependency columns return null `text`
    if ([string]::IsNullOrEmpty($Text)) { $Text = Get-Prop $ColumnValue 'display_value' }
    $Row = @(
        $BoardId, $ItemId, $ColId, $Title,
        "$(Get-Prop $ColumnValue 'type')", "$Text", "$(Get-Prop $ColumnValue 'value')"
    ) | ForEach-Object { ConvertTo-CsvField $_ }
    return $Row -join ","
}

function Build-ActivityRow {
    param([string]$BoardId, $LogEntry, [hashtable]$ColumnTitles)
    $DataJson = "$(Get-Prop $LogEntry 'data')"
    $Parsed   = ConvertFrom-ActivityData $DataJson
    $Title    = $Parsed.column_title
    if (-not $Title -and $Parsed.column_id -and $ColumnTitles -and $ColumnTitles.ContainsKey($Parsed.column_id)) {
        $Title = $ColumnTitles[$Parsed.column_id]
    }
    $Row = @(
        "$(Get-Prop $LogEntry 'id')", $BoardId, $Parsed.item_id,
        "$(Get-Prop $LogEntry 'event')", "$(Get-Prop $LogEntry 'user_id')",
        (ConvertFrom-MondayLogTimestamp "$(Get-Prop $LogEntry 'created_at')"),
        $Parsed.column_id, $Title, $Parsed.previous_value_text, $Parsed.value_text, $DataJson
    ) | ForEach-Object { ConvertTo-CsvField $_ }
    return $Row -join ","
}

function ConvertTo-CsvBool { param($Value) if ($Value) { "true" } else { "false" } }

# Invoke-RestMethod (pwsh 7) turns ISO date strings into [DateTime], which would
# otherwise stringify in the local culture ("09/03/2024 12:10:04"). Re-emit ISO UTC.
function ConvertTo-IsoUtc {
    param($Value)
    if ($Value -is [DateTimeOffset]) { return $Value.UtcDateTime.ToString("yyyy-MM-ddTHH:mm:ssZ") }
    if ($Value -is [DateTime])       { return $Value.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ") }
    return "$Value"
}

# ── Step 1: Resolve the RCM workspace ────────────────────────
if (-not $WorkspaceId) {
    Write-Host "Resolving workspace '$WorkspaceName'..."
    $Ws = (Invoke-GQL 'query { workspaces(limit: 100) { id name } }').data.workspaces |
        Where-Object { $_.name -eq $WorkspaceName }
    $Ws = @($Ws)
    if ($Ws.Count -eq 0) { throw "No workspace named '$WorkspaceName' is visible to this token. Set MONDAY_RCM_WORKSPACE_ID or check access." }
    if ($Ws.Count -gt 1) { throw "Multiple workspaces named '$WorkspaceName' ($($Ws.id -join ', ')). Set MONDAY_RCM_WORKSPACE_ID." }
    $WorkspaceId = $Ws[0].id
}
Write-Host "Workspace ID: $WorkspaceId"

# ── Step 2: Activity-log window ──────────────────────────────
$ActivityTo = $RunStart.ToString("yyyy-MM-ddTHH:mm:ssZ")
if (Test-Path $StateFile) {
    $ActivityFrom = (Get-Content $StateFile -Raw | ConvertFrom-Json).last_activity_to
    # ConvertFrom-Json in pwsh may have parsed it into a DateTime
    if ($ActivityFrom -is [DateTime]) { $ActivityFrom = $ActivityFrom.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ") }
} else {
    $ActivityFrom = $RunStart.AddDays(-$InitialLookbackDays).ToString("yyyy-MM-ddTHH:mm:ssZ")
    Write-Host "No rcm_state.json found — first run, looking back $InitialLookbackDays day(s)"
}
Write-Host "Activity window: $ActivityFrom -> $ActivityTo"

# ── Step 3: Boards + column titles ───────────────────────────
Write-Host "Fetching boards..."
$Boards = [System.Collections.Generic.List[object]]::new()
$Page = 1
do {
    $Q = "query { boards(workspace_ids: [$WorkspaceId], state: active, limit: 100, page: $Page) { id name columns { id title type } } }"
    $Batch = @((Invoke-GQL $Q).data.boards)
    foreach ($b in $Batch) { $Boards.Add($b) }
    $Page++
    Start-Sleep -Milliseconds 150
} while ($Batch.Count -eq 100)
Write-Host "  Found $($Boards.Count) board(s)"
# RCM is a closed workspace: boards are only visible to subscribers, even for
# account admins. Zero boards almost always means missing access — fail rather
# than advance the watermark past events we never saw.
if ($Boards.Count -eq 0) {
    throw "No boards visible in workspace $WorkspaceId. The token's user must be subscribed to (or have read access on) the RCM boards."
}

$ItemLines     = [System.Collections.Generic.List[string]]::new()
$ValueLines    = [System.Collections.Generic.List[string]]::new()
$ActivityLines = [System.Collections.Generic.List[string]]::new()

$ItemsHeader    = '"board_id","board_name","group_title","item_id","item_name","state","created_at_utc","updated_at_utc","creator_id","item_url"'
$ValuesHeader   = '"board_id","item_id","column_id","column_title","column_type","text","value_json"'
$ActivityHeader = '"event_id","board_id","item_id","event","user_id","created_at_utc","column_id","column_title","previous_value_text","value_text","data_json"'
$UsersHeader    = '"user_id","name","email","title","enabled","is_admin","is_guest","teams","created_at","last_activity"'

$ItemFields = @"
id name state created_at updated_at creator_id url
group { title }
column_values {
  id type text value
  ... on MirrorValue { display_value }
  ... on BoardRelationValue { display_value }
  ... on DependencyValue { display_value }
}
"@

foreach ($Board in $Boards) {
    $BoardId   = "$($Board.id)"
    $BoardName = "$($Board.name)"
    Write-Host "Board: $BoardName ($BoardId)"

    $ColumnTitles = @{}
    foreach ($c in @($Board.columns)) { $ColumnTitles["$($c.id)"] = "$($c.title)" }

    # ── Items (cursor pagination) ──
    $Q = "query { boards(ids: [$BoardId]) { items_page(limit: 500) { cursor items { $ItemFields } } } }"
    $ItemsPage = (Invoke-GQL $Q).data.boards[0].items_page
    $ItemCount = 0
    while ($true) {
        foreach ($Item in @($ItemsPage.items)) {
            $Group = Get-Prop $Item 'group'
            $Row = @(
                $BoardId, $BoardName, "$(Get-Prop $Group 'title')", "$($Item.id)", "$($Item.name)",
                "$(Get-Prop $Item 'state')", (ConvertTo-IsoUtc (Get-Prop $Item 'created_at')), (ConvertTo-IsoUtc (Get-Prop $Item 'updated_at')),
                "$(Get-Prop $Item 'creator_id')", "$(Get-Prop $Item 'url')"
            ) | ForEach-Object { ConvertTo-CsvField $_ }
            $ItemLines.Add($Row -join ",")
            foreach ($cv in @($Item.column_values)) {
                $ValueLines.Add((Build-ColumnValueRow $BoardId "$($Item.id)" $cv $ColumnTitles))
            }
            $ItemCount++
        }
        if (-not $ItemsPage.cursor) { break }
        Start-Sleep -Milliseconds 150
        $Q = "query { next_items_page(limit: 500, cursor: $(ConvertTo-GqlString $ItemsPage.cursor)) { cursor items { $ItemFields } } }"
        $ItemsPage = (Invoke-GQL $Q).data.next_items_page
    }

    # ── Activity log (page until a short page) ──
    $LogLimit = 1000
    $Page = 1
    $EventCount = 0
    do {
        $Q = "query { boards(ids: [$BoardId]) { activity_logs(from: $(ConvertTo-GqlString $ActivityFrom), to: $(ConvertTo-GqlString $ActivityTo), limit: $LogLimit, page: $Page) { id event data entity user_id created_at } } }"
        $Logs = @((Invoke-GQL $Q).data.boards[0].activity_logs)
        foreach ($Log in $Logs) {
            if ($null -eq $Log) { continue }
            $ActivityLines.Add((Build-ActivityRow $BoardId $Log $ColumnTitles))
            $EventCount++
        }
        $Page++
        Start-Sleep -Milliseconds 150
    } while ($Logs.Count -eq $LogLimit)

    Write-Host "  $ItemCount item(s), $EventCount activity event(s)"
    Start-Sleep -Milliseconds 300
}

# ── Step 4: Users with job titles ────────────────────────────
Write-Host "Fetching users..."
$UserLines = [System.Collections.Generic.List[string]]::new()
$Page = 1
do {
    $Q = "query { users(limit: 100, page: $Page) { id name email title enabled is_admin is_guest created_at last_activity teams { name } } }"
    $Batch = @((Invoke-GQL $Q).data.users)
    foreach ($u in $Batch) {
        $Teams = (@(Get-Prop $u 'teams') | Where-Object { $_ } | ForEach-Object { $_.name }) -join "; "
        $Row = @(
            "$($u.id)", "$($u.name)", "$($u.email)", "$(Get-Prop $u 'title')",
            (ConvertTo-CsvBool $u.enabled), (ConvertTo-CsvBool $u.is_admin), (ConvertTo-CsvBool $u.is_guest),
            $Teams, (ConvertTo-IsoUtc (Get-Prop $u 'created_at')), (ConvertTo-IsoUtc (Get-Prop $u 'last_activity'))
        ) | ForEach-Object { ConvertTo-CsvField $_ }
        $UserLines.Add($Row -join ",")
    }
    $Page++
    Start-Sleep -Milliseconds 150
} while ($Batch.Count -eq 100)
Write-Host "  $($UserLines.Count) user(s)"

# ── Step 5: Write files, then advance the watermark ──────────
# Header-only files are still written so warehouse loaders never miss a file.
@($ItemsHeader)    + $ItemLines     | Set-Content -Path $ItemsFile    -Encoding UTF8
@($ValuesHeader)   + $ValueLines    | Set-Content -Path $ValuesFile   -Encoding UTF8
@($ActivityHeader) + $ActivityLines | Set-Content -Path $ActivityFile -Encoding UTF8
@($UsersHeader)    + $UserLines     | Set-Content -Path $UsersFile    -Encoding UTF8

# Only reached if every board succeeded (any error above throws),
# so a failed run re-pulls the same activity window next time.
@{ last_activity_to = $ActivityTo } | ConvertTo-Json | Set-Content -Path $StateFile -Encoding UTF8

Write-Host ""
Write-Host "Done! Files written to $ReportsDir"
Write-Host "  $(Split-Path -Leaf $ItemsFile)     ($($ItemLines.Count) rows)"
Write-Host "  $(Split-Path -Leaf $ValuesFile)    ($($ValueLines.Count) rows)"
Write-Host "  $(Split-Path -Leaf $ActivityFile)  ($($ActivityLines.Count) rows)"
Write-Host "  $(Split-Path -Leaf $UsersFile)     ($($UserLines.Count) rows)"
Write-Host "  Watermark advanced to $ActivityTo"
