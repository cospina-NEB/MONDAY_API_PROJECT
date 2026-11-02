# ============================================================
# monday_rcm_report.ps1
# READ-ONLY report on the CORAL -- RCM workspace: for each board,
# its items, their column values (status, work queue, dates,
# owner, team), the board activity log (who changed what, and
# when), plus the user list with job titles.
#
# Only GraphQL `query` operations are issued — nothing in monday
# is created or changed.
#
# Usage:
#   .\scripts\monday_rcm_report.ps1 [-ActivityDays 30]
#
# Output (Reports/monday_rcm/):
#   coral_rcm_report_YYYY-MM-DD.html           executive dashboard
#   coral_rcm_report_YYYY-MM-DD.xlsx           Items / Activity / Users sheets
#   coral_rcm_items_YYYY-MM-DD.csv             raw: one row per item
#   coral_rcm_column_values_YYYY-MM-DD.csv     raw: one row per item x column
#   coral_rcm_activity_YYYY-MM-DD.csv          raw: one row per activity-log event
#   coral_rcm_users_YYYY-MM-DD.csv             raw: users with job titles
# ============================================================

param(
    # How many days of board activity log to include
    [int]$ActivityDays = 30
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
# MONDAY_RCM_API_TOKEN lets this report run as a different monday user
# (e.g. an RCM team member, or Coral's data_loader@ account) without
# touching MONDAY_API_TOKEN, which the monthly user report uses.
$Token   = if ($env:MONDAY_RCM_API_TOKEN) { $env:MONDAY_RCM_API_TOKEN } else { $env:MONDAY_API_TOKEN }
if (-not $Token) { throw "Set MONDAY_RCM_API_TOKEN or MONDAY_API_TOKEN env var first" }

$WorkspaceId   = $env:MONDAY_RCM_WORKSPACE_ID
$WorkspaceName = if ($env:MONDAY_RCM_WORKSPACE_NAME) { $env:MONDAY_RCM_WORKSPACE_NAME } else { "CORAL -- RCM" }

$RunStart   = [DateTime]::UtcNow
$Date       = Get-Date -Format "yyyy-MM-dd"
$ReportsDir = Join-Path $ProjectRoot "Reports\monday_rcm"
if (-not (Test-Path $ReportsDir)) { New-Item -ItemType Directory -Path $ReportsDir | Out-Null }

$ItemsFile    = Join-Path $ReportsDir "coral_rcm_items_$Date.csv"
$ValuesFile   = Join-Path $ReportsDir "coral_rcm_column_values_$Date.csv"
$ActivityFile = Join-Path $ReportsDir "coral_rcm_activity_$Date.csv"
$UsersFile    = Join-Path $ReportsDir "coral_rcm_users_$Date.csv"
$HtmlFile     = Join-Path $ReportsDir "coral_rcm_report_$Date.html"
$ExcelFile    = Join-Path $ReportsDir "coral_rcm_report_$Date.xlsx"

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
    # Empty / all-null objects ({} or {"name":null}) mean "no value";
    # a single scalar field ({"workspace_id":15742286}) is its own text.
    if ($Value -isnot [array]) {
        $Props   = @($Value.PSObject.Properties)
        $NonNull = @($Props | Where-Object { $null -ne $_.Value })
        if ($NonNull.Count -eq 0) { return "" }
        if ($Props.Count -eq 1 -and ($NonNull[0].Value -is [string] -or $NonNull[0].Value -is [ValueType])) {
            return "$($NonNull[0].Value)"
        }
    }
    return ($Value | ConvertTo-Json -Depth 10 -Compress)
}

# ── Helper: parse an activity-log `data` JSON string ─────────
function ConvertFrom-ActivityData {
    param([string]$Data)
    $Result = [ordered]@{ item_id = ""; item_name = ""; column_id = ""; column_title = ""; previous_value_text = ""; value_text = "" }
    if ([string]::IsNullOrWhiteSpace($Data)) { return $Result }
    try { $Parsed = $Data | ConvertFrom-Json } catch { return $Result }
    if ($null -eq $Parsed -or $Parsed -is [string] -or $Parsed -is [ValueType]) { return $Result }

    # `pulse_id` is the item. Some board-level events (e.g. subscribe) carry
    # "pulse_id": null plus an `item_id` that is really the board ID
    # (item_type "Board"), so only fall back to item_id for non-board events.
    $ItemId = Get-Prop $Parsed 'pulse_id'
    if ($null -eq $ItemId -and "$(Get-Prop $Parsed 'item_type')" -ne 'Board') { $ItemId = Get-Prop $Parsed 'item_id' }
    if ($null -ne $ItemId) { $Result.item_id = "$ItemId" }
    # Name at the time of the event — still readable after the item is deleted
    $ItemName = Get-Prop $Parsed 'pulse_name'
    if ($null -ne $ItemName) { $Result.item_name = "$ItemName" }

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
        "$(Get-Prop $LogEntry 'id')", $BoardId, $Parsed.item_id, $Parsed.item_name,
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

# ── Report helpers ────────────────────────────────────────────
# Friendly names for the activity-log events RCM leaders care about;
# anything else is shown as monday's raw event name.
function Get-EventLabel {
    param([string]$EventName)
    $Labels = @{
        update_column_value   = "Changed value"
        create_pulse          = "Created item"
        update_name           = "Renamed item"
        move_pulse_into_group = "Moved to group"
        move_pulse_into_board = "Moved to board"
        archive_pulse         = "Archived item"
        restore_pulse         = "Restored item"
        delete_pulse          = "Deleted item"
        create_update         = "Posted update"
        subscribe             = "Subscribed"
        unsubscribe           = "Unsubscribed"
        create_column         = "Added column"
        delete_column         = "Removed column"
        create_group          = "Added group"
        update_board_name     = "Renamed board"
        board_workspace_id_changed = "Moved board to workspace"
    }
    if ($Labels.ContainsKey($EventName)) { return $Labels[$EventName] } else { return $EventName }
}

function ConvertTo-HtmlText { param([string]$Value) [System.Net.WebUtility]::HtmlEncode($Value) }

# People-column changes are logged as "person:<id>; team:<id>" (the raw CSV
# keeps that for joins); in the report, show names where we know them.
function Resolve-PeopleText {
    param([string]$Value, [hashtable]$UserNames)
    if (-not $Value -or -not $UserNames) { return $Value }
    return [regex]::Replace($Value, 'person:(\d+)', {
        param($m)
        if ($UserNames.ContainsKey($m.Groups[1].Value)) { $UserNames[$m.Groups[1].Value] } else { $m.Value }
    })
}

# ISO UTC ("2026-05-28T18:58:06.586Z") -> "2026-05-28 18:58"
function Format-ReportTime {
    param([string]$Value)
    if ($Value -match '^(\d{4}-\d{2}-\d{2})T(\d{2}:\d{2})') { return "$($Matches[1]) $($Matches[2])" }
    return $Value
}

# Pivot the long column-values CSV into one row per item, with one column
# per board column title (union across boards, in first-seen order).
# Every row gets every key so Export-Excel sees the full header.
function ConvertTo-WideItems {
    param($Items, $Values, [hashtable]$UserNames)
    $Fixed        = @('Board', 'Group', 'Item', 'Created (UTC)', 'Updated (UTC)', 'Created By', 'Link')
    $ValuesByItem = @{}
    $ColOrder     = [System.Collections.Generic.List[string]]::new()
    foreach ($v in @($Values)) {
        if ($v.column_type -eq 'subtasks') { continue }
        if (-not $ValuesByItem.ContainsKey($v.item_id)) { $ValuesByItem[$v.item_id] = [ordered]@{} }
        $Key = if ($v.column_title) { $v.column_title } else { $v.column_id }
        # Duplicate titles on one board, or a board column named like a fixed field
        if ($ValuesByItem[$v.item_id].Contains($Key) -or $Fixed -contains $Key) { $Key = "$Key [$($v.column_id)]" }
        if (-not $ColOrder.Contains($Key)) { $ColOrder.Add($Key) }
        $ValuesByItem[$v.item_id][$Key] = $v.text
    }
    foreach ($i in @($Items)) {
        $Vals = $ValuesByItem[$i.item_id]
        $Row  = [ordered]@{ Board = $i.board_name; Group = $i.group_title; Item = $i.item_name }
        foreach ($k in $ColOrder) { $Row[$k] = if ($Vals -and $Vals.Contains($k)) { $Vals[$k] } else { "" } }
        $Row['Created (UTC)'] = Format-ReportTime $i.created_at_utc
        $Row['Updated (UTC)'] = Format-ReportTime $i.updated_at_utc
        $Row['Created By']    = if ($UserNames -and $UserNames.ContainsKey($i.creator_id)) { $UserNames[$i.creator_id] } else { $i.creator_id }
        $Row['Link']          = $i.item_url
        [PSCustomObject]$Row
    }
}

# ── Step 0: Whose access is this run using? ──────────────────
# Board visibility depends entirely on the token's user, so record it.
$Me = (Invoke-GQL 'query { me { name email } }').data.me
$RunAs = "$($Me.email)"
$TokenSource = if ($env:MONDAY_RCM_API_TOKEN) { "MONDAY_RCM_API_TOKEN" } else { "MONDAY_API_TOKEN" }
Write-Host "Running as: $($Me.name) <$RunAs> (from $TokenSource)"

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
$ActivityTo   = $RunStart.ToString("yyyy-MM-ddTHH:mm:ssZ")
$ActivityFrom = $RunStart.AddDays(-$ActivityDays).ToString("yyyy-MM-ddTHH:mm:ssZ")
Write-Host "Activity window: $ActivityFrom -> $ActivityTo ($ActivityDays day(s))"

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
# than produce an empty report that looks like "no activity".
if ($Boards.Count -eq 0) {
    throw "No boards visible to $RunAs in workspace $WorkspaceId. That user must be a member of the workspace (directly or via one of its teams), or set MONDAY_RCM_API_TOKEN to the token of someone who is."
}

$ItemLines     = [System.Collections.Generic.List[string]]::new()
$ValueLines    = [System.Collections.Generic.List[string]]::new()
$ActivityLines = [System.Collections.Generic.List[string]]::new()

$ItemsHeader    = '"board_id","board_name","group_title","item_id","item_name","state","created_at_utc","updated_at_utc","creator_id","item_url"'
$ValuesHeader   = '"board_id","item_id","column_id","column_title","column_type","text","value_json"'
$ActivityHeader = '"event_id","board_id","item_id","item_name","event","user_id","created_at_utc","column_id","column_title","previous_value_text","value_text","data_json"'
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

# ── Step 5: Write raw CSVs ────────────────────────────────────
@($ItemsHeader)    + $ItemLines     | Set-Content -Path $ItemsFile    -Encoding UTF8
@($ValuesHeader)   + $ValueLines    | Set-Content -Path $ValuesFile   -Encoding UTF8
@($ActivityHeader) + $ActivityLines | Set-Content -Path $ActivityFile -Encoding UTF8
@($UsersHeader)    + $UserLines     | Set-Content -Path $UsersFile    -Encoding UTF8
Write-Host "Raw CSVs written to $ReportsDir"

# ── Step 6: Generate HTML report ─────────────────────────────
Write-Host "Generating HTML report..."

$Items    = @(Import-Csv -Path $ItemsFile    -Encoding UTF8)
$Values   = @(Import-Csv -Path $ValuesFile   -Encoding UTF8)
$Activity = @(Import-Csv -Path $ActivityFile -Encoding UTF8 | Sort-Object created_at_utc -Descending)
$Users    = @(Import-Csv -Path $UsersFile    -Encoding UTF8)

$UserById   = @{}; foreach ($u in $Users) { $UserById[$u.user_id] = $u }
$UserNames  = @{}; foreach ($u in $Users) { $UserNames[$u.user_id] = $u.name }
$BoardNames = @{}; foreach ($b in $Boards) { $BoardNames["$($b.id)"] = "$($b.name)" }
$ItemNames  = @{}; foreach ($i in $Items) { $ItemNames[$i.item_id] = $i.item_name }

function Get-UserName  { param([string]$Id) if ($UserById.ContainsKey($Id)) { $UserById[$Id].name }  elseif ($Id) { "User $Id" } else { "(system)" } }
function Get-UserTitle { param([string]$Id) if ($UserById.ContainsKey($Id)) { $UserById[$Id].title } else { "" } }
# Prefer the item's current name; fall back to its name at the time of the event (deleted/moved items)
function Get-ItemName  { param($Evt) if ($Evt.item_id -and $ItemNames.ContainsKey($Evt.item_id)) { $ItemNames[$Evt.item_id] } else { $Evt.item_name } }

function Get-SortableTh {
    param([string]$Label, [switch]$Num)
    $type = if ($Num) { " data-type='num'" } else { "" }
    "<th class='sortable' onclick='sortTable(this)'$type>$Label<span class='sort-icon'>&#8645;</span></th>"
}

# ── Helper: render a table section or an empty-state message ──
function New-HtmlSection {
    param([string]$Title, [string]$Head, [string[]]$Rows, [string]$EmptyMsg, [string]$Intro = "", [string]$TableId = "")
    $rowCount = ($Rows | Measure-Object).Count
    $idAttr   = if ($TableId) { " id='$TableId'" } else { "" }
    $inner = if ($rowCount -gt 0) {
        "$Intro<table$idAttr><thead><tr>$Head</tr></thead><tbody>$($Rows -join '')</tbody></table>"
    } else {
        "<p class='empty'>$EmptyMsg</p>"
    }
    "<section><h2>$Title</h2>$inner</section>"
}

$H = { param($v) ConvertTo-HtmlText $v }   # short alias for escaping cell text

# ── Boards ──
$BoardRows = foreach ($b in ($Boards | Sort-Object name)) {
    $bid      = "$($b.id)"
    $bItems   = @($Items    | Where-Object { $_.board_id -eq $bid })
    $bEvents  = @($Activity | Where-Object { $_.board_id -eq $bid })
    $last     = if ($bEvents.Count) { Format-ReportTime $bEvents[0].created_at_utc } else { "&mdash;" }
    $url      = "https://coral.monday.com/boards/$bid"
    "<tr><td><a href='$url' target='_blank'>$(& $H $b.name)</a></td><td class='num'>$($bItems.Count)</td><td class='num'>$($bEvents.Count)</td><td>$last</td></tr>"
}
$BoardSection = New-HtmlSection -Title "Boards" `
    -Head ((Get-SortableTh "Board") + (Get-SortableTh "Items" -Num) + (Get-SortableTh "Changes ($ActivityDays d)" -Num) + (Get-SortableTh "Last Change (UTC)")) `
    -Rows $BoardRows -EmptyMsg "No boards."

# ── Activity by associate ──
$ByUser = $Activity | Group-Object user_id | Sort-Object Count -Descending
$AssociateRows = foreach ($g in $ByUser) {
    $touched = @($g.Group | Where-Object { $_.item_id } | Select-Object -ExpandProperty item_id -Unique).Count
    $changes = @($g.Group | Where-Object { $_.event -eq 'update_column_value' }).Count
    "<tr><td>$(& $H (Get-UserName $g.Name))</td><td>$(& $H (Get-UserTitle $g.Name))</td><td class='num'>$($g.Count)</td><td class='num'>$changes</td><td class='num'>$touched</td><td>$(Format-ReportTime $g.Group[0].created_at_utc)</td></tr>"
}
$AssociateSection = New-HtmlSection -Title "Activity by Associate &mdash; Last $ActivityDays Days" `
    -Head ((Get-SortableTh "Associate") + (Get-SortableTh "Job Title") + (Get-SortableTh "Events" -Num) + (Get-SortableTh "Value Changes" -Num) + (Get-SortableTh "Items Touched" -Num) + (Get-SortableTh "Last Activity (UTC)")) `
    -Rows $AssociateRows -EmptyMsg "No board activity in the last $ActivityDays days."

# ── Status breakdown (every status-type column, e.g. Status, Work Queue) ──
$StatusGroups = $Values | Where-Object { $_.column_type -eq 'status' } |
    Group-Object board_id, column_title, text |
    Sort-Object @{ e = { $BoardNames[$_.Group[0].board_id] } }, @{ e = { $_.Group[0].column_title } }, @{ e = 'Count'; Descending = $true }
$StatusRows = foreach ($g in $StatusGroups) {
    $v     = $g.Group[0]
    $label = if ($v.text) { & $H $v.text } else { "<span class='muted'>(blank)</span>" }
    "<tr><td>$(& $H $BoardNames[$v.board_id])</td><td>$(& $H $v.column_title)</td><td>$label</td><td class='num'>$($g.Count)</td></tr>"
}
$StatusSection = New-HtmlSection -Title "Items by Status" `
    -Head ((Get-SortableTh "Board") + (Get-SortableTh "Column") + (Get-SortableTh "Value") + (Get-SortableTh "Items" -Num)) `
    -Rows $StatusRows -EmptyMsg "No status columns found."

# ── Owner / team breakdown (every people-type column) ──
$OwnerPairs = foreach ($v in ($Values | Where-Object { $_.column_type -eq 'people' })) {
    $names = if ($v.text) { $v.text -split ',\s*' } else { @('(unassigned)') }
    foreach ($n in $names) { [PSCustomObject]@{ board_id = $v.board_id; column_title = $v.column_title; owner = $n } }
}
$OwnerRows = foreach ($g in (@($OwnerPairs) | Where-Object { $_ } | Group-Object board_id, column_title, owner | Sort-Object Count -Descending)) {
    $v = $g.Group[0]
    "<tr><td>$(& $H $BoardNames[$v.board_id])</td><td>$(& $H $v.column_title)</td><td>$(& $H $v.owner)</td><td class='num'>$($g.Count)</td></tr>"
}
$OwnerSection = New-HtmlSection -Title "Items by Owner / Team" `
    -Head ((Get-SortableTh "Board") + (Get-SortableTh "Column") + (Get-SortableTh "Assigned To") + (Get-SortableTh "Items" -Num)) `
    -Rows $OwnerRows -EmptyMsg "No people columns found."

# ── Recent changes (filterable by associate / board) ──
$RecentLimit = 500
$Recent      = @($Activity | Select-Object -First $RecentLimit)
$ChangeRows = foreach ($e in $Recent) {
    $who   = Get-UserName $e.user_id
    $board = $BoardNames[$e.board_id]
    "<tr data-user='$(& $H $who)' data-board='$(& $H $board)'><td>$(Format-ReportTime $e.created_at_utc)</td><td>$(& $H $who)</td><td>$(& $H $board)</td><td>$(& $H (Get-ItemName $e))</td><td>$(& $H (Get-EventLabel $e.event))</td><td>$(& $H $e.column_title)</td><td class='from'>$(& $H (Resolve-PeopleText $e.previous_value_text $UserNames))</td><td>$(& $H (Resolve-PeopleText $e.value_text $UserNames))</td></tr>"
}
$UserOptions  = (@($Recent | ForEach-Object { Get-UserName $_.user_id }) | Sort-Object -Unique | ForEach-Object { "<option value='$(& $H $_)'>$(& $H $_)</option>" }) -join ""
$BoardOptions = (@($Recent | ForEach-Object { $BoardNames[$_.board_id] }) | Sort-Object -Unique | ForEach-Object { "<option value='$(& $H $_)'>$(& $H $_)</option>" }) -join ""
$Truncated    = if ($Activity.Count -gt $RecentLimit) { " &mdash; showing latest $RecentLimit of $($Activity.Count); full log in the Excel/CSV" } else { "" }
$FilterBar = @"
<div class="filter-bar">
  <label for="user-filter">Associate</label>
  <select id="user-filter" onchange="filterChanges()"><option value="">All</option>$UserOptions</select>
  <label for="board-filter">Board</label>
  <select id="board-filter" onchange="filterChanges()"><option value="">All</option>$BoardOptions</select>
  <span class="filter-count" id="change-count">$($Recent.Count) event(s)</span>
</div>
"@
$ChangeSection = New-HtmlSection -Title "Change Log &mdash; Who Changed What, When$Truncated" `
    -Head ((Get-SortableTh "When (UTC)") + (Get-SortableTh "Associate") + (Get-SortableTh "Board") + (Get-SortableTh "Item") + (Get-SortableTh "Change") + (Get-SortableTh "Column") + (Get-SortableTh "From") + (Get-SortableTh "To")) `
    -Rows $ChangeRows -Intro $FilterBar -TableId "change-table" -EmptyMsg "No board activity in the last $ActivityDays days."

# ── Users & job titles ──
$UserRows = foreach ($u in ($Users | Sort-Object name)) {
    $st      = if ($u.enabled -eq 'true') { "<span class='active'>Active</span>" } else { "<span class='warn'>Inactive</span>" }
    $titleTd = if ($u.title) { & $H $u.title } else { "<span class='muted'>&mdash;</span>" }
    "<tr><td>$(& $H $u.name)</td><td>$(& $H $u.email)</td><td>$titleTd</td><td>$st</td><td>$(& $H $u.teams)</td><td>$(Format-ReportTime $u.last_activity)</td></tr>"
}
$UserSection = New-HtmlSection -Title "Users &amp; Job Titles" `
    -Head ((Get-SortableTh "Name") + (Get-SortableTh "Email") + (Get-SortableTh "Job Title") + (Get-SortableTh "Status") + (Get-SortableTh "Teams") + (Get-SortableTh "Last Active (UTC)")) `
    -Rows $UserRows -EmptyMsg "No users."

$WsLabel  = & $H $WorkspaceName
$ActiveAssociates = @($ByUser).Count
$Html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>Coral RCM Activity Report - $Date</title>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: 'Segoe UI', Arial, sans-serif; background: #f4f5f7; color: #172b4d; }
  header { background: #0052cc; color: #fff; padding: 24px 40px; }
  header h1 { font-size: 22px; font-weight: 600; }
  header p  { font-size: 13px; opacity: 0.75; margin-top: 5px; }
  main { padding: 32px 40px; max-width: 1400px; }
  .cards { display: flex; gap: 16px; margin-bottom: 28px; flex-wrap: wrap; }
  .card { background: #fff; border-radius: 6px; box-shadow: 0 1px 3px rgba(0,0,0,.12); padding: 16px 20px; min-width: 160px; }
  .card .n { font-size: 26px; font-weight: 600; color: #0052cc; }
  .card .l { font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: .05em; color: #5e6c84; margin-top: 4px; }
  section { background: #fff; border-radius: 6px; box-shadow: 0 1px 3px rgba(0,0,0,.12); margin-bottom: 28px; overflow: hidden; }
  section h2 { font-size: 12px; font-weight: 700; text-transform: uppercase; letter-spacing: .06em;
               color: #5e6c84; padding: 14px 20px; border-bottom: 1px solid #ebecf0; }
  table { width: 100%; border-collapse: collapse; font-size: 13px; }
  th { background: #f4f5f7; color: #5e6c84; font-weight: 700; text-transform: uppercase;
       font-size: 11px; letter-spacing: .05em; padding: 10px 16px; text-align: left; }
  td { padding: 9px 16px; border-top: 1px solid #ebecf0; vertical-align: middle; }
  tr:hover td { background: #fafbfc; }
  td.num { text-align: center; }
  td.from { color: #7a869a; }
  .active { color: #006644; font-weight: 600; }
  .warn   { color: #bf2600; font-weight: 600; }
  .muted  { color: #a5adba; }
  a { color: #0052cc; text-decoration: none; }
  a:hover { text-decoration: underline; }
  .empty { color: #7a869a; font-style: italic; padding: 16px 20px; font-size: 13px; }
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
  <h1>Coral &mdash; RCM Activity Report</h1>
  <p>$WsLabel &nbsp;&bull;&nbsp; Generated $Date &nbsp;&bull;&nbsp; Activity window: last $ActivityDays days ($(Format-ReportTime $ActivityFrom) &ndash; $(Format-ReportTime $ActivityTo) UTC) &nbsp;&bull;&nbsp; Read-only extract from monday.com as $(& $H $RunAs)</p>
</header>
<main>
<div class="cards">
  <div class="card"><div class="n">$($Boards.Count)</div><div class="l">Boards</div></div>
  <div class="card"><div class="n">$($Items.Count)</div><div class="l">Items</div></div>
  <div class="card"><div class="n">$($Activity.Count)</div><div class="l">Activity Events</div></div>
  <div class="card"><div class="n">$ActiveAssociates</div><div class="l">Active Associates</div></div>
  <div class="card"><div class="n">$($Users.Count)</div><div class="l">Users</div></div>
</div>
$AssociateSection
$ChangeSection
$StatusSection
$OwnerSection
$BoardSection
$UserSection
</main>
<script>
function filterChanges() {
  var u = document.getElementById('user-filter').value;
  var b = document.getElementById('board-filter').value;
  var rows = document.querySelectorAll('#change-table tbody tr');
  var visible = 0;
  rows.forEach(function(row) {
    var show = (!u || row.dataset.user === u) && (!b || row.dataset.board === b);
    row.style.display = show ? '' : 'none';
    if (show) visible++;
  });
  document.getElementById('change-count').textContent = visible + ' event(s)';
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

# ── Step 7: Generate Excel report ─────────────────────────────
Write-Host "Generating Excel report..."

$ItemsSheet = @(ConvertTo-WideItems $Items $Values $UserNames)
$ActivitySheet = foreach ($e in $Activity) {
    [PSCustomObject][ordered]@{
        'When (UTC)' = Format-ReportTime $e.created_at_utc
        'Associate'  = Get-UserName $e.user_id
        'Job Title'  = Get-UserTitle $e.user_id
        'Board'      = $BoardNames[$e.board_id]
        'Item'       = Get-ItemName $e
        'Change'     = Get-EventLabel $e.event
        'Column'     = $e.column_title
        'From'       = Resolve-PeopleText $e.previous_value_text $UserNames
        'To'         = Resolve-PeopleText $e.value_text $UserNames
        'Item ID'    = $e.item_id
        'Event ID'   = $e.event_id
    }
}
$UsersSheet = foreach ($u in ($Users | Sort-Object name)) {
    [PSCustomObject][ordered]@{
        'Name'               = $u.name
        'Email'              = $u.email
        'Job Title'          = $u.title
        'Status'             = if ($u.enabled -eq 'true') { 'Active' } else { 'Inactive' }
        'Admin'              = $u.is_admin
        'Guest'              = $u.is_guest
        'Teams'              = $u.teams
        'Joined (UTC)'       = Format-ReportTime $u.created_at
        'Last Active (UTC)'  = Format-ReportTime $u.last_activity
    }
}

try {
    if (Test-Path $ExcelFile) { Remove-Item $ExcelFile -Force -ErrorAction Stop }
    $XlOpts = @{ Path = $ExcelFile; AutoFilter = $true; FreezeTopRow = $true; BoldTopRow = $true; AutoSize = $true }
    if ($ItemsSheet.Count)    { $ItemsSheet    | Export-Excel @XlOpts -WorksheetName "Items" }
    if (@($ActivitySheet).Count) { $ActivitySheet | Export-Excel @XlOpts -WorksheetName "Activity" }
    if (@($UsersSheet).Count) { $UsersSheet    | Export-Excel @XlOpts -WorksheetName "Users" }
    Write-Host "   Excel report saved to: $ExcelFile"
} catch {
    Write-Warning "Excel report skipped ($($_.Exception.Message)). Install ImportExcel, or close '$ExcelFile' in Excel and re-run."
}

Write-Host ""
Write-Host "Done! $($Boards.Count) board(s), $($Items.Count) item(s), $($Activity.Count) activity event(s), $($Users.Count) user(s)"
Write-Host "  Report: $HtmlFile"
