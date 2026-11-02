# ============================================================
# monday_rcm_report.Tests.ps1
# Pester v5 tests for monday_rcm_report.ps1
#
# Run with:
#   Invoke-Pester -Path tests/monday_rcm_report.Tests.ps1 -Output Detailed
# ============================================================

BeforeAll {
    # ── Load only the helper functions from the script ──
    # Parse the script and define each top-level function, without running its
    # body (which needs a live API token and network). Unlike copy-pasting the
    # helpers, this keeps the tests pinned to the real implementation.
    $ScriptPath = Join-Path $PSScriptRoot "..\scripts\monday_rcm_report.ps1"
    $Ast = [System.Management.Automation.Language.Parser]::ParseFile($ScriptPath, [ref]$null, [ref]$null)
    $Ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $false) |
        ForEach-Object { . ([scriptblock]::Create($_.Extent.Text)) }
}

Describe "ConvertFrom-MondayLogTimestamp" {
    It "converts a 17-digit 100ns timestamp to ISO UTC" {
        # 2026-10-30T12:34:56.789Z = 1793363696789 ms
        ConvertFrom-MondayLogTimestamp "17933636967890000" | Should -Be "2026-10-30T12:34:56.789Z"
    }
    It "passes through values that are not 17 digits" {
        ConvertFrom-MondayLogTimestamp "2026-10-30T12:00:00Z" | Should -Be "2026-10-30T12:00:00Z"
        ConvertFrom-MondayLogTimestamp "" | Should -Be ""
    }
}

Describe "ConvertFrom-ActivityData" {
    It "parses a status change (update_column_value)" {
        $Data = '{"board_id":1,"pulse_id":555,"pulse_name":"Claim 42","column_id":"status","column_title":"Status","column_type":"color","value":{"label":{"index":1,"text":"Done"}},"previous_value":{"label":{"index":0,"text":"Working on it"}}}'
        $r = ConvertFrom-ActivityData $Data
        $r.item_id             | Should -Be "555"
        $r.column_id           | Should -Be "status"
        $r.column_title        | Should -Be "Status"
        $r.previous_value_text | Should -Be "Working on it"
        $r.value_text          | Should -Be "Done"
    }
    It "parses date and people values" {
        $Data = '{"pulse_id":7,"column_id":"date4","value":{"date":"2026-10-01","time":null},"previous_value":{"personsAndTeams":[{"id":11,"kind":"person"},{"id":3,"kind":"team"}]}}'
        $r = ConvertFrom-ActivityData $Data
        $r.value_text          | Should -Be "2026-10-01"
        $r.previous_value_text | Should -Be "person:11; team:3"
    }
    It "handles create_pulse with no column" {
        $r = ConvertFrom-ActivityData '{"board_id":1,"pulse_id":9,"pulse_name":"New item"}'
        $r.item_id    | Should -Be "9"
        $r.column_id  | Should -Be ""
        $r.value_text | Should -Be ""
    }
    It "does not report a board ID as the item for board-level events (subscribe)" {
        $r = ConvertFrom-ActivityData '{"item_id":18415321643,"item_type":"Board","board_id":18415321643,"pulse_id":null}'
        $r.item_id | Should -Be ""
    }
    It "renders empty, all-null and single-field values as plain text" {
        $r = ConvertFrom-ActivityData '{"pulse_id":1,"previous_value":{},"value":{"workspace_id":15742286}}'
        $r.previous_value_text | Should -Be ""
        $r.value_text          | Should -Be "15742286"
        (ConvertFrom-ActivityData '{"previous_value":{"name":null}}').previous_value_text | Should -Be ""
    }
    It "returns blanks for empty or malformed JSON" {
        foreach ($Bad in "", "   ", "{not json", '"just a string"') {
            $r = ConvertFrom-ActivityData $Bad
            $r.item_id   | Should -Be ""
            $r.column_id | Should -Be ""
        }
    }
}

Describe "ConvertTo-IsoUtc" {
    It "formats a parsed UTC DateTime as ISO, not local culture" {
        $d = [DateTime]::SpecifyKind([DateTime]"2024-09-03 12:10:04", [DateTimeKind]::Utc)
        ConvertTo-IsoUtc $d | Should -Be "2024-09-03T12:10:04Z"
    }
    It "passes strings and nulls through" {
        ConvertTo-IsoUtc "2024-09-03" | Should -Be "2024-09-03"
        ConvertTo-IsoUtc $null        | Should -Be ""
    }
}

Describe "ConvertTo-GqlString" {
    It "escapes quotes and backslashes" {
        ConvertTo-GqlString 'a"b\c' | Should -Be '"a\"b\\c"'
    }
}

Describe "Build-ColumnValueRow" {
    It "uses the board column-title map and escapes value_json" {
        $cv  = [PSCustomObject]@{ id = "status"; type = "status"; text = "Done"; value = '{"index":1,"note":"a, b"}' }
        $row = Build-ColumnValueRow "100" "555" $cv @{ status = "Work Queue" }
        $row | Should -Be '"100","555","status","Work Queue","status","Done","{""index"":1,""note"":""a, b""}"'
    }
    It "leaves the title blank for an unknown column" {
        $cv  = [PSCustomObject]@{ id = "x"; type = "text"; text = "hi"; value = $null }
        Build-ColumnValueRow "1" "2" $cv @{} | Should -Be '"1","2","x","","text","hi",""'
    }
    It "falls back to display_value when text is null (mirror columns)" {
        $cv  = [PSCustomObject]@{ id = "m"; type = "mirror"; text = $null; value = $null; display_value = "Alice" }
        Build-ColumnValueRow "1" "2" $cv @{ m = "Owner" } | Should -Be '"1","2","m","Owner","mirror","Alice",""'
    }
}

Describe "Build-ActivityRow" {
    It "builds a full row, including item name, and falls back to the board map for column title" {
        $log = [PSCustomObject]@{
            id = "evt-1"; event = "update_column_value"; user_id = "11"; entity = "pulse"
            created_at = "17933636967890000"
            data = '{"pulse_id":555,"pulse_name":"Claim 42","column_id":"status","value":{"label":{"text":"Done"}}}'
        }
        $fields = (Build-ActivityRow "100" $log @{ status = "Work Queue" }) -split '","'
        $fields[0]  | Should -Be '"evt-1'
        $fields[2]  | Should -Be "555"
        $fields[3]  | Should -Be "Claim 42"
        $fields[6]  | Should -Be "2026-10-30T12:34:56.789Z"
        $fields[8]  | Should -Be "Work Queue"
        $fields[10] | Should -Be "Done"
    }
}

Describe "Get-EventLabel" {
    It "maps known events to friendly names and passes unknown ones through" {
        Get-EventLabel "update_column_value" | Should -Be "Changed value"
        Get-EventLabel "create_pulse"        | Should -Be "Created item"
        Get-EventLabel "some_new_event"      | Should -Be "some_new_event"
    }
}

Describe "Format-ReportTime" {
    It "shortens ISO UTC to date + minutes" {
        Format-ReportTime "2026-05-28T18:58:06.586Z" | Should -Be "2026-05-28 18:58"
        Format-ReportTime "2026-05-28T18:58:06Z"     | Should -Be "2026-05-28 18:58"
    }
    It "passes through blanks and non-ISO values" {
        Format-ReportTime ""        | Should -Be ""
        Format-ReportTime "unknown" | Should -Be "unknown"
    }
}

Describe "ConvertTo-HtmlText" {
    It "escapes markup in monday item names" {
        ConvertTo-HtmlText "<b>A & B's</b>" | Should -Be "&lt;b&gt;A &amp; B&#39;s&lt;/b&gt;"
    }
}

Describe "ConvertTo-WideItems" {
    BeforeAll {
        $Items = @(
            [PSCustomObject]@{ board_id = "1"; board_name = "Claims"; group_title = "Open"; item_id = "10"; item_name = "Claim A"; created_at_utc = "2026-10-01T09:00:00Z"; updated_at_utc = "2026-10-02T10:30:00Z"; creator_id = "7"; item_url = "u1" }
            [PSCustomObject]@{ board_id = "2"; board_name = "Denials"; group_title = "New"; item_id = "20"; item_name = "Denial B"; created_at_utc = ""; updated_at_utc = ""; creator_id = "99"; item_url = "u2" }
        )
        $Values = @(
            [PSCustomObject]@{ item_id = "10"; column_id = "status";  column_title = "Status";     column_type = "status";   text = "Done" }
            [PSCustomObject]@{ item_id = "10"; column_id = "queue";   column_title = "Work Queue"; column_type = "status";   text = "Billing" }
            [PSCustomObject]@{ item_id = "10"; column_id = "subs";    column_title = "Subitems";   column_type = "subtasks"; text = "" }
            [PSCustomObject]@{ item_id = "20"; column_id = "status";  column_title = "Status";     column_type = "status";   text = "Stuck" }
            [PSCustomObject]@{ item_id = "20"; column_id = "date1";   column_title = "Date";       column_type = "date";     text = "2026-10-05" }
            [PSCustomObject]@{ item_id = "20"; column_id = "date2";   column_title = "Date";       column_type = "date";     text = "2026-10-06" }
            [PSCustomObject]@{ item_id = "20"; column_id = "board_x"; column_title = "Board";      column_type = "text";     text = "clash" }
        )
        $Rows = @(ConvertTo-WideItems $Items $Values @{ "7" = "Alice" })
    }
    It "gives one row per item, with the same columns on every row" {
        $Rows.Count | Should -Be 2
        ($Rows[0].PSObject.Properties.Name -join "|") | Should -Be ($Rows[1].PSObject.Properties.Name -join "|")
    }
    It "puts each column title in its own column, merging same-titled columns across boards" {
        $Rows[0].Status       | Should -Be "Done"
        $Rows[1].Status       | Should -Be "Stuck"
        $Rows[0].'Work Queue' | Should -Be "Billing"
        $Rows[1].'Work Queue' | Should -Be ""
    }
    It "skips subitem columns and disambiguates duplicate or reserved titles" {
        $Rows[0].PSObject.Properties.Name | Should -Not -Contain "Subitems"
        $Rows[1].Date            | Should -Be "2026-10-05"
        $Rows[1].'Date [date2]'  | Should -Be "2026-10-06"
        $Rows[1].Board           | Should -Be "Denials"
        $Rows[1].'Board [board_x]' | Should -Be "clash"
    }
    It "resolves creator names and formats timestamps" {
        $Rows[0].'Created By'    | Should -Be "Alice"
        $Rows[1].'Created By'    | Should -Be "99"
        $Rows[0].'Updated (UTC)' | Should -Be "2026-10-02 10:30"
    }
}

Describe "Resolve-PeopleText" {
    It "replaces known person IDs with names and leaves teams/unknown IDs alone" {
        Resolve-PeopleText "person:7; team:3; person:99" @{ "7" = "Alice" } | Should -Be "Alice; team:3; person:99"
    }
    It "passes through blanks and non-people values" {
        Resolve-PeopleText ""     @{ "7" = "Alice" } | Should -Be ""
        Resolve-PeopleText "Done" @{ "7" = "Alice" } | Should -Be "Done"
    }
}
