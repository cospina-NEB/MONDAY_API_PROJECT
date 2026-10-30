# ============================================================
# monday_rcm_export.Tests.ps1
# Pester v5 tests for monday_rcm_export.ps1
#
# Run with:
#   Invoke-Pester -Path tests/monday_rcm_export.Tests.ps1 -Output Detailed
# ============================================================

BeforeAll {
    # ── Load only the helper functions from the script ──
    # Parse the script and define each top-level function, without running its
    # body (which needs a live API token and network). Unlike copy-pasting the
    # helpers, this keeps the tests pinned to the real implementation.
    $ScriptPath = Join-Path $PSScriptRoot "..\scripts\monday_rcm_export.ps1"
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
    It "builds a full row and falls back to the board map for column title" {
        $log = [PSCustomObject]@{
            id = "evt-1"; event = "update_column_value"; user_id = "11"; entity = "pulse"
            created_at = "17933636967890000"
            data = '{"pulse_id":555,"column_id":"status","value":{"label":{"text":"Done"}}}'
        }
        $fields = (Build-ActivityRow "100" $log @{ status = "Work Queue" }) -split '","'
        $fields[0]  | Should -Be '"evt-1'
        $fields[2]  | Should -Be "555"
        $fields[5]  | Should -Be "2026-10-30T12:34:56.789Z"
        $fields[7]  | Should -Be "Work Queue"
        $fields[9]  | Should -Be "Done"
    }
}
