# ============================================================
# monday_user_report.Tests.ps1
# Pester v5 tests for monday_user_report.ps1
#
# Run with:
#   Install-Module Pester -MinimumVersion 5.0 -Force
#   Invoke-Pester ./monday_user_report.Tests.ps1 -Output Detailed
# ============================================================

BeforeAll {
    # ── Dot-source only the helper functions from the script ──
    # We re-define the helpers here so we can test them in isolation
    # without triggering the script's top-level execution (which
    # would require a live API token and network).

    function ConvertTo-CsvField {
        param([string]$Value)
        $escaped = $Value -replace '"', '""'
        return "`"$escaped`""
    }

    function Get-MemberRole {
        param([PSCustomObject]$Member)
        if     ($Member.is_admin)     { return "Admin"  }
        elseif ($Member.is_guest)     { return "Guest"  }
        elseif ($Member.is_view_only) { return "Viewer" }
        else                          { return "Member" }
    }

    function Get-MemberStatus {
        param([PSCustomObject]$Member)
        if ($Member.enabled) { return "Active" } else { return "Inactive" }
    }

    function Get-MemberTeams {
        param([PSCustomObject]$Member)
        if (-not $Member.teams) { return "No Teams" }
        $names = @($Member.teams | ForEach-Object { $_.name })
        if ($names.Count -gt 0) { return $names -join "; " } else { return "No Teams" }
    }

    function Get-MemberProducts {
        param([PSCustomObject]$Member)
        if (-not $Member.products) { return "No Products" }
        $names = @($Member.products | ForEach-Object { $_.name })
        if ($names.Count -gt 0) { return $names -join "; " } else { return "No Products" }
    }

    function Get-InvitedBy {
        param([PSCustomObject]$Member)
        if ($Member.invited_by) { return $Member.invited_by.name } else { return "N/A" }
    }

    function Build-CsvRow {
        param(
            [PSCustomObject]$Member,
            [string]$WsName,
            [string]$WsUrl
        )
        $role       = Get-MemberRole     $Member
        $products   = Get-MemberProducts $Member
        $status     = Get-MemberStatus   $Member
        $teams      = Get-MemberTeams    $Member
        $invitedBy  = Get-InvitedBy      $Member
        $joined     = if ($Member.created_at)    { $Member.created_at }    else { "" }
        $lastActive = if ($Member.last_activity) { $Member.last_activity } else { "Never logged in" }

        $fields = @($WsName, $Member.name, $Member.email, $role, $products, $status,
                    $teams, $joined, $invitedBy, $lastActive, "Disabled", $WsUrl)
        return ($fields | ForEach-Object { ConvertTo-CsvField $_ }) -join ","
    }

    # ── Factory: build a test member object ───────────────────
    function New-Member {
        param(
            [string]$Name         = "Alice Smith",
            [string]$Email        = "alice@example.com",
            [bool]  $IsAdmin      = $false,
            [bool]  $IsGuest      = $false,
            [bool]  $IsViewOnly   = $false,
            [bool]  $Enabled      = $true,
            [string]$CreatedAt    = "2024-01-15",
            [string]$LastActivity = "2024-06-01",
            [array] $Teams        = @(@{ name = "Engineering" }),
            [array] $Products     = @(@{ name = "WorkForms" }),
            $InvitedBy            = @{ name = "Bob Jones" }
        )
        return [PSCustomObject]@{
            name          = $Name
            email         = $Email
            is_admin      = $IsAdmin
            is_guest      = $IsGuest
            is_view_only  = $IsViewOnly
            enabled       = $Enabled
            created_at    = $CreatedAt
            last_activity = $LastActivity
            teams         = $Teams
            products      = $Products
            invited_by    = $InvitedBy
        }
    }
}

# ── ConvertTo-CsvField ────────────────────────────────────────
Describe "ConvertTo-CsvField" {
    It "wraps plain text in double-quotes" {
        ConvertTo-CsvField "hello" | Should -Be '"hello"'
    }

    It "escapes embedded double-quotes by doubling them" {
        ConvertTo-CsvField 'say "hi"' | Should -Be '"say ""hi"""'
    }

    It "handles empty string" {
        ConvertTo-CsvField "" | Should -Be '""'
    }

    It "preserves commas inside the quoted field" {
        ConvertTo-CsvField "a,b,c" | Should -Be '"a,b,c"'
    }

    It "preserves newlines inside the quoted field" {
        ConvertTo-CsvField "line1`nline2" | Should -Be "`"line1`nline2`""
    }
}

# ── Get-MemberRole ────────────────────────────────────────────
Describe "Get-MemberRole" {
    It "returns Admin when is_admin is true" {
        $m = New-Member -IsAdmin $true
        Get-MemberRole $m | Should -Be "Admin"
    }

    It "returns Guest when is_guest is true" {
        $m = New-Member -IsGuest $true
        Get-MemberRole $m | Should -Be "Guest"
    }

    It "returns Viewer when is_view_only is true" {
        $m = New-Member -IsViewOnly $true
        Get-MemberRole $m | Should -Be "Viewer"
    }

    It "returns Member for a standard user" {
        $m = New-Member
        Get-MemberRole $m | Should -Be "Member"
    }

    It "Admin takes priority over Guest when both flags are set" {
        $m = New-Member -IsAdmin $true -IsGuest $true
        Get-MemberRole $m | Should -Be "Admin"
    }
}

# ── Get-MemberStatus ──────────────────────────────────────────
Describe "Get-MemberStatus" {
    It "returns Active when enabled is true" {
        $m = New-Member -Enabled $true
        Get-MemberStatus $m | Should -Be "Active"
    }

    It "returns Inactive when enabled is false" {
        $m = New-Member -Enabled $false
        Get-MemberStatus $m | Should -Be "Inactive"
    }
}

# ── Get-MemberTeams ───────────────────────────────────────────
Describe "Get-MemberTeams" {
    It "returns a single team name" {
        $m = New-Member -Teams @(@{ name = "Engineering" })
        Get-MemberTeams $m | Should -Be "Engineering"
    }

    It "joins multiple team names with semicolon-space" {
        $m = New-Member -Teams @(@{ name = "Engineering" }, @{ name = "Design" })
        Get-MemberTeams $m | Should -Be "Engineering; Design"
    }

    It "returns 'No Teams' when teams array is empty" {
        $m = New-Member -Teams @()
        Get-MemberTeams $m | Should -Be "No Teams"
    }

    It "returns 'No Teams' when teams is null" {
        $m = New-Member -Teams $null
        Get-MemberTeams $m | Should -Be "No Teams"
    }
}

# ── Get-MemberProducts ───────────────────────────────────────
Describe "Get-MemberProducts" {
    It "returns a single product name" {
        $m = New-Member -Products @(@{ name = "WorkForms" })
        Get-MemberProducts $m | Should -Be "WorkForms"
    }

    It "joins multiple product names with semicolon-space" {
        $m = New-Member -Products @(@{ name = "WorkForms" }, @{ name = "WorkCanvas" })
        Get-MemberProducts $m | Should -Be "WorkForms; WorkCanvas"
    }

    It "returns 'No Products' when products array is empty" {
        $m = New-Member -Products @()
        Get-MemberProducts $m | Should -Be "No Products"
    }

    It "returns 'No Products' when products is null" {
        $m = New-Member -Products $null
        Get-MemberProducts $m | Should -Be "No Products"
    }
}

# ── Get-InvitedBy ─────────────────────────────────────────────
Describe "Get-InvitedBy" {
    It "returns the inviter's name when invited_by is set" {
        $m = New-Member -InvitedBy @{ name = "Bob Jones" }
        Get-InvitedBy $m | Should -Be "Bob Jones"
    }

    It "returns 'N/A' when invited_by is null" {
        $m = New-Member -InvitedBy $null
        Get-InvitedBy $m | Should -Be "N/A"
    }
}

# ── Build-CsvRow ──────────────────────────────────────────────
Describe "Build-CsvRow" {
    BeforeAll {
        $WsName = "Main Workspace"
        $WsUrl  = "https://coral.monday.com/workspaces/42"
    }

    It "produces exactly 12 quoted comma-separated fields" {
        $m   = New-Member
        $row = Build-CsvRow -Member $m -WsName $WsName -WsUrl $WsUrl
        # Split on commas NOT inside quotes
        $fields = $row | ConvertFrom-Csv -Header (1..12)
        ($fields | Get-Member -MemberType NoteProperty).Count | Should -Be 12
    }

    It "puts workspace name in the first field" {
        $row = Build-CsvRow -Member (New-Member) -WsName $WsName -WsUrl $WsUrl
        $row | Should -Match "^`"Main Workspace`","
    }

    It "puts workspace URL in the last field" {
        $row = Build-CsvRow -Member (New-Member) -WsName $WsName -WsUrl $WsUrl
        $row | Should -Match ",`"https://coral\.monday\.com/workspaces/42`"$"
    }

    It "always sets the 2FA field to Disabled" {
        $m      = New-Member
        $row    = Build-CsvRow -Member $m -WsName $WsName -WsUrl $WsUrl
        $fields = $row -split '","'
        $twofa  = $fields[10] -replace '"', ''
        $twofa | Should -Be "Disabled"
    }

    It "uses empty string for missing created_at" {
        $m = New-Member -CreatedAt ""
        $row = Build-CsvRow -Member $m -WsName $WsName -WsUrl $WsUrl
        $fields = $row -split '","'
        $joined = $fields[7] -replace '"', ''
        $joined | Should -Be ""
    }

    It "uses 'Never logged in' for missing last_activity" {
        $m = [PSCustomObject]@{
            name          = "Test"
            email         = "t@t.com"
            is_admin      = $false
            is_guest      = $false
            is_view_only  = $false
            enabled       = $true
            created_at    = "2024-01-01"
            last_activity = $null
            teams         = @()
            products      = @()
            invited_by    = $null
        }
        $row    = Build-CsvRow -Member $m -WsName $WsName -WsUrl $WsUrl
        $fields = $row -split '","'
        $last   = $fields[9] -replace '"', ''
        $last | Should -Be "Never logged in"
    }

    It "correctly reflects Admin role" {
        $m      = New-Member -IsAdmin $true
        $row    = Build-CsvRow -Member $m -WsName $WsName -WsUrl $WsUrl
        $fields = $row -split '","'
        ($fields[3] -replace '"', '') | Should -Be "Admin"
    }

    It "correctly reflects Inactive status" {
        $m      = New-Member -Enabled $false
        $row    = Build-CsvRow -Member $m -WsName $WsName -WsUrl $WsUrl
        $fields = $row -split '","'
        ($fields[5] -replace '"', '') | Should -Be "Inactive"
    }

    It "puts products in the fifth field" {
        $m      = New-Member -Products @(@{ name = "WorkForms" })
        $row    = Build-CsvRow -Member $m -WsName $WsName -WsUrl $WsUrl
        $fields = $row -split '","'
        ($fields[4] -replace '"', '') | Should -Be "WorkForms"
    }

    It "puts invited-by name in the ninth field" {
        $m      = New-Member -InvitedBy @{ name = "Bob Jones" }
        $row    = Build-CsvRow -Member $m -WsName $WsName -WsUrl $WsUrl
        $fields = $row -split '","'
        ($fields[8] -replace '"', '') | Should -Be "Bob Jones"
    }
}

# ── .env parsing (regex logic) ────────────────────────────────
Describe ".env line parsing" {
    It "matches a valid KEY=VALUE line" {
        $line = "MONDAY_API_TOKEN=abc123"
        $line -match '^\s*([^#][^=]+)=(.*)$' | Should -Be $true
        $Matches[1].Trim() | Should -Be "MONDAY_API_TOKEN"
        $Matches[2].Trim() | Should -Be "abc123"
    }

    It "ignores comment lines starting with #" {
        $line = "# This is a comment"
        $line -match '^\s*([^#][^=]+)=(.*)$' | Should -Be $false
    }

    It "handles values that contain an equals sign" {
        $line = "TOKEN=abc=def"
        $line -match '^\s*([^#][^=]+)=(.*)$' | Should -Be $true
        $Matches[2].Trim() | Should -Be "abc=def"
    }

    It "handles leading whitespace before the key" {
        $line = "  MY_VAR=value"
        $line -match '^\s*([^#][^=]+)=(.*)$' | Should -Be $true
        $Matches[1].Trim() | Should -Be "MY_VAR"
    }
}

# ── Workspace URL construction ────────────────────────────────
Describe "Workspace URL construction" {
    It "builds the correct URL from a workspace ID" {
        $WsId  = "99"
        $WsUrl = "https://coral.monday.com/workspaces/$WsId"
        $WsUrl | Should -Be "https://coral.monday.com/workspaces/99"
    }
}
