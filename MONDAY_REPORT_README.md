# Monday.com User-Workspace Report — PowerShell

## What This Does

Automates the manual Coral User Report (XLSX) by querying the Monday.com GraphQL API.  
Produces a CSV with one row per workspace member:

| Column | Description |
|--------|-------------|
| Workspace | Workspace name |
| Name | Display name |
| Email | User email |
| User Role | `Admin`, `Member`, `Viewer`, or `Guest` |
| Products | Always `N/A` — not available via the workspace `users_subscribers` API field |
| Status | `Active` or `Inactive` |
| Teams | Semicolon-separated list, or `No Teams` |
| Joined | `created_at` date from the API |
| Invited By | Always `N/A` — not available via the workspace `users_subscribers` API field |
| Last Active | `last_activity` date, or `Never logged in` |
| 2FA | Always `Disabled` (API does not expose this field) |
| Workspace URL | Direct link: `https://coral.monday.com/workspaces/{id}` |

> **Guest/External User note:** Users with `is_guest: true` only exist in the Monday.com universe — they are not provisioned through Entra ID or SCIM. The `User Role` column flags these as `Guest` so they can be identified separately from Entra-managed full members.

---

## Prerequisites

| Requirement | Minimum version | Check |
|-------------|----------------|-------|
| Windows PowerShell | 5.1 | `$PSVersionTable.PSVersion` |
| Internet access to `api.monday.com` | — | — |
| Monday.com API token | — | See below |

No third-party modules are required. The script uses only built-in cmdlets (`Invoke-RestMethod`, `ConvertTo-Json`, `Add-Content`).

---

## Getting Your API Token

1. Log into Monday.com
2. Click your **avatar** (bottom-left corner)
3. Go to **Developers → My Access Tokens**
4. Copy the token — treat it like a password, never commit it to source control

---

## Setup

Create a `.env` file next to the script with your token:

```
MONDAY_API_TOKEN=eyJhbGci...
```

The script loads this file automatically on startup. Comment lines (starting with `#`) and extra whitespace are ignored.

---

## Running the Script

```powershell
# From the directory containing monday_user_report.ps1
.\monday_user_report.ps1
```

If your execution policy blocks unsigned scripts, run once with:

```powershell
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
```

**Output:** `coral_user_report_YYYY-MM-DD.csv` in the current directory.

Console output during a run:

```
Fetching workspaces...
  -> Processing workspace: Main workspace (id: 12345, kind: open)
    Page 1: 39 user(s) returned
    Total members in 'Main workspace': 39
  -> Processing workspace: IT Projects (id: 67890, kind: closed)
    Page 1: 12 member(s) returned
    Total members in 'IT Projects': 12

Removing deleted-member records...
Done! Report saved to: coral_user_report_2026-05-08.csv
   Rows written: 266
   Deleted-member rows removed: 0
```

---

## Filtering the Output

Open the CSV in Excel or run quick PowerShell filters:

```powershell
# All Guest/External users (Monday.com only — not in Entra)
Import-Csv coral_user_report_*.csv | Where-Object { $_.'User Role' -eq 'Guest' }

# All active users in a specific workspace
Import-Csv coral_user_report_*.csv |
    Where-Object { $_.Workspace -eq 'Main Workspace' -and $_.Status -eq 'Active' }

# Users who have never logged in
Import-Csv coral_user_report_*.csv | Where-Object { $_.'Last Active' -eq 'Never logged in' }

# Export guests to a separate file
Import-Csv coral_user_report_*.csv |
    Where-Object { $_.'User Role' -eq 'Guest' } |
    Export-Csv guests_only.csv -NoTypeInformation
```

---

## API Notes

- **Endpoint:** `https://api.monday.com/v2` (single GraphQL endpoint)
- **Auth header:** `Authorization: your_token` — no `Bearer` prefix
- **API version header:** `API-Version: 2024-07` pins to a stable schema version
- **Rate limits:** ~5,000 complexity points/minute on most plans; the script includes a 300ms delay between workspace calls and 150ms between pagination pages
- **Workspace limit:** fetches up to 100 workspaces per run; increase `limit: 100` in the query if needed
- **Pagination:** member and user queries loop through pages of 100 until a page returns fewer than 100 results — no manual limit increase needed
- **Open vs closed workspaces:** the script checks each workspace's `kind` field and uses different queries accordingly (see below)

---

## Key GraphQL Queries

### List all workspaces (including kind)
```graphql
query {
  workspaces(limit: 100) {
    id
    name
    kind
  }
}
```

### Members of a closed workspace (explicit membership)
`users_subscribers` is correct for closed workspaces. Loop `page` from 1 until the result is smaller than `limit`.
```graphql
query {
  workspaces(ids: [15284606]) {
    members: users_subscribers(limit: 100, page: 1) {
      id
      name
      email
      is_admin
      is_guest
      is_view_only
      enabled
      created_at
      last_activity
      teams { name }
    }
  }
}
```

### All account users (used for open workspaces)
Open workspaces (including Main workspace) grant access to all account members implicitly — `users_subscribers` only returns explicit subscribers and misses the bulk of the membership. Use the top-level `users` query instead. `kind: all` ensures guests are included alongside regular members.
```graphql
query {
  users(limit: 100, page: 1, kind: all) {
    id
    name
    email
    is_admin
    is_guest
    is_view_only
    enabled
    created_at
    last_activity
    teams { name }
  }
}
```

> **Pending invitations note:** Users who have been invited but have not yet accepted their invitation appear in the Monday.com admin UI but are not returned by the `users` API endpoint. This is an API limitation — there is no endpoint to list pending invitations.

---

## Running the Tests

Tests use [Pester v5](https://pester.dev/), the standard PowerShell testing framework.

### Install Pester

```powershell
Install-Module Pester -MinimumVersion 5.0 -Force -Scope CurrentUser -SkipPublisherCheck
```

### Run the full test suite

```powershell
Invoke-Pester -Path tests/monday_user_report.Tests.ps1 -Output Detailed
```

Expected result: **37 tests, 0 failures.**

### What the tests cover

| Describe block | Tests | What it validates |
|----------------|------:|-------------------|
| `ConvertTo-CsvField` | 5 | CSV quoting, double-quote escaping, commas and newlines in values |
| `Get-MemberRole` | 5 | Role priority (`Admin > Guest > Viewer > Member`), edge case where both `is_admin` and `is_guest` are set |
| `Get-MemberStatus` | 2 | `enabled` flag maps correctly to `Active` / `Inactive` |
| `Get-MemberTeams` | 4 | Single team, multiple teams joined with `; `, empty array, and `null` value all return correct output |
| `Get-MemberProducts` | 4 | Single product, multiple products joined with `; `, empty array, and `null` value all return correct output |
| `Get-InvitedBy` | 2 | Returns inviter name when set; returns `N/A` when `invited_by` is null |
| `Build-CsvRow` | 10 | Full row has 12 fields in the correct column order; null dates degrade gracefully; Products and Invited By at correct indices |
| `.env` line parsing | 4 | Token loading regex handles comments, `=` in values, and leading whitespace |
| Workspace URL | 1 | Deep-link URL is constructed correctly from a workspace ID |

The test file is self-contained — it re-defines the helper functions in a `BeforeAll` block and does not make any network calls, so it runs without a real API token.

### Unblocking the test file (first run only)

If PowerShell blocks the test file because it came from a network location (e.g. OneDrive sync):

```powershell
Unblock-File -Path tests/monday_user_report.Tests.ps1
```

---

## Scheduling (Optional)

### Windows Task Scheduler

```powershell
$action  = New-ScheduledTaskAction -Execute 'powershell.exe' `
               -Argument '-NonInteractive -File "C:\path\to\monday_user_report.ps1"' `
               -WorkingDirectory 'C:\path\to\'

$trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday -At '08:00'

Register-ScheduledTask -TaskName 'MondayUserReport' -Action $action -Trigger $trigger -RunLevel Highest
```

### GitHub Actions

```yaml
on:
  schedule:
    - cron: '0 8 * * 1'   # Every Monday at 8am UTC
jobs:
  report:
    runs-on: windows-latest
    steps:
      - uses: actions/checkout@v4
      - name: Run report
        env:
          MONDAY_API_TOKEN: ${{ secrets.MONDAY_API_TOKEN }}
        shell: pwsh
        run: .\monday_user_report.ps1
      - uses: actions/upload-artifact@v4
        with:
          name: user-report
          path: coral_user_report_*.csv
```

Store the token as a GitHub secret — never hardcode it in the workflow file.

---

## Test Results

See [TEST_RESULTS.md](TEST_RESULTS.md) for the full breakdown of the last test run, including details on each test case and the bugs that were fixed.
