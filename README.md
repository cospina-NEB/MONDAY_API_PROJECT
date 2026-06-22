# Monday.com Coral User Report

Automates the Coral User Report — a complete mapping of every Monday.com workspace to its members, with role, status, team, and activity data. Replaces a manual XLSX process and runs automatically on the 1st of each month.

---

## What It Produces

Each run writes three files to the `Reports/` folder:

| File | Description |
| ---- | ----------- |
| `coral_user_report_YYYY-MM-DD.csv` | Raw data — one row per workspace member |
| `coral_user_report_YYYY-MM-DD.html` | Executive HTML report with workspace breakdown, inactive users, guest users, and new hire sections |
| `coral_user_report_YYYY-MM-DD.xlsx` | Excel workbook with the same data; new hire rows highlighted blue |

### CSV Columns

| Column | Source | Notes |
| ------ | ------ | ----- |
| Workspace | API | Name of the Monday.com workspace |
| Name | API | User's display name |
| Email | API | User's email address |
| User Role | API | `Admin`, `Member`, `Viewer`, `Guest`, or `Pending` — derived from the `kind` field |
| Status | API | `Active` or `Inactive` (based on `enabled` flag) |
| Teams | API | Semicolon-separated list, or `No Teams` |
| Joined | API | `created_at` date from the API |
| Last Active | API | `last_activity` date, or `Never logged in` |
| Invitation Method | API | `email`, `link`, `api`, or the inviter's display name (resolved via audit logs for `user` invites); `N/A` for the original account owner |
| 2FA | — | Always `Disabled` — not exposed by the API |
| Workspace URL | Derived | `https://coral.monday.com/workspaces/{id}` |

### HTML Report Sections

All tables have sortable column headers — click a header to sort ascending (↑), click again for descending (↓), click a third time to reset (⇅). Each table sorts independently.

- **New Hires (last 30 days)** — users whose `Joined` date is within the past 30 days; filterable by workspace, status, and invitation method
- **Workspace Breakdown** — per-workspace totals for each role and active/inactive counts
- **Inactive Users** — users who have never logged in or whose last activity is older than 30 days
- **Guest / External Users** — all users with `User Role = Guest` (Monday.com-only accounts, not provisioned through Entra ID)

---

## Project Structure

```
monday_api_project/
├── monday_user_report.ps1              # Primary script (PowerShell)
├── monday_user_report.py               # Python port (identical logic)
├── upload_to_sharepoint.ps1            # Uploads reports to SharePoint
├── RunScript.bat                       # Double-click launcher (bypasses execution policy)
├── query.graphql                       # Reference GraphQL queries
├── probe_schema.ps1                    # Utility: introspects the Monday.com API schema
├── .env                                # API tokens — not committed to source control
├── .github/
│   └── workflows/
│       └── monthly_report.yml          # GitHub Actions automation
├── tests/
│   ├── monday_user_report.Tests.ps1    # Pester v5 test suite (PowerShell)
│   └── test_monday_user_report.py      # pytest test suite (Python)
└── Reports/                            # Output files (gitignored)
```

---

## How It Works

### Data Flow

```
.env
  └─ MONDAY_API_TOKEN
           │
           ▼
GraphQL API ──► POST /v2  (API-Version: 2026-07)
           │
           ├─► audit_logs(events: ["monday-user-invite","user-invite"])
           │        └─► build invitee email → inviter name map
           │
           ├─► workspaces(limit: 100)
           │        ├─ kind = "open"   ──► users(limit: 100, page: N, kind: all)
           │        └─ kind = "closed" ──► workspaces(ids:[id]).users_subscribers(limit:100, page:N)
           │
           ▼
Normalize each member into an 11-column row
           │
           ├─► Strip "Deleted member" rows
           ├─► Write CSV   → Reports/coral_user_report_YYYY-MM-DD.csv
           ├─► Generate HTML → Reports/coral_user_report_YYYY-MM-DD.html
           └─► Generate Excel → Reports/coral_user_report_YYYY-MM-DD.xlsx
```

### Why Two Different User Queries

Monday.com has two workspace types:

- **Open workspaces** — all account members have implicit access. `users_subscribers` only returns users who explicitly subscribed, missing most of the membership. The top-level `users(kind: all)` query is used instead.
- **Closed workspaces** — membership is explicit. `users_subscribers` is the correct query.

### Pagination

Both member queries are paginated (page size 100). The script loops pages until a page returns fewer than 100 results. The audit log query uses `has_more_pages` / `next_page_number` cursor-based pagination. A 150ms delay is inserted between pages and a 300ms delay between workspaces to stay within API rate limits (~5,000 complexity points/minute).

### Role Detection

Role is derived from the `kind` field on the `User` type (API version 2026-07):

| `kind` value | Reported role |
| ------------ | ------------- |
| `admin` | Admin |
| `guest` | Guest |
| `view_only` | Viewer |
| `PENDING` | Pending |
| anything else | Member |

### Invitation Method Enrichment

The `invitation_method` field on `User` returns `"email"`, `"link"`, `"api"`, `"user"`, or `null`. When the value is `"user"` (meaning someone manually invited them), the script resolves the inviter's display name from audit logs (`monday-user-invite` / `user-invite` events) and writes the name instead of `"user"`. Falls back to `"user"` if the invitee's email isn't in the audit log. Returns `"N/A"` when `invitation_method` is null (e.g. the original account owner).

---

## Prerequisites

### PowerShell script

| Requirement | Version | Check |
| ----------- | ------- | ----- |
| Windows PowerShell | 5.1+ | `$PSVersionTable.PSVersion` |
| ImportExcel module | 7.0+ | Required for Excel output only |

```powershell
# One-time: install ImportExcel
Install-Module ImportExcel -MinimumVersion 7.0 -Force -Scope CurrentUser -SkipPublisherCheck
```

### Python script

```bash
pip install requests python-dotenv openpyxl truststore
```

---

## Setup

Create a `.env` file in the project root:

```
MONDAY_API_TOKEN=eyJhbG...
```

To get your token: Monday.com → avatar (bottom-left) → **Developers → My Access Tokens**. Treat it like a password — never commit it to source control.

---

## Running

### PowerShell (primary)

```powershell
.\monday_user_report.ps1
```

Or double-click `RunScript.bat` — bypasses execution policy automatically.

If PowerShell blocks the script on first run:
```powershell
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
```

### Python (alternative)

```bash
python monday_user_report.py
```

Both implementations produce identical output.

### Console output during a run

```
Fetching workspaces...
Fetching audit logs...
   Audit log entries indexed: 49
  -> Processing workspace: Main Workspace (id: 12345, kind: open)
    Page 1: 39 user(s) returned
    Total members in 'Main Workspace': 39
  -> Processing workspace: IT Projects (id: 67890, kind: closed)
    Page 1: 12 member(s) returned
    Total members in 'IT Projects': 12
Removing deleted-member records...
Done! Report saved to: Reports\coral_user_report_2026-06-01.csv
   Rows written: 266
   Deleted-member rows removed: 0
   HTML report saved to: Reports\coral_user_report_2026-06-01.html
   Excel report saved to: Reports\coral_user_report_2026-06-01.xlsx
```

---

## Automation

### GitHub Actions (monthly)

The workflow at `.github/workflows/monthly_report.yml` runs automatically on the **1st of each month at 8:00 AM UTC**. It:
1. Checks out the repo
2. Writes `.env` from GitHub Secrets
3. Runs `monday_user_report.ps1`
4. Uploads CSV and HTML to SharePoint via `upload_to_sharepoint.ps1`
5. Archives all report files as workflow artifacts (retained 90 days)

Can also be triggered manually from the GitHub Actions UI via **workflow_dispatch**.

**Required GitHub Secrets:**

| Secret | Description |
| ------ | ----------- |
| `MONDAY_API_TOKEN` | Monday.com API token |
| `SHAREPOINT_TENANT_ID` | Azure AD tenant ID |
| `SHAREPOINT_CLIENT_ID` | App registration client ID |
| `SHAREPOINT_CLIENT_SECRET` | App registration secret |
| `SHAREPOINT_SITE_URL` | `https://nebulasco487.sharepoint.com/sites/sp_softwaredevelopment` |
| `SHAREPOINT_FOLDER` | `Shared Documents/COR/monday_api_project` |

### SharePoint Upload (manual)

```powershell
.\monday_user_report.ps1
.\upload_to_sharepoint.ps1
```

**Azure AD app requirements (one-time setup):**
- App registration with **Microsoft Graph → Application → `Sites.ReadWrite.All`** permission
- Admin consent granted

---

## Testing

### PowerShell — Pester v5 (primary suite)

```powershell
# One-time install
Install-Module Pester -MinimumVersion 5.0 -Force -Scope CurrentUser -SkipPublisherCheck

# Run
Invoke-Pester -Path tests/monday_user_report.Tests.ps1 -Output Detailed
```

If the test file is blocked by Windows (e.g. from OneDrive sync):
```powershell
Unblock-File -Path tests/monday_user_report.Tests.ps1
```

### Python — pytest

```bash
pip install pytest requests python-dotenv
pytest tests/test_monday_user_report.py -v
```

The test suites are self-contained and make no network calls — no API token needed to run them.

### What the tests cover

| Area | What is validated |
| ---- | ----------------- |
| CSV escaping | RFC 4180 quoting; embedded quotes, commas, and newlines don't break row structure |
| Role logic | `kind` field maps correctly to `Admin`, `Guest`, `Viewer`, `Pending`, `Member` |
| Status logic | `enabled` flag maps to `Active` / `Inactive` |
| Teams formatting | Single team, multiple teams joined with `; `, empty and null produce `No Teams` |
| Invitation Method | Returns field value or resolved inviter name; falls back to `N/A` when null |
| Full row assembly | 11 fields in correct column order; null dates degrade gracefully |
| `.env` parsing | Token loading regex handles comments, `=` in values, leading whitespace |
| Workspace URL | Deep-link URL built correctly from workspace ID |

---

## API Details

- **Endpoint:** `https://api.monday.com/v2` (single GraphQL endpoint)
- **Auth:** `Authorization: <token>` header — no `Bearer` prefix
- **Version:** `API-Version: 2026-07` (pinned; required for `kind` field and `invitation_method`)
- **Rate limits:** ~5,000 complexity points/minute — built-in delays handle this

### Known API Limitations

| Field | Status |
| ----- | ------ |
| `2FA` | Not exposed by the API — always reported as `Disabled` |
| Pending invitations | Users who have been invited but not yet accepted do not appear in any `users` API query |

---

## Quick Filters (PowerShell)

```powershell
$data = Import-Csv Reports\coral_user_report_*.csv | Select-Object -Last 1

# Guest / external users (Monday.com only — not in Entra ID)
$data | Where-Object { $_.'User Role' -eq 'Guest' }

# Users who have never logged in
$data | Where-Object { $_.'Last Active' -eq 'Never logged in' }

# All active users in a specific workspace
$data | Where-Object { $_.Workspace -eq 'Main Workspace' -and $_.Status -eq 'Active' }

# New hires (joined in last 30 days)
$cutoff = (Get-Date).AddDays(-30)
$data | Where-Object { $_.Joined -and [datetime]::Parse($_.Joined) -gt $cutoff }

# Users invited by a specific person
$data | Where-Object { $_.'Invitation Method' -eq 'Jane Smith' }
```
