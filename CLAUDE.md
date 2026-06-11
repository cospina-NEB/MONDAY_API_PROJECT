# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What This Project Does

Automates generation of the "Coral User Report" — a CSV mapping of Monday.com workspaces to their members with role, status, team, and activity data. It replaces a manual XLSX process and outputs `coral_user_report_YYYY-MM-DD.csv`.

## Running the Script

```powershell
.\monday_user_report.ps1
```

Or double-click `RunScript.bat` (bypasses execution policy automatically).

Requires a `.env` file in the project root:
```
MONDAY_API_TOKEN=<your_token>
```

## Testing

**PowerShell (Pester v5) — primary test suite:**
```powershell
# One-time install
Install-Module Pester -MinimumVersion 5.0 -Force -Scope CurrentUser -SkipPublisherCheck

# Run tests
Invoke-Pester -Path tests/monday_user_report.Tests.ps1 -Output Detailed
```

**Python (pytest) — alternative suite:**
```bash
pip install pytest requests python-dotenv
pytest tests/test_monday_user_report.py -v
```

## Architecture

The project is a single PowerShell script (`monday_user_report.ps1`) with no build step or external runtime dependencies.

**Data flow:**
1. Parse `.env` → extract `MONDAY_API_TOKEN`
2. Query Monday.com GraphQL API for all workspaces (up to 100)
3. For each workspace, fetch members with different queries depending on workspace kind:
   - **Open workspaces** → top-level `users` query (implicit membership)
   - **Closed workspaces** → `users_subscribers` query (explicit membership)
4. Paginate results (page size 100, 150ms delay between pages, 300ms between workspaces)
5. Normalize each member into a CSV row and append to output file
6. Strip any "Deleted member" rows from the final file

**Key functions in `monday_user_report.ps1`:**
- `Invoke-GQL` — HTTP layer; sends POST to `https://api.monday.com/v2` with auth headers and pinned API version `2024-07`
- `ConvertTo-CsvField` — RFC 4180 CSV escaping (wraps in quotes, doubles embedded quotes)
- `Build-CsvRow` — main data pipeline; calls role/status/teams getters and assembles the 10-column row
- Role/Status/Teams getter functions — normalize raw API values to report-friendly strings

**CSV output columns (in order):** Workspace, Name, Email, User Role, Status, Teams, Joined, Last Active, 2FA, Workspace URL

**Known API limitations (hardcoded placeholders):**
- `2FA` is always `"Disabled"` — not exposed by the API
- `Products` is always `"N/A"` — not available via `users_subscribers`
- Pending invitations are not visible via the API

## Automation & SharePoint Upload

The report runs automatically on the **1st of each month at 8:00 AM UTC** via GitHub Actions (`.github/workflows/monthly_report.yml`). After generating the CSV and HTML, it uploads both files to SharePoint using `upload_to_sharepoint.ps1`.

**To upload manually (local run):**
```powershell
.\monday_user_report.ps1
.\upload_to_sharepoint.ps1
```

**Required `.env` variables for SharePoint:**
```
SHAREPOINT_TENANT_ID=<azure-ad-tenant-id>
SHAREPOINT_CLIENT_ID=<app-registration-client-id>
SHAREPOINT_CLIENT_SECRET=<app-registration-secret>
SHAREPOINT_SITE_URL=https://nebulasco487.sharepoint.com/sites/sp_softwaredevelopment
SHAREPOINT_FOLDER=Shared Documents/COR/monday_api_project
```

**Required GitHub Secrets** (set in repo → Settings → Secrets → Actions):
`MONDAY_API_TOKEN`, `SHAREPOINT_TENANT_ID`, `SHAREPOINT_CLIENT_ID`, `SHAREPOINT_CLIENT_SECRET`, `SHAREPOINT_SITE_URL`, `SHAREPOINT_FOLDER`

**Azure AD app requirements** (one-time setup):
- App registration with **Microsoft Graph → Application → `Sites.ReadWrite.All`** permission
- Admin consent granted

The workflow can also be triggered manually from the GitHub Actions UI via `workflow_dispatch`.

## API Details

- Endpoint: `https://api.monday.com/v2`
- API version header: `2024-07` (pinned for schema stability)
- Rate limits: ~5,000 complexity points/minute — the built-in delays handle this
- Reference queries: `query.graphql`
