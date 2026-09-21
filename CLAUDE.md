# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What This Project Does

Automates generation of the "Coral User Report" — a CSV mapping of Monday.com workspaces to their members with role, status, team, and activity data. It replaces a manual XLSX process and outputs three files per run:
- `coral_user_report_YYYY-MM-DD.csv` — raw data
- `coral_user_report_YYYY-MM-DD.html` — executive HTML report with new-hire highlighting and workspace filter
- `coral_user_report_YYYY-MM-DD.xlsx` — Excel workbook with new-hire rows highlighted blue (requires ImportExcel module)

## Running the Script

```powershell
.\scripts\monday_user_report.ps1
```

Or double-click `RunScript.bat` (bypasses execution policy automatically).

Requires a `.env` file in the project root:
```
MONDAY_API_TOKEN=<your_token>
```

The Excel output requires the ImportExcel module (one-time install):
```powershell
Install-Module ImportExcel -MinimumVersion 7.0 -Force -Scope CurrentUser -SkipPublisherCheck
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

The project is a single PowerShell script (`scripts/monday_user_report.ps1`) with no build step or external runtime dependencies.

**Data flow:**
1. Parse `.env` → extract `MONDAY_API_TOKEN`
2. Query Monday.com GraphQL API for all workspaces (up to 100)
3. For each workspace, fetch members with different queries depending on workspace kind:
   - **Open workspaces** → top-level `users` query (implicit membership)
   - **Closed workspaces** → `users_subscribers` query (explicit membership)
4. Paginate results (page size 100, 150ms delay between pages, 300ms between workspaces)
5. Normalize each member into a CSV row and append to output file
6. Strip any "Deleted member" rows from the final file

**Key functions in `scripts/monday_user_report.ps1`:**
- `Invoke-GQL` — HTTP layer; sends POST to `https://api.monday.com/v2` with auth headers and pinned API version `2026-07`
- `ConvertTo-CsvField` — RFC 4180 CSV escaping (wraps in quotes, doubles embedded quotes)
- `Build-CsvRow` — main data pipeline; calls role/status/teams getters and assembles the 11-column row
- Role/Status/Teams getter functions — normalize raw API values to report-friendly strings

**CSV output columns (in order):** Workspace, Name, Email, User Role, Status, Teams, Joined, Last Active, Invitation Method, 2FA, Workspace URL

**Known API limitations (hardcoded placeholders):**
- `2FA` is always `"Disabled"` — not exposed by the API
- `Products` is always `"N/A"` — not available via `users_subscribers`
- `Invitation Method` — fetched from the `invitation_method` field on the `User` GraphQL type (e.g. `"email"`, `"link"`, `"api"`); returns `"N/A"` when the API returns null (e.g. the original account owner)
- Pending invitations are not visible via the API

## Automation & SharePoint Upload

The report runs automatically on the **1st of each month at 8:00 AM UTC** via GitHub Actions (`.github/workflows/monthly_report.yml`). After generating the CSV and HTML, it uploads both files to SharePoint using `scripts/upload_to_sharepoint.ps1`.

**To upload manually (local run):**
```powershell
.\scripts\monday_user_report.ps1
.\scripts\upload_to_sharepoint.ps1
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
- API version header: `2026-07` (pinned for schema stability; required for `invitation_method` and the `kind` role field)
- Rate limits: ~5,000 complexity points/minute — the built-in delays handle this
- Reference queries: `query.graphql`

## ShareFile User Report (AES-315)

`scripts/sharefile_user_report.ps1` is a ShareFile counterpart to the Monday.com script, built for AES-315. It replicates the manually-exported ShareFile "UserList.xlsx" (Admin > Users > Export) via the ShareFile REST API v3, and produces the same three-file bundle (CSV/HTML/XLSX) as the Monday report. It is currently standalone — not wired into GitHub Actions or the SharePoint upload.

**One-time setup — obtaining a refresh token:**

ShareFile's OAuth2 password grant has no way to satisfy an MFA challenge, so any account with MFA enabled (the norm for this org) always fails password grant with `invalid_grant: invalid username or password`, regardless of correct credentials. Instead, this script authenticates via the Authorization Code flow's refresh token, obtained once interactively:

```powershell
.\scripts\sharefile_get_refresh_token.ps1
```

This opens a browser to ShareFile's login (MFA happens there), listens locally for the OAuth redirect, exchanges the code for tokens, and writes `SHAREFILE_REFRESH_TOKEN` into `.env`. It requires `SHAREFILE_SUBDOMAIN`, `SHAREFILE_CLIENT_ID`, `SHAREFILE_CLIENT_SECRET`, and `SHAREFILE_REDIRECT_URI` already set in `.env`, and the redirect URI must exactly match what's registered on the API app in ShareFile's Admin → API/App Management.

**Running:**
```powershell
.\scripts\sharefile_user_report.ps1
```

Requires these `.env` variables:
```
SHAREFILE_SUBDOMAIN=<your-subdomain>
SHAREFILE_CLIENT_ID=<api-app-client-id>
SHAREFILE_CLIENT_SECRET=<api-app-client-secret>
SHAREFILE_REDIRECT_URI=<redirect-uri-registered-on-the-api-app>
SHAREFILE_REFRESH_TOKEN=<from sharefile_get_refresh_token.ps1>
```

**Data flow:**
1. OAuth2 refresh-token grant against `https://{subdomain}.sharefile.com/oauth/token` → access token + account subdomain/apicp (ShareFile may rotate the refresh token on use; the script persists the new one back to `.env` automatically)
2. Fetch `Accounts/Employees` and `Accounts/Clients` from `https://{subdomain}.{apicp}/sf/v3` (this split gives us the `UserType` column directly, instead of guessing at a field for it)
3. Paginate via OData `$top`/`$skip` (page size 100)
4. Normalize into the same 10 columns as the manual export: Email, FirstName, LastName, Company, UserType, CreationDate, LastLoginDate, UserDisabled, SharedAddressBook, SecondaryEmail
5. Fetch each user's secondary email individually via `Users({id})?$select=Id,EmailAddresses` — the `EmailAddresses` array (with `IsPrimary` flags) only appears on the full single-item `User` object, not on the `Accounts/Employees`/`Accounts/Clients` list feed, and OData `$select` can't project it from that feed either (confirmed: unknown/unavailable `$select` names are silently dropped, not errored). This costs one extra API call per user.
6. Generate HTML dashboard (new users in last 30 days, company breakdown, inactive users, external/Client users) and XLSX (new-user rows highlighted blue), same conventions as the Monday report

**Field mapping:** ShareFile's public API docs don't publish a full `Contact`/`User` field list, so column values are read via a `Get-Field` helper that tries several likely property names per column and leaves the cell blank + logs a warning if none match. Verified against a real account (`-DumpRawSample`, plus manual API probing) against the `coralconnect` tenant and cross-checked against a real `UserList.xlsx` export:
- `Email`, `FirstName`, `LastName`, `Company`, `CreatedDate`, `IsDisabled` map directly from the list feed.
- Last-login is exposed as `LastAnyLogin`, not `LastLoginDate`.
- `SecondaryEmail` = the non-primary entry in `EmailAddresses` on the full `Users({id})` object (see step 5 above).
- `SharedAddressBook` has no backing field anywhere in the API (checked the full `Users({id})` property list and the `Contacts` feed) — it was `Yes` for all users in the reference export with zero variance, so the script hardcodes `"Yes"` rather than pretending it's a real per-user lookup. If a real field ever surfaces, or a user with a different value turns up, this needs revisiting.
- If a `Company` warning appears, it's typically real data (some contacts have no company on file), not a mapping issue — worth spot-checking the CSV before assuming a bug.

**Timezone:** `CreatedDate`/`LastAnyLogin` come back from the API in UTC, but ShareFile's own UI/exports show account-local time. There's no timezone field exposed via the API (checked `AccountPreferences`), so the script hardcodes conversion to US Eastern (`ConvertTo-EasternString`, DST-aware via `TimeZoneInfo`) — confirmed empirically against a real `UserList.xlsx` export (50/51 sampled timestamps matched exactly after conversion; the one remaining diff was a genuine login that happened between the two report runs). If this script is ever pointed at a non-Eastern ShareFile account, this hardcode needs revisiting.
