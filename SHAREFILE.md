# ShareFile User Report

Companion to the [Monday.com report](README.md), built for **AES-315**. It replicates the manually-exported ShareFile "UserList.xlsx" (Admin → Users → Export) using the ShareFile REST API v3, and adds an `AddedBy` column for **AES-316** (see below). It is currently standalone — not wired into GitHub Actions or the SharePoint upload.

This document (build notes + the Postman testing record) is **AES-317**.

---

## What It Produces

Each run writes three files to `Reports\sharefile\`:

| File | Description |
| ---- | ----------- |
| `sharefile_user_report_YYYY-MM-DD.csv` | Raw data — one row per Employee/Client user |
| `sharefile_user_report_YYYY-MM-DD.html` | Executive dashboard — new users (30 days), company breakdown, inactive users, external/Client users |
| `sharefile_user_report_YYYY-MM-DD.xlsx` | Excel workbook, new-user rows highlighted blue |

### CSV Columns

The first 10 columns match the manual `UserList.xlsx` export exactly; `AddedBy` is new (see AES-316 below).

| Column | Source | Notes |
| ------ | ------ | ----- |
| Email | `Email` / `Username` | |
| FirstName | `FirstName` | |
| LastName | `LastName` | |
| Company | `Company` | Blank is often real (contact has no company on file), not a mapping bug |
| UserType | Derived | `Employee` or `Client`, from which feed the row came from |
| CreationDate | `CreatedDate` / `CreationDate` / `DateCreated` / `Created` | Converted UTC → US Eastern |
| LastLoginDate | `LastAnyLogin` (not `LastLoginDate`, despite the CSV header) | Converted UTC → US Eastern; ShareFile's "never logged in" sentinel (`1900-01-01`) is blanked |
| UserDisabled | `IsDisabled` / `Disabled` / `UserDisabled` / `IsInactive` | `Yes` / `No` |
| SharedAddressBook | — | Hardcoded `"Yes"` — no backing API field exists (see below) |
| SecondaryEmail | `EmailAddresses` (non-primary entry) | Requires a per-user `Users({id})` call |
| **AddedBy** *(new, AES-316)* | `ReferredBy`, resolved to a name/email | Who added the user — see AES-316 section |

---

## Why This Script Exists This Way

### Password grant doesn't work — MFA

ShareFile's OAuth2 **password grant** has no way to satisfy an MFA challenge. Any account with MFA enabled (the norm for this org) fails password grant with `invalid_grant: invalid username or password`, regardless of correct credentials — confirmed by testing it directly (see Postman section below).

The workaround is the **Authorization Code flow**, run once interactively via `scripts\sharefile_get_refresh_token.ps1`:

1. Opens the default browser to ShareFile's login page (MFA happens there, in the real browser).
2. This API app's redirect URI is ShareFile's own built-in landing page (`https://secure.sharefile.com/oauth/oauthcomplete.aspx`), not a URL this script controls — so after login, ShareFile displays the resulting code on screen rather than redirecting back to a local listener. The script has you paste that URL/code back in manually.
3. Exchanges the code for an access token + a long-lived **refresh token**, and writes `SHAREFILE_REFRESH_TOKEN` into `.env`.

After that one-time setup, `sharefile_user_report.ps1` mints new access tokens on every run via the refresh-token grant — no MFA needed again until the refresh token itself is revoked or expires. ShareFile may rotate the refresh token on use; the script detects a changed value in the token response and persists it back to `.env` automatically, so the next run doesn't fail with a stale token.

### No published field list — reverse-engineered via Postman

ShareFile's public API docs don't publish a full `Contact`/`User` field list. Column values are read through a `Get-Field` helper (`scripts\sharefile_user_report.ps1`) that tries several likely property names per column and leaves the cell blank + logs a one-time warning if none match. The actual field names were confirmed by:

- Running the script with `-DumpRawSample`, which writes one raw Employee JSON object and one raw Client JSON object to `Reports\sharefile\debug\`.
- Manually probing the API in Postman (see below).
- Cross-checking the resulting CSV against a real `UserList.xlsx` export from the `coralconnect` tenant.

### Split Employees/Clients feed

`Accounts/Employees` and `Accounts/Clients` are fetched as two separate OData feeds. This split gives `UserType` directly, instead of guessing at a field for it on a single combined feed.

### Two known non-fields

- **SharedAddressBook** has no backing field anywhere in the API — checked the full `Users({id})` object and the `Contacts` feed, neither has it. It was `Yes` for all 34 users in the reference export with zero variance, so the script hardcodes it. Revisit if a real field ever surfaces or a user with a different value turns up.
- **Timezone** isn't exposed anywhere (checked `Accounts/Preferences` — see Postman section). `CreatedDate`/`LastAnyLogin` come back in UTC, but the ShareFile UI/exports show account-local time. The script hardcodes conversion to US Eastern (`ConvertTo-EasternString`, DST-aware), confirmed against 34 sampled timestamps from a real export (50/51 matched exactly across two report runs; the one diff was a genuine login between runs). Revisit if this script is ever pointed at a non-Eastern account.

---

## AES-316: Manual vs. Automated Provisioning (`AddedBy`)

**Goal:** distinguish users added by SCIM/automated provisioning from users added manually, without an explicit API flag for it.

**What was tried, and what came back:**

| Approach | Result |
| -------- | ------ |
| Look for a provisioning-source field on `Accounts/Preferences` (`SCIM`, `Provision`, `Sync`, `Directory`, etc.) | Only found `EnableSync` / `EnableSyncAutoUpdate` — these gate the **desktop sync client**, unrelated to identity provisioning |
| Look for a provisioning-source field on the full `Users({id})` object (`ExternalId`, `Provider`, `SSOId`, `ManagedBy`, `Source`, etc.) | None of these field names exist on this account as of 2026-09 |
| Probe `GET /scim/v2/ServiceProviderConfig` on `{subdomain}.sharefile.com` (the conventional unauthenticated SCIM discovery endpoint) | Returns HTTP 200 but with the ShareFile web app's HTML shell (SPA catch-all), not real SCIM JSON |
| Probe `GET /scim/v2/Users` on `{subdomain}.sharefile.com` | Same — 200, HTML shell, not real SCIM JSON |
| Probe `GET /scim/v2/Users` on `{subdomain}.{apicp}` (the `sf-api.com` API host) | Plain 404 |

**Conclusion:** no accessible SCIM 2.0 endpoint exists under this API app's OAuth token, on either host. ShareFile *does* have a real `AdminSCIM` permission — seen on the account's MasterAdmin (`ryan.rollo@coralconnect.com`)'s `Roles` — gating some SCIM/provisioning feature in the Admin UI. That's the place to check for ground truth on whether SCIM is actually configured for this tenant, since the REST API can't confirm it either way.

**The proxy that was shipped instead:** the script surfaces `ReferredBy` (who added the user) as a new `AddedBy` column, resolved to a name/email via a second `Users({id})` lookup. It's only visible on the full `Users({id})` object — not on the `Accounts/Employees`/`Accounts/Clients` list feed, and OData `$select` silently drops unknown/unavailable field names from that feed rather than erroring, which is how this was confirmed it can't be projected from there either. The account owner's `ReferredBy` comes back as the literal string `"none"` rather than null; the script normalizes both to blank.

**Empirical pattern (as of 2026-09-22):** every current employee except the account owner is referred by `sa_nebulas@coralconnect.com` (a service/admin account); every client is referred by a specific named human employee (a real invite). This makes `AddedBy` a reasonable manual-vs-automated signal in practice — but the script deliberately does **not** hardcode a boolean "is this a service account" rule on top of it. That would mean guessing which accounts count as automation, which breaks the moment a new service account or a new human onboarding-admin shows up. Report readers are expected to interpret the resolved name themselves.

---

## Postman Testing

All ShareFile API exploration was done through the **"ShareFile API"** Postman collection (`postman/collections/sharefile api/`), against the `coralconnect` tenant. This is where the field-name and endpoint findings above were actually verified before being encoded into the script.

### Collection variables (`.resources/definition.yaml`)

| Variable | Purpose |
| -------- | ------- |
| `sharefile_subdomain` | `coralconnect` |
| `sharefile_client_id` / `sharefile_client_secret` | API app credentials (secret) |
| `sharefile_refresh_token` | From the one-time `sharefile_get_refresh_token.ps1` flow (secret) |
| `sharefile_access_token` | Short-lived token minted by "Get Access Token" (secret) |
| `sharefile_api_base` | `https://coralconnect.sf-api.com/sf/v3` |
| `user_id` | Scratch variable for the by-ID requests |

### Requests, in order

1. **Get Access Token** — `POST https://{subdomain}.sharefile.com/oauth/token`, `grant_type=refresh_token`. Confirms the refresh-token exchange works end-to-end and mints `sharefile_access_token` for the rest of the collection. (This is also where the password-grant/MFA failure was originally confirmed, before switching to the refresh-token approach.)
2. **Get Accounts Employees** — `GET {api_base}/Accounts/Employees?$top=100`. Verified this feed's field names and confirmed it returns only employee-type users.
3. **Get Accounts Clients** — `GET {api_base}/Accounts/Clients?$top=100`. Same, for client-type (external) users — confirms the Employees/Clients split is how `UserType` should be derived.
4. **Get User By Id (full fields)** — `GET {api_base}/Users({user_id})`, no `$select`. Returns every property the `User` entity exposes for this account. Used to find `SecondaryEmail`'s backing field (`EmailAddresses`) and, for AES-316, to check for any provisioning-source field before settling on `ReferredBy`.
5. **Get Account Preferences** — `GET {api_base}/Accounts/Preferences`. Checked (AES-316) for SCIM/Provision/Sync/Directory-named flags; only found `EnableSync`/`EnableSyncAutoUpdate` (desktop sync client, unrelated). Also the place that was checked for a timezone field — none exists.
6. **Probe SCIM ServiceProviderConfig** — `GET https://{subdomain}.sharefile.com/scim/v2/ServiceProviderConfig`. AES-316: the conventional unauthenticated SCIM discovery endpoint. Returned the web app's HTML shell, not real SCIM JSON.
7. **Probe SCIM Users** — `GET https://{subdomain}.sharefile.com/scim/v2/Users` (also tried against the `sf-api.com` host by editing the URL). AES-316: `sharefile.com` fell through to the SPA catch-all (200 + HTML); `sf-api.com` returned a plain 404. Neither is a real, reachable SCIM endpoint for this account/API app.

Requests 5–7 carry inline `description` fields in their `.request.yaml` source with the exact dates and conclusions above — check those first if picking this investigation back up, before re-probing.

### Re-running this investigation

If ShareFile's provisioning/SCIM feature gets turned on for this tenant later, or if a new field shows up:

1. Re-run **Get Account Preferences** and **Get User By Id (full fields)** — new fields would show up directly in the JSON body.

2. Re-run the two **Probe SCIM …** requests — a real SCIM setup would return actual SCIM-shaped JSON instead of the HTML shell / 404.

3. Cross-check any Admin UI change against the `AdminSCIM` permission on the account's MasterAdmin.

---

## Setup

```
SHAREFILE_SUBDOMAIN=coralconnect
SHAREFILE_CLIENT_ID=<api-app-client-id>
SHAREFILE_CLIENT_SECRET=<api-app-client-secret>
SHAREFILE_REDIRECT_URI=<redirect-uri-registered-on-the-api-app>
SHAREFILE_REFRESH_TOKEN=<from sharefile_get_refresh_token.ps1>
```

One-time refresh token setup:

```powershell
.\scripts\sharefile_get_refresh_token.ps1
```

## Running

```powershell
.\scripts\sharefile_user_report.ps1

# First run against a real account, or after a schema change —
# writes raw JSON samples to Reports\sharefile\debug\ for field verification:
.\scripts\sharefile_user_report.ps1 -DumpRawSample
```

## API Details

- **Endpoint:** `https://{subdomain}.{apicp}/sf/v3` (the `apicp` host and account subdomain come back in the OAuth token response — not necessarily the same as the login subdomain)
- **Auth:** `Authorization: Bearer {access_token}`
- **Pagination:** OData `$top`/`$skip`, page size 100
- **Rate limiting:** 150ms delay between list pages, 100ms delay per per-user `Users({id})` lookup
