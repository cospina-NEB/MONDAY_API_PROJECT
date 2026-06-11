# Test Results — `monday_user_report.ps1`

**Date:** 2026-05-07  
**Test runner:** Pester v5.7.1  
**PowerShell:** Windows PowerShell 5.1  
**Test file:** `tests/monday_user_report.Tests.ps1`

---

## Summary

| Result | Count |
|--------|------:|
| Passed | 29 |
| Failed | 0 |
| Skipped | 0 |
| **Total** | **29** |

---

## Test Cases by Describe Block

### `ConvertTo-CsvField` — CSV output escaping (5 tests)

| Test | Status |
|------|--------|
| Wraps plain text in double-quotes | PASS |
| Escapes embedded double-quotes by doubling them | PASS |
| Handles empty string | PASS |
| Preserves commas inside the quoted field | PASS |
| Preserves newlines inside the quoted field | PASS |

**What this covers:** The CSV serializer correctly quotes every field and doubles any embedded `"` characters per RFC 4180. Commas and newlines inside a field value do not break the row structure.

---

### `Get-MemberRole` — User role classification (5 tests)

| Test | Status |
|------|--------|
| Returns `Admin` when `is_admin` is true | PASS |
| Returns `Guest` when `is_guest` is true | PASS |
| Returns `Viewer` when `is_view_only` is true | PASS |
| Returns `Member` for a standard user | PASS |
| `Admin` takes priority over `Guest` when both flags are set | PASS |

**What this covers:** The role hierarchy (`Admin > Guest > Viewer > Member`) is evaluated in the correct order. This is the field that distinguishes Monday.com-only guest/external users from Entra-provisioned full members.

---

### `Get-MemberStatus` — Active/Inactive flag (2 tests)

| Test | Status |
|------|--------|
| Returns `Active` when `enabled` is true | PASS |
| Returns `Inactive` when `enabled` is false | PASS |

**What this covers:** Suspended or deprovisioned accounts are correctly labelled `Inactive` regardless of role.

---

### `Get-MemberTeams` — Team membership formatting (4 tests)

| Test | Status |
|------|--------|
| Returns a single team name | PASS |
| Joins multiple team names with `; ` separator | PASS |
| Returns `No Teams` when teams array is empty | PASS |
| Returns `No Teams` when teams value is null | PASS |

**What this covers:** Team names are joined into a single cell-safe string. The null guard ensures users with no team assignment (common for external guests) produce `No Teams` rather than a blank or pipeline error.

---

### `Build-CsvRow` — Full row assembly (7 tests)

| Test | Status |
|------|--------|
| Produces exactly 10 quoted comma-separated fields | PASS |
| Puts workspace name in the first field | PASS |
| Puts workspace URL in the last field | PASS |
| Always sets the 2FA field to `Disabled` | PASS |
| Uses empty string for missing `created_at` | PASS |
| Uses `Never logged in` for missing `last_activity` | PASS |
| Correctly reflects `Admin` role | PASS |
| Correctly reflects `Inactive` status | PASS |

**What this covers:** End-to-end column ordering matches the expected CSV schema (`Workspace, Name, Email, User Role, Status, Teams, Joined, Last Active, 2FA, Workspace URL`). Null date fields degrade gracefully instead of emitting blank cells with no label.

---

### `.env` line parsing — Token loading (4 tests)

| Test | Status |
|------|--------|
| Matches a valid `KEY=VALUE` line | PASS |
| Ignores comment lines starting with `#` | PASS |
| Handles values that contain an `=` sign (e.g. base64 tokens) | PASS |
| Handles leading whitespace before the key | PASS |

**What this covers:** The `.env` regex correctly isolates keys and values for common real-world token formats, including base64-encoded API tokens that contain `=` padding.

---

### Workspace URL construction (1 test)

| Test | Status |
|------|--------|
| Builds the correct URL from a workspace ID | PASS |

**What this covers:** The workspace deep-link (`https://coral.monday.com/workspaces/{id}`) is assembled correctly and included in the final CSV column.

---

## Bugs Fixed Prior to This Run

| File | Bug | Fix Applied |
|------|-----|-------------|
| `monday_user_report.ps1` | Teams null guard was positioned before `$TeamNames` was computed and then immediately overwritten, making it a no-op | Replaced with `if ($Member.teams)` pre-check + `@()` wrapper + `.Count -gt 0` comparison |
| `tests/monday_user_report.Tests.ps1` | Newline test assertion used `\"` (not a valid PS escape); string parsed as a single `\` character | Changed to backtick-quote: `` "`"line1`nline2`"" `` |
| `tests/monday_user_report.Tests.ps1` | `Get-MemberTeams` BeforeAll used `@($null)` which has Count 1, bypassing the empty guard | Added early-exit `if (-not $Member.teams)` before the pipeline |
