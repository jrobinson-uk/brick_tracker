# JR's Bricks Profit Tracker — Claude Code Context

## Project Overview

**Goal:** A Google Sheets + Apps Script tool that automatically pulls BrickLink store data via API and gives a clear picture of true profitability.

**Owner:** legojames / JR's Bricks, Cambridgeshire
**Stack:** Google Sheets · Google Apps Script · BrickLink API v3 · OAuth 1.0a
**Distribution:** Shareable Google Sheet — sellers make a copy and enter their own API keys
**Repo:** GitHub synced to Google Apps Script via `clasp` and GitHub Actions

---

## Repository Structure

```
brick_tracker/
├── Code.js                    # All Apps Script logic (monolithic for now — Phase 3.5 will split)
├── Sidebar.html               # Setup sidebar HTML/CSS/JS for credential entry
├── appsscript.json            # Apps Script project manifest (timezone, runtime V8)
├── .github/workflows/
│   └── clasp-push.yml         # Manual workflow: pushes a branch to Apps Script via clasp
├── .gitignore                 # Ignores node_modules/, .clasprc.json, .clasp.json
└── jrs-bricks-tracker.code-workspace  # VS Code workspace file
```

### Planned module structure (Phase 3.5 target)
After the upcoming refactor, `Code.js` should be split into:
- `Auth.gs` — OAuth 1.0a signing only
- `Orders.gs` — order sync logic
- `UI.gs` — menu, sidebars, status bar
- `Config.gs` — settings read/write via PropertiesService
- `Utils.gs` — shared helpers (date formatting, sheet utilities)
- `Debug.gs` — debug functions, log writer, flush/clear

---

## Development Workflow

### Branching
- Each phase gets its own feature branch.
- Merge to `master` only on phase completion.
- Current working branch convention: `claude/<description>-<hash>`.

### Syncing code to Google Apps Script
The GitHub Action at `.github/workflows/clasp-push.yml` is triggered **manually** (`workflow_dispatch`). It:
1. Checks out the specified branch.
2. Installs `clasp` globally.
3. Writes `~/.clasprc.json` from the `CLASPRC_JSON` repo secret.
4. Writes `.clasp.json` from the `CLASP_JSON` repo secret (never commit these files).
5. Runs `clasp push --force`.

To deploy a branch: trigger the workflow in GitHub Actions and select the branch.

### Local clasp usage
`.clasprc.json` and `.clasp.json` are gitignored. If working locally, obtain them from the project owner.

---

## Code Architecture

### Spreadsheet tabs
| Tab | Purpose |
|-----|---------|
| `Settings` | Status bar (row 10), API call log (row 13), credential setup labels |
| `Orders` | Synced order data; columns 7, 12, 13 are manually editable |

### Key constants (`Code.js`)
```js
BRICKLINK_API_BASE_URL = 'https://api.bricklink.com/api/store/v1/'
OAUTH_SIGNATURE_METHOD = 'HMAC-SHA1'
OAUTH_VERSION = '1.0'
SETTINGS_TAB = 'Settings'
ORDERS_TAB = 'Orders'
ORDERS_HEADERS = ['Order ID','Date','Buyer','Items','Lots','Shipping Charged',
                  'Shipping Actual','Grand Total','Currency','Payment Method',
                  'Status','Refund','Notes']
MANUAL_COLS = [7, 12, 13]   // Shipping Actual, Refund, Notes (1-indexed)
MANUAL_COLOUR = '#FFF9C4'   // light yellow background for manual-entry cells
HEADER_COLOUR = '#f3f3f3'
```

### Credential storage
Credentials are **never** written to the sheet permanently. They are stored via `PropertiesService.getUserProperties()` using these keys:
- `BL_CONSUMER_KEY`
- `BL_CONSUMER_SECRET`
- `BL_ACCESS_TOKEN`
- `BL_ACCESS_TOKEN_SECRET`
- `BL_STORE_USERNAME`

The Settings sidebar (`Sidebar.html`) calls `saveCredentialsFromSidebar(creds)` and `clearCredentials()` via `google.script.run`.

### API request pattern
All BrickLink calls go through `bricklinkRequest(endpoint, method, queryParams, bodyParams)`:
1. Reads credentials from PropertiesService.
2. Builds OAuth 1.0a signature (nonce, timestamp, HMAC-SHA1).
3. Constructs `Authorization` header via `buildAuthorizationHeader()`.
4. Calls `UrlFetchApp.fetch()` with `muteHttpExceptions: true`.
5. Logs the call via `trackApiCall_()` and flushes to the Settings tab via `flushApiCallLog()`.
6. Returns parsed JSON.

### Status bar
`logStatus(message)` writes to `Settings!A10:B10`. Call it after every significant operation.

### API call tracking
- `resetApiCallLog()` — call before a sync function starts.
- `trackApiCall_(endpoint)` — called internally by `bricklinkRequest`.
- `flushApiCallLog(functionName)` — writes a summary to `Settings!B13` and to `Logger`; call at the end of a sync function.

### Sheet protection
`setupOrdersTab()` protects the Orders sheet, leaving only `MANUAL_COLS` columns editable. Recreate protection by calling `setupOrdersTab()` again if columns are added.

---

## BrickLink API v3 — Key Reference

**Base URL:** `https://api.bricklink.com/api/store/v1`
**Auth:** OAuth 1.0a — HMAC-SHA1, sent in `Authorization` header
**Encoding:** UTF-8; dates ISO 8601 (`yyyy-MM-dd'T'HH:mm:ss.SSSZ`)
**Financial values:** 4 decimal places, rounded up

### Order fields relevant to this project
| Field | Notes |
|-------|-------|
| `order_id` | Integer — unique order ID |
| `date_ordered` | Timestamp |
| `buyer_name` | String |
| `total_count` | Total items |
| `unique_count` | Unique lots |
| `cost.shipping` | Postage charged to buyer — **currently not returning from API (open bug)** |
| `cost.grand_total` | Order total |
| `cost.currency_code` | ISO 4217 |
| `payment.method` | Payment method string |
| `status` | Order status string |

### Known issue: `cost.shipping`
`cost.shipping` (postage charged to buyer) is not returned by the BrickLink API in the current order sync. This is the primary investigation target for Phase 3.5. The field maps to the "Shipping Charged" column (col 6) in the Orders tab.

### Common endpoints
```
GET  /orders                    # List orders (params: direction=in|out)
GET  /orders/{order_id}         # Single order detail
GET  /inventories               # List all inventory lots
GET  /inventories/{id}          # Single inventory lot
POST /inventories               # Create lot
PUT  /inventories/{id}          # Update lot (quantity uses +/- prefix)
DELETE /inventories/{id}        # Delete lot
GET  /items/{type}/{no}/price   # Price guide
GET  /colors                    # All colours
GET  /categories                # All categories
GET  /settings/shipping_methods # Your store shipping methods
GET  /notifications             # Unread push notifications
```

---

## Phase Status

| # | Phase | Status |
|---|-------|--------|
| 1 | Foundation & API connection | ✓ Complete |
| 2 | Custom menu & UX shell | ✓ Complete |
| 3 | Order sync | ⚠ Partial — `cost.shipping` not returning |
| 3.5 | Refactor & debug infrastructure | **Up next** |
| 4 | Fee & net revenue calculation | Not started |
| 5 | Purchase cost logger | Not started |
| 6 | Inventory sync | Not started |
| 7 | Profit attribution engine | Not started |
| 8 | Dashboard | Not started |
| 9 | Real-world testing | Not started |
| 10 | Portability & polish | Not started |
| 11 | Launch | Not started |

---

## Phase 3.5 Objectives (Current Priority)

1. **Investigate `cost.shipping`** — why it isn't returned by the API; consider fetching individual order detail (`GET /orders/{order_id}`) which may include shipping breakdown.
2. **Split `Code.js` into modules** — confirm module boundaries against the live file before splitting; the planned module list is above.
3. **Add Debug tab** — timestamped log entries, API call traces, error capture, manual flush/clear button.
4. **Add `Debug.gs`** — debug-specific functions and log writer.
5. **Done when:** Code is modular, Debug tab works, postage issue resolved or diagnosed, all existing functionality intact.

---

## Conventions & Guardrails

- **No personal data hardcoded** — all credentials via PropertiesService, not the sheet.
- **Portable by design** — any seller should be able to copy the sheet, run setup, and use it.
- **API-first** — no manual CSV pasting unless the API genuinely cannot provide the data.
- **Test before advancing** — each phase has a "done when" condition; don't move to the next phase until it's met.
- **Phase-per-chat** — each phase gets its own Claude Code session; attach this document to every new session.
- **Phase-per-branch** — feature branches merge to `master` only on phase completion.
- **Menu stubs** — functions for future phases (`syncInventory`, `logPurchase`, `viewDashboard`) exist as stubs with placeholder alerts; fill them in during their respective phases, don't touch them earlier.
- **Sheet protection** — the Orders tab is protected; manual-entry columns are unprotected. Rebuild protection via `setupOrdersTab()` if the schema changes.
- **No editing `.clasprc.json` or `.clasp.json`** — these are gitignored secrets; manage them via GitHub Actions secrets (`CLASPRC_JSON`, `CLASP_JSON`).

---

## Working Process

- At the end of each phase chat: mark items completed / partial / deferred.
- Roadmap document (`JRs_Bricks_ClaudeCode_Context_v*.docx`) is updated at the end of each phase and re-attached to the next session.
- This `CLAUDE.md` should be updated whenever the phase status or code structure changes materially.

---

## Collaboration Style

This is a **collaborative** project. Claude may write most or all of the code, but the owner (James) should be consulted on approach before significant work begins.

### How decisions work
- **High-level design**: Always discuss before implementing. If a task involves a new artefact, a meaningful structural change, or more than a few lines of non-trivial code, propose the approach and wait for agreement before writing it.
- **Low-level coding decisions**: These can be made autonomously — naming, implementation detail, minor refactors within an agreed design don't need sign-off.
- **When in doubt**: Err on the side of a brief "here's what I'm thinking, does that sound right?" rather than silently proceeding.

### The testing loop
1. Claude proposes and explains the approach.
2. James agrees (or redirects).
3. Claude writes the code.
4. James deploys and tests it.
5. James reports the result back to Claude.
6. Claude helps interpret the result — whether it's a success, an unexpected behaviour, or a failure — and proposes next steps.
