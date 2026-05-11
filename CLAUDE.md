# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Running locally

```bash
pip install -r requirements.txt
python app.py          # starts on http://localhost:5000
```

Requires a `.env` file (gitignored) with:
```
SUPABASE_URL=...
SUPABASE_KEY=...   # anon public key
```

The app runs without Supabase — Steps 1–2 work fully offline. Steps 3–4 degrade gracefully (driver/assignment API calls return 503 but the workflow still navigates).

**Production:** Railway auto-deploys from `main` via `Procfile` (`gunicorn app:app`). Push to `main` = deploy.

## Architecture

Single-file Flask backend (`app.py`) + four Jinja templates. The server is **stateless between requests** — all in-progress workflow data travels via the browser's `sessionStorage`. Only the driver roster and terminal/city assignments hit Supabase for persistence.

### Four-step workflow

| Step | Route | Guards | Writes to session |
|------|-------|--------|-------------------|
| 1 | `/` | — | `atm_terminals` |
| 2 | `/step2` | `atm_terminals` | `atm_cash_amounts` |
| 3 | `/step3` | `atm_terminals`, `atm_cash_amounts` | `atm_drivers_today`, `atm_terminal_assignments` |
| 4 | `/step4` | all three above | — |

`static/utils.js` defines `STORAGE_KEYS` and shared helpers (`guardStep`, `loadFromSession`, `saveToSession`, `formatCurrency`, `urgencyClass`, `toTitleCase`, `attachSortHeaders`, `triggerBase64Download`).

### Key backend logic

- **`/process` (POST):** Parses two `.xls` reports (xlrd) + `.xlsm` template (openpyxl with `keep_vba=True`), returns JSON containing `excel_b64` (base64 XLSM for client-side download) and `terminals` array.
- **`update_template()`** returns `(BytesIO, all_rows, tulsa_rows)`. Tulsa filter: only cities exactly `tulsa` or `tulsa_jd` (case-insensitive) go to the Tulsa tab — `West Tulsa` and `North Tulsa` stay in the main sheet.
- **`/generate-pdf` (POST):** Receives `{drivers: [{name, terminals: [...]}]}`, builds an A4 PDF via fpdf2. Column widths sum to exactly 180mm (A4 210mm − 15mm margins each side). Uses Helvetica (Latin-1 only — avoid Unicode characters like `•`).

### Supabase tables

| Table | Key | Purpose |
|-------|-----|---------|
| `drivers` | `id` (UUID) | Persistent driver roster |
| `terminal_assignments` | `terminal_id` (TEXT) | Per-terminal driver assignment |
| `terminal_status` | `terminal_id` (TEXT) | active / inactive / seasonal |
| `city_assignments` | `city` (TEXT) | Legacy — no longer used by frontend |

All tables use RLS with a permissive `allow_all` policy (internal tool, no user auth). The `_require_db()` helper returns a 503 tuple if Supabase is unconfigured.

### City / terminal key normalization

Cities are stored and compared as `city.strip().lower()` (`city_key` in the terminal JSON). Display uses `toTitleCase()` in JS. This is important for Supabase lookups and session state matching.

## Deploying code changes

The local `~/Documents` path is inaccessible from the shell (macOS TCC). All edits are made to `/tmp/atm_push/` and pushed to GitHub via the REST API:

```bash
# Create blob
b64=$(base64 -i <file> | tr -d '\n')
BLOB_SHA=$(echo "{\"encoding\":\"base64\",\"content\":\"$b64\"}" | \
  gh api repos/ghassan-abunada/atm-cash-pos/git/blobs --input - --jq '.sha')

# Get current state
PARENT_SHA=$(gh api repos/ghassan-abunada/atm-cash-pos/git/ref/heads/main --jq '.object.sha')
PARENT_TREE=$(gh api "repos/ghassan-abunada/atm-cash-pos/git/commits/$PARENT_SHA" --jq '.tree.sha')

# Create tree → commit → update ref
NEW_TREE=$(echo "{\"base_tree\":\"$PARENT_TREE\",\"tree\":[{\"path\":\"<repo-path>\",\"mode\":\"100644\",\"type\":\"blob\",\"sha\":\"$BLOB_SHA\"}]}" | \
  gh api repos/ghassan-abunada/atm-cash-pos/git/trees --input - --jq '.sha')
NEW_COMMIT=$(gh api repos/ghassan-abunada/atm-cash-pos/git/commits \
  -f message="<msg>" -f "tree=$NEW_TREE" -f "parents[]=$PARENT_SHA" --jq '.sha')
gh api repos/ghassan-abunada/atm-cash-pos/git/refs/heads/main -X PATCH -f sha=$NEW_COMMIT
```

Batch multiple files by including more objects in the `tree` array.
