# AI Research Publication Verification System

AI-assisted research publication management and verification system built with Google Apps Script, Google Sheets, Google Forms, Crossref and Groq AI.

## What it does

- Supports multiple Google Forms in one system
- Keeps one centralized `MASTER_DATA` sheet
- Dynamically maps different Form question/column names
- Synchronizes Form responses into Master Data
- Verifies DOI/title metadata using Crossref
- Stores verification results separately in `VERIFICATION`
- Uses Groq AI only as an additional comparison layer for ambiguous cases
- Tracks verification status with dropdowns
- Maintains system logs and diagnostics
- Uses caching and batching to reduce unnecessary API calls
- Keeps API keys in Apps Script Properties instead of source code

## Architecture

`FORM_REGISTRY` → Google Forms → `MASTER_DATA` → Crossref → `VERIFICATION` → optional Groq AI

`SYSTEM_LOG` stores operational diagnostics and errors.

## Important security rule

Do **not** commit:

- Groq API keys
- Google Form IDs/URLs that should remain private
- Google Sheet IDs that should remain private
- researcher personal information
- confidential institutional data
- production-only configuration

The public repository should contain sanitized source code and documentation only.

## Setup

1. Create/open the Google Sheet used by the application.
2. Open **Extensions → Apps Script**.
3. Replace the Apps Script source with `Code.gs` from this repository.
4. Save and reload the spreadsheet.
5. Run **Research System → Setup / Repair System**.
6. Authorize the script with the Google account that can access the Forms.
7. Add Forms through **Research System → Add Form**.
8. Run **Sync All Forms**.
9. Run **Verify All Publications**.

## Groq configuration

The code reads the key from Apps Script Script Properties. Do not put the key in `Code.gs`.

Property name:

`GROQ_API_KEY`

The configured model and endpoint are defined in the `CONFIG.AI` section of `Code.gs`.

## Main sheets

- `FORM_REGISTRY` — registered Forms and synchronization metadata
- `MASTER_DATA` — centralized submitted data from multiple Forms
- `VERIFICATION` — Crossref and AI verification results
- `SYSTEM_LOG` — diagnostics and errors

## AI design

Groq AI does not replace authoritative metadata verification. Crossref remains the primary metadata source. AI is used as a supporting comparison layer for title/author similarity and ambiguous cases.

## License

See `LICENSE`.
