# L.U.C.A.S macOS Port Handoff

L.U.C.A.S means Lot Upload, Comping & Assignment System.

## Current Snapshot

- Project copy: `C:\Users\User\Documents\Codex\2026-06-13\card-pipeline-mac`
- Source copied from: `C:\Users\User\Documents\Codex\2026-06-04\card_pipeline`
- Purpose: macOS-downloadable fork/copy so Mac setup can change without touching the original Windows project.
- Current visible tabs: `Home`, `Create`, `Comp`, `Receive`, `Assignment`, `Payouts/Tabs`, `Profit`.
- Mac setup walkthrough: `FIRST_RUN_SETUP.md`

The old visible `Review` workflow was split into `Receive` and `Assignment`. Many internal Python names still use `review_*`; that is intentional legacy naming to avoid risky churn.

## Mac Port Changes

- Added `install_dependencies.sh`.
- Added `run_card_pipeline.sh`.
- Added double-click launcher `Run Card Pipeline.command`.
- Added optional Finder launcher builder `create_macos_app.sh`.
- Removed Windows-only `.bat` and `.vbs` launchers from this Mac copy.
- Updated `.env.example` and example company paths to macOS Google Drive paths.
- Updated Card Ladder extension docs to use POSIX-style paths.
- Updated local identity fallback to use `LUCAS_DISPLAY_NAME`, then `$USER`, then `$USERNAME`.
- Added Assignment Rules company list filters for name, active companies, and inactive companies.
- Removed copied generated `work/` cache content and Python cache folders from the Mac copy.

## Project Layout

- `app.py`: main Tkinter desktop app and UI workflows.
- `assignment_engine.py`: company rule parsing, payout parsing, category/player matching, and best-company recommendation.
- `assignment_config_ui.py`: Assignment Rules popup.
- `intake_io.py`: spreadsheet import/export, receive marking, company-sheet append, and profit record extraction.
- `shared_state.py`: atomic JSON writes and shared-folder lock helpers.
- `google_sheets_import.py`: desktop OAuth flow and Google Sheets export/read helpers.
- `comp_engine`: Card Ladder bridge, workbook rows, comp strategy, and screenshot OCR fallback.
- `photo_tool`: bundled photo OCR helper used by Create and Receive.
- `cardladder-autocomp/extension`: unpacked Chrome extension for Card Ladder automation.
- `tests/test_shared_workflows.py`: committed offline regression suite.

## Required Mac Setup

For a full setup, each Mac needs:

- Google Chrome.
- Python 3.11+ with Tkinter.
- Project `.venv` created by `./install_dependencies.sh`.
- Local `.env` created from `.env.example`.
- `GOOGLE_API_KEY` for Photo OCR and Card Ladder screenshot OCR fallback.
- Google billing, payment method, spend cap, and billing budget alerts.
- Google Sheets OAuth desktop credentials in `.env`.
- `Connect Google` completed once inside Assignment Rules.
- Google Drive for desktop if using the shared Drive pipeline.
- Card Ladder account, active Chrome login, and the unpacked extension loaded from `cardladder-autocomp/extension`.
- Shared pipeline folder selected through the `Working Folder` button.
- Active assignment companies with rule and payout sources.

## Local-Only Files

Do not commit these:

- `.env`
- `.venv`
- `lucas_settings.json`
- `lucas_user_identity.json`
- `lucas_google_sheets_token.json`
- `assignment_companies.json`
- generated debug screenshots/logs
- `work/`
- `outputs/`

## Shared Pipeline Folder

The user chooses the actual `WORKING SHEETS` folder with the `Working Folder` button. L.U.C.A.S uses that folder's parent as the pipeline root.

Expected shared root shape:

```text
CARD_PIPELINE
  WORKING SHEETS
  INCOMING SHEETS
  RECEIVED SHEETS
  COMPANY SHEETS
  ASSIGNMENT RULES
  sheet_markers.json
  profit_ledger.json
  unassigned_players.json
  assignment_player_overrides.json
  .locks
```

Typical Google Drive for desktop location on macOS:

```text
/Users/yourname/Library/CloudStorage/GoogleDrive-your.email@gmail.com/My Drive/CARD_PIPELINE/WORKING SHEETS
```

Each user keeps their own app folder, `.env`, OAuth token, and Card Ladder extension install. Shared writes use `.locks` plus atomic JSON writes to reduce Drive conflict risk.

## Workflows

### Home

Home lists `Incoming`, `Working`, and `Received` sheets. `Edit Markers` handles incoming state, tracking number, all-received state, and assigned person. Payment state is handled in `Payouts/Tabs`.

### Create

Create supports barcode scanner rows, photo OCR rows, and existing spreadsheet import. Output is saved as a working sheet.

### Comp

Comp runs Card Ladder through the local bridge and Chrome extension. It writes Card Ladder value, comps, assignment results, and statuses. Best-company assignment is recalculated when comp values are added or edited.

### Receive

Receive is for physically receiving cards and marking them received in sheets. If `Company Pile` is checked for a row and the row has a real Best Company, marking received appends that card to the company's weekly sheet under `COMPANY SHEETS/<Company Name>`.

Rows marked `NOBODY TAKES` are not appended to company sheets.

### Assignment

Assignment is for loading received/unassigned sheets and recalculating best company/payout when needed.

### Payouts/Tabs

Tracks active balances by assigned person. It includes incoming/unreceived and received/unpaid sheets. Clicking active balances can mark all matching person sheets paid.

### Profit

Profit reads `profit_ledger.json` and backfills from company sheets. It has a person filter, daily profit line chart, `Sold Cards` view, and `Sold Sheets` grouped view.

## Card Ladder

Chrome extension folder:

```text
cardladder-autocomp/extension
```

The desktop bridge binds to the first available port from `8765` to `8772`. The extension manifest grants access to the same local range.

Common gotchas:

- user must be logged into Card Ladder in Chrome
- old unpacked extension versions should be removed/disabled
- app warns if the extension version seen by the bridge is stale
- no-results pages should still preserve the Card Ladder card title when available
- BGS OCR rejects subgrades as slab grades

## Tests And Verification

Primary offline test suite:

```bash
.venv/bin/python -m unittest discover -s tests -v
```

Useful sanity checks:

```bash
.venv/bin/python -m compileall -q .
.venv/bin/python -c "import app; root = app.CardPipelineApp(); root.update_idletasks(); root.destroy(); print('app startup ok')"
```

Current limitation: the Mac launcher and setup scripts were authored in this Windows Codex workspace. They should be validated on an actual Mac before calling the port fully production-ready.
