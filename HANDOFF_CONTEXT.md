# L.U.C.A.S macOS Handoff Context

L.U.C.A.S means Lot Upload, Comping & Assignment System.

## Current Snapshot

- Repo: `C:\Users\User\Documents\Codex\2026-06-13\card-pipeline-mac`
- Remote: `https://github.com/mikegrossbarth/card_pipeline_mac.git`
- Branches kept current: `main` and `master`
- Latest commit at handoff time: this recovery handoff commit
- Purpose: macOS-focused fork/copy so Mac setup and CourtYard automation can change without touching the Windows project.
- Current visible tabs: `Home`, `Create`, `Comp`, `Receive`, `Assignment`, `Payouts/Tabs`, `Profit`
- Mac setup walkthrough: `FIRST_RUN_SETUP.md`

The old visible `Review` workflow was split into `Receive` and `Assignment`. Many internal names still use `review_*`; that is intentional legacy naming to avoid risky churn.

## Latest Completed Work

- Mac repo was brought forward with the Windows parser/company-sheet/manual-create work while preserving Mac-only CourtYard automation.
- Arena Club assignment sheets now support both old and new Arena formats.
- Rules handle broad labels like `Pre-1989` and `ALL Grades`.
- Obvious duplicated-player parse mistakes such as `Tom Tom Brady` are treated as the intended player.
- CY-style sheets can be ingested without confusing `Estimate` with `Purchase`.
- `CY Estimate` and `CY Confidence` are preserved in imported sheets, working sheets, output sheets, company sheets, and display tables.
- Assignment rules can choose their value source per company:
  - `Comps`
  - `Card Ladder value`
  - `CY Estimate`
- Company sheets now use one workbook per company with weekly tabs:
  - `COMPANY SHEETS/<Company>/<Company>.xlsx`
  - weekly tab name: `Week of YYYY-MM-DD`
- Old legacy weekly company files remain readable for profit backfill.
- Sunday at midnight rolls forward to the next Monday's company-sheet tab.
- Create now has `Manual Entry` mode. Use the `+ Add row` line in the Create table, then double-click cells to edit. The extra toolbar button was removed.
- Recovery note, 2026-06-17: Card Ladder was rolled back to the last working comping flow after the later DOM-sweep/grader-verification changes broke real usage. Do not reapply those changes without live browser QA against actual Card Ladder cert searches.

## Mac-Only CourtYard/CY Automation

Mac has active CourtYard lookup support. Windows intentionally does not.

Mac-only pieces include:

- `comp_engine/cy_automation/`
- `comp_engine/cy_automation/cy_macos.py`
- `lookup_cy_buy_price`
- comp source selector:
  - `Card Ladder + CY`
  - `Card Ladder`
  - `CY`
- CY-only runs close the CourtYard app after the last lookup.

Keep this Mac-only unless the user explicitly requests Windows CourtYard support later.

## Mac Port Setup Pieces

- `install_dependencies.sh`
- `run_card_pipeline.sh`
- double-click launcher: `Run Card Pipeline.command`
- optional Finder launcher builder: `create_macos_app.sh`
- macOS Google Drive path examples in docs/config examples
- Card Ladder extension docs use POSIX-style paths
- local identity fallback uses `LUCAS_DISPLAY_NAME`, then `$USER`, then `$USERNAME`

## Project Layout

- `app.py`: main Tkinter desktop app and UI workflows.
- `assignment_engine.py`: company rules, payout parsing, category/player matching, and recommendation logic.
- `assignment_config_ui.py`: Assignment Rules popup.
- `intake_io.py`: spreadsheet import/export, receive marking, company-sheet append, weekly tabs, archive/profit extraction helpers.
- `shared_state.py`: atomic JSON writes and shared-folder locks.
- `google_sheets_import.py`: OAuth and Google Sheets export/read helpers.
- `comp_engine`: Card Ladder bridge, workbook row model, comp strategy, screenshot OCR fallback, and Mac CourtYard automation.
- `photo_tool`: photo OCR helper used by Create and Receive.
- `cardladder-autocomp/extension`: unpacked Chrome extension for Card Ladder automation.
- `tests/test_shared_workflows.py`: offline regression suite.

## Required Mac Setup

Each Mac needs:

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
- CourtYard installed/configured for Mac CY lookup workflows.
- Shared pipeline folder selected through the `Working Folder` button.
- Active assignment companies with rule and payout sources.

## Shared Pipeline Folder

The user chooses the actual `WORKING SHEETS` folder with the `Working Folder` button. L.U.C.A.S uses that folder's parent as the pipeline root.

Expected shared root:

```text
CARD_PIPELINE
  WORKING SHEETS
  INCOMING SHEETS
  RECEIVED SHEETS
  ARCHIVED SHEETS
  COMPANY SHEETS
  ASSIGNMENT RULES
  sheet_markers.json
  weekly_company_sheets.json
  profit_ledger.json
  unassigned_players.json
  assignment_player_overrides.json
  .locks
```

Typical Google Drive for desktop location on macOS:

```text
/Users/yourname/Library/CloudStorage/GoogleDrive-your.email@gmail.com/My Drive/CARD_PIPELINE/WORKING SHEETS
```

Each user keeps their own app folder, `.env`, OAuth token, local app settings, Card Ladder extension install, and CourtYard local setup.

## Local-Only Files

Do not commit:

- `.env`
- `.venv`
- `lucas_settings.json`
- `lucas_user_identity.json`
- `lucas_google_sheets_token.json`
- `assignment_companies.json`
- generated debug screenshots/logs
- generated `work/` or `outputs/` content

## Workflow Notes

### Home

Home lists `Incoming`, `Working`, and `Received` sheets. Payment state is handled in `Payouts/Tabs`.

### Create

Create supports:

- `Barcode Scanner`
- `Manual Entry`
- `Photo OCR`
- `Existing Spreadsheet`

Manual Entry uses the `+ Add row` line in the Create table. Double-click table cells to edit cert, grader, card, purchase, Card Ladder, comps, CY Estimate, or CY Confidence.

### Comp

Comp can run Card Ladder, CY, or both depending on the comp source selector. Best-company assignment recalculates for rows touched by the current comp run, not every row merely because a sheet loaded.

### Receive

Receive marks cards received in source sheets. If `Company Pile` is checked and the row has a real Best Company, marking received appends that card to the current weekly tab in the company workbook. `NOBODY TAKES` rows are not appended.

### Assignment

Assignment can recalculate best company and payout using the configured value source. Failed valued assignments are recorded in `unassigned_players.json`.

### Profit

Profit reads `profit_ledger.json`, current company workbook tabs, and legacy weekly company files. It includes person filters, daily profit chart, `Sold Cards`, and grouped `Sold Sheets`.

## Assignment Rules

Assignment companies live in local-only `assignment_companies.json`.

Important behavior:

- Companies can choose `Comps`, `Card Ladder value`, or `CY Estimate` as assignment value source.
- If a company requires Card Ladder value and the row has none, that company is ignored.
- If a company requires CY Estimate and the row has none, that company is ignored.
- Default value source remains comps first, then Card Ladder value, then CY Estimate.
- A company must accept the card and have a matching payout tier/rate.
- Highest estimated payout wins.
- If no company can take the card, Best Company becomes `NOBODY TAKES`.

## Card Ladder

Chrome extension folder:

```text
cardladder-autocomp/extension
```

Current comping flow uses the normal Chrome profile/session with the unpacked extension loaded. The app queues rows through the local desktop bridge; the extension checks in, opens Card Ladder Sales History, selects the requested grader, and submits cert searches. The grader selector first tries the DOM path, then briefly attaches Chrome debugger input only for trusted clicks on the visible grader bar if Card Ladder ignores synthetic clicks.

The desktop bridge binds to the first available port from `8765` to `8772`. The extension manifest grants access to the same local range.

Common gotchas:

- User must be logged into Card Ladder in the Chrome profile where the unpacked extension is loaded.
- Old unpacked extension versions should be removed or disabled.
- Current restored extension/background version: `2026-06-17-trusted-grader-click-v18`.
- Current restored content-script version: `2026-06-17-trusted-grader-click-v18`.
- App warns if the extension version seen by the bridge is stale.
- No-results pages preserve the Card Ladder title when available.

## Tests And Verification

Recent verification:

```text
C:\Users\User\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe -m py_compile app.py
C:\Users\User\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe -m unittest tests.test_shared_workflows -v
```

Last full Mac result after recovery in this Windows Codex workspace: `39 tests OK`.

On an actual Mac, prefer:

```bash
.venv/bin/python -m unittest discover -s tests -v
.venv/bin/python -m compileall -q .
.venv/bin/python -c "import app; root = app.CardPipelineApp(); root.update_idletasks(); root.destroy(); print('app startup ok')"
```

## Current Git State At Handoff

- `main` and `master` should both point at this recovery handoff commit.
- Working tree was clean after this handoff file update was committed.
