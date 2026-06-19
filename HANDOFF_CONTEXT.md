# L.U.C.A.S macOS Handoff Context

L.U.C.A.S means Lot Upload, Comping & Assignment System.

## Current Snapshot

- Repo: `C:\Users\User\Documents\Codex\2026-06-13\card-pipeline-mac`
- Remote: `https://github.com/mikegrossbarth/card_pipeline_mac.git`
- Branches kept current: `main` and `master`
- Latest commit at handoff time: see `git log -1 --oneline`
- Purpose: macOS-focused fork/copy so Mac setup and CourtYard automation can change without touching the Windows project.
- Current visible tabs: `Home`, `Create`, `Comp`, `Receive`, `Assignment`, `Payouts/Tabs`, `Inventory`, `Profit`
- Mac setup walkthrough: `FIRST_RUN_SETUP.md`

The old visible `Review` workflow was split into `Receive` and `Assignment`. Many internal names still use `review_*`; that is intentional legacy naming to avoid risky churn.

## Future Ideas

### iPhone Companion App / TestFlight Path

User is interested in a future L.U.C.A.S phone companion, not a full desktop clone. Preferred scope is inventory-only:

- Inventory lookup/search/filter.
- Quick add to active inventory.
- Mark inventory cards sold with sale price and optional company/buyer.
- No Card Ladder, CourtYard, Chrome extension, or other automation on the phone.

Recommended path:

1. Start with a local/private mobile web companion served by desktop L.U.C.A.S or a small local backend.
2. Keep the backend/API boundaries clean so the same inventory lookup/add/sold actions can later power a PWA or TestFlight app.
3. Later wrap/rebuild the mobile UI with React Native, Expo, Capacitor, or native Swift for TestFlight if the workflow proves useful.
4. If remote access is needed, move the API behind proper auth and concurrency-safe shared data handling.

Important design note: the hard part is not the mobile UI; it is safely exposing shared inventory/profit ledger actions so two people cannot conflict while adding or marking cards sold. Desktop automation should stay desktop/server-side.

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
- `Inventory` tracks active person-level inventory in `inventory_ledger.json`.
- Received cards that are not checked for the company pile are automatically added to active inventory for the assigned person.
- `Inventory` fast-loads the saved ledger on app open. `Reconcile Received` backfills active inventory from received-marked sheet rows that are not already present in company sheets. This is needed for sheets received before inventory capture existed.
- Inventory reconcile skips orphan/deleted sheet files that no longer have an assigned-person Home marker, so deleted sheets are not re-backfilled as `Unassigned`.
- `Inventory` displays and exports `Best Company` and `Payout`; table export uses only the current filtered rows, and the Inventory `Refresh` button enriches only rows visible under the current person/sport/search/price/status filters.
- Inventory preserves separate `Comps`, `Card Ladder`, and `CY` values internally; the visible `Value` updates to the source value used by the winning best-company recommendation, so Card Ladder-only companies like Fanatics are not forced to use comps.
- New or reconciled inventory rows recalculate stale `NOBODY TAKES` assignment fields for the assigned person before saving, so person-specific payout rules apply in Inventory.
- Changing a sheet's assigned person retargets existing inventory rows from that source sheet to the new person and rebuilds their inventory keys, so reconcile does not duplicate the same cert/source under multiple owners.
- Inventory shows the filtered card count, purchase-price total, and inventory value total in the upper-right header, with vertical/horizontal scrollbars on the inventory table.
- Inventory includes a `Search Cert/Card` filter above the table for quickly narrowing hundreds of rows by cert number or card title text.
- Inventory rows support right-click `Copy Cell` and `Copy Row`; copy actions never edit ledger data.
- Active inventory cards with a real Best Company can be moved to company sheets from the Inventory table right-click menu; `NOBODY TAKES` rows do not show the move option. The move runs assignment recommendations, writes company/profit rows, and marks those inventory records as `Company Sheet`.
- Inventory cards can be sold from the right-click menu with `Mark Sold`; one app-styled modal captures sale price and optional company/buyer. Blank company/buyer records the sale under that person's `General Sold` sheet.
- Create has local `Network Mode` in `lucas_settings.json`. When off, Create hides Seller/Sheet Type and seller terms do not apply, keeping normal Open Team UI clean. When on, Create exposes optional seller terms for buying from people at fixed rates. The shared runtime file is `ASSIGNMENT RULES/seller_terms.csv` with columns `Seller`, `Sheet Type`, `Seller Rate`, and optional `Deduction`; see `seller_terms.example.csv`. `Sheet Type` must match an assignment-rule company, so that company rule decides the per-card value source. If Create's Seller or Sheet Type is blank, L.U.C.A.S leaves purchase prices alone. If both are filled, it sets the Create `Purchase` column to the seller's sell price, separate from the normal best-company assignment payout. With `Seller Rate`, it pays that percentage of the matching company rule value. With `Deduction`, it follows the `Sheet Type` company's existing assignment/payout rules per card and subtracts that deduction from the company payout, e.g. Arena Club with `5%` pays all Arena Club rule payouts minus five percentage points.
- Create saves the selected `Seller` onto the new working sheet marker as its assigned person. When that sheet is moved/marked as `Incoming`, the marker follows it, so `Payouts/Tabs` automatically shows the incoming sheet under that seller without manual payout-marker editing.
- `Payouts/Tabs` active balances pay seller-term people their sheet purchase total only after the sheet is `Received`; seller/non-team buy payouts do not appear while the sheet is merely `Incoming`. Assigned people not listed in `seller_terms.csv` are treated as team members and their active balance is half realized sold profit from `profit_ledger.json`, floored at zero. Unsold estimated payout never creates team profit owed.
- Profit rows can be refunded individually from the `Profit` tab. Refunds remove the sold-card profit/company-sheet row and return that card to active inventory.
- `Payouts/Tabs` has `Delete Person`; it removes a person's name from sheet markers, inventory ownership, and profit ownership while leaving cards/sheets/ledger rows intact.
- Create now has `Manual Entry` mode. Use the `+ Add row` line in the Create table, then double-click cells to edit. The extra toolbar button was removed.
- Card Ladder recovery note, 2026-06-17: known-good helper version is `2026-06-17-no-blind-grader-option-v22`. The verified CGC grader test opens the cert modal, uses trusted debugger clicks only when synthetic clicks fail, selects CGC, and leaves the modal open. Do not restore blind guessed grader-option coordinates; they closed/submitted the modal.

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
  inventory_ledger.json
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

Home lists `Incoming`, `Working`, and `Received` sheets. It supports right-click move between `Incoming`, `Working`, and `Received`. Moving a sheet out of `Received` clears received/paid marker state, clears workbook received marks, removes company-sheet rows created from that source sheet, and removes matching profit ledger rows. Payment state is handled in `Payouts/Tabs`.

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
- Person-specific payout overlays can be configured at the top level of `assignment_companies.json` with `person_payouts_source` (or embedded `person_payouts`). The source may be local text/CSV/JSON/XLSX/`.gsheet` or a web/Google Sheet URL. CSV example:

```csv
Person,Company,Min,Max,Rate
Lucas,Arena Club,0,,95%
Mikey,Arena Club,0,,90%
```

- If a person has any policy rows, they are locked to only those listed companies. Listed rates override the company payout tiers for that person; people without policy rows keep the normal company rules and payouts.

## Card Ladder

Chrome extension folder:

```text
cardladder-autocomp/extension
```

Current comping flow uses the normal Chrome profile/session with the unpacked extension loaded. The app queues rows through the local desktop bridge; the extension checks in, opens Card Ladder Sales History, selects the requested grader, and submits cert searches. The grader selector first tries the DOM path, then briefly attaches Chrome debugger input only for trusted clicks on the visible grader bar if Card Ladder ignores synthetic clicks.

`chrome.debugger` is scoped to the Card Ladder `tabId`, but Chrome may still show a browser-level debugger warning/banner while trusted clicks are active. That warning is expected when the trusted fallback runs.

The desktop bridge binds to the first available port from `8765` to `8772`. The extension manifest grants access to the same local range.

Common gotchas:

- User must be logged into Card Ladder in the Chrome profile where the unpacked extension is loaded.
- Old unpacked extension versions should be removed or disabled.
- Current extension/background version: `2026-06-17-no-blind-grader-option-v22`.
- Current content-script version: `2026-06-17-no-blind-grader-option-v22`.
- Current bridge expected helper version: `2026-06-17-no-blind-grader-option-v22`.
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

- `main` and `master` should both be pushed after the cleanup/handoff commit.
- Working tree should be clean after the handoff commit.

## New Chat Bootstrap

Tell a new chat:

```text
Work in C:\Users\User\Documents\Codex\2026-06-04\card_pipeline for Windows and C:\Users\User\Documents\Codex\2026-06-13\card-pipeline-mac for Mac. Read HANDOFF_CONTEXT.md first. Current known-good Card Ladder helper is 2026-06-17-no-blind-grader-option-v22. Do not reintroduce blind guessed grader-option coordinates; v22 fixed CGC by opening the cert modal, avoiding blind option clicks, re-preparing the modal if synthetic selection closes it, then using trusted chrome.debugger clicks on the visible grader bar only as fallback. The debugger banner is expected during trusted fallback because Chrome owns that UI. Windows has no CourtYard automation; Mac keeps CY automation. Keep main and master in both repos pushed.
```
