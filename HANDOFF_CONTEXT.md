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

Deferred future work, not for the current build: true live-anywhere mobile access would need a small secure cloud backend, real auth/invite tokens, and a cloud queue or database. The preferred low-risk shape is to keep the current mobile/offline queue model, add a hosted queue service such as Supabase later, let the phone write inventory add/sold/expense actions to that queue from anywhere, and let desktop L.U.C.A.S pull/apply those actions through the same ledger-safe queue importer. This is intentionally postponed because it is more setup/security/deployment work than the current local/offline mobile companion needs.

## Latest Completed Work

- Inventory photos are now linked from a phone-exported album folder. L.U.C.A.S creates/uses `CARD_PIPELINE/INVENTORY PHOTOS`, or the user can click `Inventory -> Photo Folder` to choose an iCloud/Drive-synced export folder without editing JSON. It scans every three hours and via `Inventory -> Scan Photos`, OCRs certs from new/changed images, links matching certs to active inventory rows, shows a `Photos` count, exports photo paths, and deletes linked active-album files when cards are sold/deleted/moved to a company sheet unless another active row still uses the same photo. During scans the Inventory status line shows the exact scanned folder and per-file progress. Linked rows expose right-click `Open Photo` and `Open Photo Folder` quick actions. This intentionally uses only the album/iCloud-folder half of the old automatic-inventory-update thread; no Instagram/R2 posting code was brought in. Setup doc: `docs/INVENTORY_PHOTO_ALBUM.md`.
- Home refresh now reconciles duplicate/accounted incoming and working sheets against inventory/profit ledgers. Fully accounted sheets are marked received, moved to `RECEIVED SHEETS`, and logged; partially accounted sheets stay put and trigger a loud warning so manually re-added incoming sheets cannot silently duplicate cards already in inventory/company/sold ledgers. Live `KEVIN_HAMLIN_DALTON.xlsx` was reconciled to Received from the Windows/shared pipeline on 2026-06-30.
- Inventory `Move to Best Company Sheet` now trusts the visible stored Best Company and Est. Payout on active inventory rows instead of requiring a fresh assignment recalculation to still match. This fixes rows like `FANATICS / $13.30` failing with "No selected inventory cards matched an assignable company" after assignment rules/context drift.
- Weekly company tabs now roll over on Monday at 8:00 PM local time. Sunday midnight no longer starts the next company-sheet week; Monday before 8 PM still writes to the prior `Week of YYYY-MM-DD` tab.
- Comp now shows a non-editable `Received` marker column for loaded Incoming/Working sheets. The marker reads the workbook `RECEIVED` column and `Save Back to Source Sheet` preserves it, so partial package check-ins remain visible while comp/edit work continues.
- Comp tab now has `Delete Selected` for loaded Incoming/Working sheet rows. It removes the selected row(s), rekeys row/source metadata, marks the comp sheet unsaved, and `Save Back to Source Sheet` persists the deletion to the source workbook.
- Receive rows now guarantee visible Best Company and Est. Payout when enough value data exists. If a matched incoming/working sheet row lacks saved assignment fields, Receive immediately recalculates through the assignment engine so barcode-scanned rows are not left blank.
- Receive barcode scans now refresh the incoming/working cert index when an existing cert match is stale and missing assignment fields, so Best Company and Est. Payout populate from sheets that were assigned after startup. Startup indexing now uses the same incoming plus working merge behavior and preserves sport/category in the receive match.
- Automatic Card Ladder comp results now fill a blank `Sport`/`Category` from the returned profile/card title, matching the manual title-edit behavior. Manually entered sport values are preserved and not overwritten.
- Receive now looks up scanned/manual rows across both `INCOMING SHEETS` and `WORKING SHEETS` and merges duplicate cert matches so nonblank assignment fields win. `write_working_sheet()` now persists `Best Company`, `Estimated Payout`, and `Status`, fixing blank Best Company/Est. Payout values in the Receive tab after assignment values were saved back to a Working sheet.
- Payouts/Tabs now includes expenses in team payout math. Sold-sheet expenses reduce that sheet's realized profit before the 50% team split, loose person expenses appear as an `Expense Adjustments` payout row, summary balances show Expenses and Net Profit, and balance owed can go negative if expenses exceed profit. Seller payouts remain based on seller terms/purchase obligations.
- Mac installer now assumes a fresh Mac may have nothing installed: `install_dependencies.sh` bootstraps Homebrew when missing, installs/verifies Homebrew Python, the matching versioned `python-tk@...` formula, `cliclick`, and `tesseract`, marks `scripts/macos/cgscroll` executable, and best-effort installs Google Chrome and Google Drive for desktop through Homebrew casks. CourtYard/CYCardScanner still has to be installed/opened/logged-in manually and granted macOS privacy permissions.
- Release-readiness audit: Company Rules naming was aligned across setup docs, user guide images, diagnostics, and OAuth error copy; the stale Network Mode handoff note was corrected so seller sheets may save before value data exists and show pending seller-payout warnings until comps/CL/CY values are available. No Card Ladder helper fallback behavior was changed.
- Documentation refresh: `docs/LUCAS_USER_GUIDE.md`, `FIRST_RUN_SETUP.md`, `README.md`, and this handoff were updated for current Network Mode/People Rules behavior, per-company Google Keep sync, Fanatics weekly sheet format, and current setup workflow. The user guide now includes `docs/images/network-mode-people-rules.svg`.
- People Rules now labels seller percentage fields as `Seller Rate %` and `Deduction %`. The UI accepts numbers only, such as `90`, `92.5`, `10`, or `10.5`, and rejects values with percent signs. Legacy CSV values with `%` still load and display as number-only values. Live shared `seller_terms.csv` was normalized to `Jon,ARENA CLUB,,10`, preserving deduction semantics without the percent sign.
- Company Rules now exposes per-company `Sync Google Keep` inside each selected company's Rule Source panel. The old app-wide Keep sync flow was removed because companies can use different Keep notes.
- Fanatics weekly company sheets now use the required visible front columns `Category`, `Card`, `Grade`, `Cert #`, `CL Value`, and `Payout`, with L.U.C.A.S tracking/profit columns trailing to the right. Existing Fanatics weekly tabs migrate when L.U.C.A.S touches the workbook. The live current-week Windows/shared Drive `FANATICS.xlsx` tab `Week of 2026-06-22` was migrated and verified.
- Operational UX pass: the old `Setup Check` header action is now `System Health` and includes setup diagnostics plus shared conflict files, stale locks, assignment companies, People Rules Health, inventory/profit ledger presence, mobile queue log presence, and the new `activity_log.json`. `Activity Log` is available from the header and records create, receive, sheet move/delete, inventory add/sold/move/delete, mobile inventory/sold actions, expense add/delete, refund, and profit recovery events.
- Ambiguous buttons were renamed for clarity: examples include `Refresh Sheet List`, `Refresh Incoming Index`, `Refresh Received Sheets`, `Refresh Home View`, `Sync Received to Inventory`, `Update Best Company/Payouts`, `Recomp Visible Cards`, `Refresh View`, and `Recover Sold Ledger`.
- Inventory rows now have right-click `Explain Assignment`, showing the exact assignment engine recommendation plus each company decision/reason/source value/payout. This is the place to debug why an inventory row says `NOBODY TAKES`, Fanatics, Arena Club, etc.
- Inventory and Profit tables now support spreadsheet-style sorting by clicking column headers. Click once to sort, click the same header again to flip direction; active sort headers show `^`/`v`. Money/count columns sort numerically and blanks stay at the bottom.
- Tab canvases now auto-fit to the current window and hide their horizontal/vertical scrollbars until actual tab content is larger than the visible area, so normal-size windows behave like a regular resizable app instead of always showing fallback scrollbars.
- Comp tab layout now avoids outer tab scrolling at the default app size by letting the Active Sheets list, comp table, and comp controls resize inside the tab; the comp table keeps its own horizontal scrollbar for spreadsheet columns.
- Main tabs other than Profit fit at the default app size without outer tab scrollbars in runtime geometry checks. Wide/tall data stays scrollable inside the relevant table/list areas instead of forcing the whole tab canvas to pan.
- Profit tab was restored to the original stacked layout with a full-width graph above search and ledger rows; this keeps the graph usable even if Profit may need the outer vertical fallback on shorter windows.
- Inventory's manual add button is now `Add Card` instead of `Add Raw Card`. It can create active inventory directly without making/receiving a sheet first. Person, card description, and purchase are still required; cert and grader are optional. If a cert is provided, the card is saved as a manual graded inventory row even when grader is blank. If cert is blank, L.U.C.A.S generates a stable `RAW-YYYYMMDD-####` item id.
- Inventory `Recomp Visible Cards` now writes CY Estimate and CY Confidence back to `inventory_ledger.json`; `Empty Comps Only` treats blank CY Confidence as missing CY data so confidence-only gaps can be repaired.
- `Recover Sold Ledger` now scans company sheets in the background and merges missing sold-ledger rows afterward, instead of freezing the UI during the scan. Errors use a copyable diagnostic popup.
- Seller/team payout classification now honors sheet-level seller-term metadata first. This prevents a person who is both a seller and a team member from having normal sold-profit team payouts hidden just because their name appears in `seller_terms.csv`. Legacy seller sheets with no marker metadata still fall back to seller-name detection.
- Mac header has `Mobile Help`, which displays/copies the current mobile URL, PIN, same-Wi-Fi note, and offline queue reminder.
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
- Assignment audit note: default `Comps` value-source assignment now correctly falls back from comps to Card Ladder value to CY Estimate. `CourtYard`, `court yard`, `CY value`, and similar aliases normalize to `CY Estimate` in both config load and Company Rules UI. Grade-bounded rules now reject rows where the grader matches but no numeric grade was parsed.
- Company sheets now use one workbook per company with weekly tabs:
  - `COMPANY SHEETS/<Company>/<Company>.xlsx`
  - weekly tab name: `Week of YYYY-MM-DD`
- CourtYard/CY weekly company sheets use the same visible front columns as ingestable CY sheets: `Grader`, `Cert`, `Description`, `Grade`, `Purchase`, `Estimate`, `Confidence`. L.U.C.A.S tracking/profit columns are kept to the right as hidden columns so profit backfill/refunds/source tracking still work.
- Old legacy weekly company files remain readable for profit backfill.
- Weekly company-sheet tabs roll forward Monday at 8:00 PM local time.
- `Inventory` tracks active person-level inventory in `inventory_ledger.json`.
- Received cards that are not checked for the company pile are automatically added to active inventory for the assigned person.
- Marking a sheet `All Received` from Home now also syncs that newly received sheet's non-company rows into active inventory for the assigned person; older sheets can still be backfilled with `Inventory` -> `Reconcile Received`.
- Moving a sheet backward out of `Received` now removes inventory rows for that source sheet in addition to clearing received marks and removing company/profit rows. This fixes the case where an Incoming sheet could still appear in Inventory after being temporarily marked/moved as Received.
- `Inventory` fast-loads the saved ledger on app open. `Reconcile Received` backfills active inventory from received-marked sheet rows that are not already present in company sheets. This is needed for sheets received before inventory capture existed.
- Inventory reconcile skips orphan/deleted sheet files that no longer have an assigned-person Home marker, so deleted sheets are not re-backfilled as `Unassigned`.
- `Inventory` displays and exports `Best Company` and `Payout`; table export uses only the current filtered rows, and the Inventory `Refresh` button enriches only rows visible under the current person/sport/search/price/status filters.
- Inventory preserves separate `Comps`, `Card Ladder`, and `CY` values internally; the visible `Value` updates to the source value used by the winning best-company recommendation, so Card Ladder-only companies like Fanatics are not forced to use comps.
- Inventory sport/category now prefers the explicit sheet/manual `Sport` value before trying to infer from the card title. Reconcile/all-received inventory creation, refund-to-inventory, and old ledger hydration all preserve or recover source-sheet sport, so manually entered baseball/football/etc. does not disappear when a title has no known player signal.
- New or reconciled inventory rows recalculate stale `NOBODY TAKES` assignment fields for the assigned person before saving, so person-specific payout rules apply in Inventory.
- Changing a sheet's assigned person retargets existing inventory rows from that source sheet to the new person and rebuilds their inventory keys, so reconcile does not duplicate the same cert/source under multiple owners.
- Inventory shows the filtered card count, purchase-price total, and inventory value total in the upper-right header, with vertical/horizontal scrollbars on the inventory table.
- Inventory includes a `Search Cert/Card` filter above the table for quickly narrowing hundreds of rows by cert number or card title text.
- Inventory rows support right-click `Copy Cell` and `Copy Row`; copy actions never edit ledger data.
- Active inventory cards with a real Best Company can be moved to company sheets from the Inventory table right-click menu; `NOBODY TAKES` rows do not show the move option. The move runs assignment recommendations, writes company/profit rows, and marks those inventory records as `Company Sheet`.
- Inventory cards can be sold from the right-click menu with `Mark Sold`; one app-styled modal captures sale price and optional company/buyer. Blank company/buyer records the sale under that person's `General Sold` sheet.
- Create has local `Network Mode` in `lucas_settings.json`. When off, Create hides Seller/Sheet Type and seller terms do not apply, keeping normal Open Team UI clean. When on, Create exposes optional seller terms for buying from people at fixed rates. Seller terms are edited from `Company Rules -> People Rules` and stored in the shared runtime file `ASSIGNMENT RULES/seller_terms.csv` with columns `Seller`, `Sheet Type`, `Seller Rate`, and optional `Deduction`; see `seller_terms.example.csv`. The People Rules UI labels the percentage fields as `Seller Rate %` and `Deduction %`; type numbers only, decimals allowed, without percent signs. `Sheet Type` must match an assignment-rule company, so that company rule decides the per-card value source. If Create's Seller or Sheet Type is blank, L.U.C.A.S leaves purchase prices alone. If both are filled, it sets the Create `Purchase` column to the seller's sell price, separate from the normal best-company assignment payout. With `Seller Rate %`, it pays that percentage of the matching company rule value. With `Deduction %`, it follows the `Sheet Type` company's existing assignment/payout rules per card and subtracts that percentage of the company source value from the company payout, e.g. Arena Club with company payout `95` on value `100` and deduction `10` writes purchase `85`.
- Create saves the selected `Seller` onto the new working sheet marker as its assigned person. Network seller saves require both Seller and Sheet Type plus a matching seller_terms.csv row. If the current rows do not yet have the value source needed to calculate seller payout, the sheet can still save with seller metadata and Home shows the pending seller-payout warning until comps/Card Ladder/CY values are added. Valid seller sheets store seller-term metadata such as `seller_terms_applied`, `seller_sheet_type`, `seller_rate`, and/or `seller_deduction` in `sheet_markers.json`. When that sheet is moved/marked as `Incoming`, the marker follows it, so `Payouts/Tabs` can show the incoming sheet under that seller once payout values are ready.
- Company Rules has a `People Rules Health` panel next to `Google Status`. It checks `ASSIGNMENT RULES/seller_terms.csv` for missing/unreadable files, empty CSVs, invalid Seller Rate/Deduction values, duplicate Seller/Sheet Type rows, inactive Sheet Type companies, missing Sheet Type assignment companies, and shows parsed rates/deductions. Buttons refresh the status, open the terms folder, and copy diagnostic details.
- `Payouts/Tabs` active balances pay seller-term people their sheet purchase total only after the sheet is `Received`; seller/non-team buy payouts do not appear while the sheet is merely `Incoming`. Assigned people not listed in `seller_terms.csv` are treated as team members, and their payout rows are generated from sold-profit entries in `profit_ledger.json`; team balance is half realized sold profit, floored at zero. Unsold estimated payout never creates team profit owed.
- Profit rows can be refunded individually from the `Profit` tab. Refunds remove the sold-card profit/company-sheet row and return that card to active inventory.
- `Profit` has `Add Expense` for person-level expenses. Expense categories are `Travel`, `Supplies`, `Travel Meal`, `Fees`, and `Shipping`; expenses record a date and deduct from that person's month/year/YTD/overall profit and charts. Expenses can be general, tied to a sold sheet, or tied to an individual sold card through dropdowns built from sold profit rows. Cert is not shown as an expense field, though card-tied expenses may keep it internally for exact matching. `Profit` also has an `Expenses` view, and the profit metric shows sales, gross profit, expenses, and net profit. Expense rows are profit-only adjustments and cannot be refunded to inventory, but selected expense rows can be deleted from the profit ledger.
- Inventory `Mark Sold` can now add an optional card-linked expense in the same popup using expense category, amount, and notes. `Profit` -> `Add Expense` keeps Sheet active when tying to `Card`, and the Card dropdown filters to the selected sold sheet so card/sheet expenses can be selected reliably.
- Mac `mobile_app` is now offline-capable for write capture. If the phone cannot reach the desktop bridge, inventory add, mark-sold, and expense actions are stored in a local Sync queue. The phone can later `Sync Now` to `/mobile/api/sync/queue` or `Export Queue` as JSON. Desktop Mac L.U.C.A.S has `Inventory` -> `Import Mobile Queue`, applies queued actions through the existing mobile add/sold/expense methods, and records applied action IDs in `CARD_PIPELINE/mobile_action_log.json` so re-importing the same queue does not duplicate ledger writes. Search/profit/payout mobile views still require a live desktop bridge.
- `Profit` normal `Refresh` is now a fast ledger-only refresh. Use `Deep Sync` when company sheets need to be scanned/backfilled into `profit_ledger.json`; this preserves the recovery path without making every filter/search refresh reread every company workbook.
- Home/startup workbook summaries are cached in memory by file modified time and size for the current app session, so repeat Home refreshes reuse unchanged summaries and only reread changed/new sheets.
- Google Sheets OAuth callback handling now waits for the real `/oauth2callback` success/error response, ignores unrelated local browser requests such as favicon probes, and reports clearer timeout/denial messages. This fixes cases where the browser said Google Sheets connected but L.U.C.A.S still popped `OAuth did not return an authorization code`.
- Date-weighted Card Ladder comps now choose the best-quality comp when multiple comps share the newest stale sale date, rather than blindly taking the first OCR/browser row for that date. This fixes the 2014 Panini Flawless Greats Patches Autographs Gold #22 Joe Montana BGS 9 case where the cleaner same-day ALT `$2,760` comp was present but the noisy BECKETT `$1,424` row appeared first.
- Card Ladder comp parsing is hardened against row bleed and date drift: the Chrome content script trims each parsed comp chunk before the next source row only after the current row already has its own date and price, prefers prices after the comp date, preserves source words that legitimately appear inside card titles, penalizes titles with conflicting dollar amounts, and dedupes same-day/same-source/similar-title rows even when the noisy duplicate has a different price. The bridge normalizes `Sept.`, embedded label text, ISO/slashed dates, and two-digit-year dates before dedupe/date weighting. Date Weighted excludes undated/unparseable comps from its average when dated comps are available.
- L.U.C.A.S writes lightweight performance timings to `CARD_PIPELINE/lucas_performance.log` for operations slower than `0.25s` by default. Set `LUCAS_PERF_LOG_SECONDS` to a lower value for more detail. Timed areas include app init, startup sheet scans, Google Sheet cache refresh, Home, Inventory, Profit/company-sheet scans, Payouts, assignment rule loads, and assignment recommendation batches.
- `Delete Person` is intentionally not exposed in the visible app UI. A backend helper still exists for maintainer/Codex-driven cleanup; it removes a person's name from sheet markers, inventory ownership, and profit ownership while leaving cards/sheets/ledger rows intact.
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
- `assignment_config_ui.py`: Company Rules and People Rules popups.
- `intake_io.py`: spreadsheet import/export, receive marking, company-sheet append, weekly tabs, archive/profit extraction helpers.
- `shared_state.py`: atomic JSON writes and shared-folder locks.
- `google_sheets_import.py`: OAuth and Google Sheets export/read helpers.
- `comp_engine`: Card Ladder bridge, workbook row model, comp strategy, screenshot OCR fallback, and Mac CourtYard automation.
- `photo_tool`: photo OCR helper used by Create and Receive.
- `cardladder-autocomp/extension`: unpacked Chrome extension for Card Ladder automation.
- `tests/test_shared_workflows.py`: offline regression suite.

## Required Mac Setup

Each Mac needs:

- Google Chrome, installed manually or by `install_dependencies.sh`.
- Python 3.11+ with Tkinter, installed manually or by `install_dependencies.sh`.
- Homebrew, `cliclick`, `tesseract`, and executable `scripts/macos/cgscroll`; `install_dependencies.sh` handles these on a clean Mac.
- Project `.venv` created by `./install_dependencies.sh`.
- Local `.env` created from `.env.example`.
- `GOOGLE_API_KEY` for Photo OCR and Card Ladder screenshot OCR fallback.
- Google billing, payment method, spend cap, and billing budget alerts.
- Google Sheets OAuth desktop credentials in `.env`.
- `Connect Google` completed once inside Company Rules.
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
  mobile_action_log.json
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
- generated `mobile_action_log.json` in the shared pipeline folder

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

Profit reads `profit_ledger.json`, current company workbook tabs, and legacy weekly company files. It includes person filters, daily profit chart, `Sold Cards`, and grouped `Sold Sheets`. Profit chart axes are based on the selected period, so sparse YTD/month/year activity still renders the full period instead of collapsing to the few sale dates.

## Company Rules

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
- If a Card Ladder lookup visibly opens a cert but the helper throws before capture completes, the extension now posts an `extension_error` row result; LUCAS records `Card Ladder extension error` instead of leaving the row blank. Bridge result posts retry before failing.

## Tests And Verification

Recent verification:

```text
C:\Users\User\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe -m py_compile app.py
C:\Users\User\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe -m unittest tests.test_shared_workflows -v
C:\Users\User\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe tests/test_extension_parser.js
```

Latest broad bug-hunt verification:

- Fixed scoped comp/assignment refresh so a result payload only updates rows it actually contains; unrelated comp rows keep their existing best company/payout.
- Synced Mac with the scoped comp assignment path, including bridge `updated_row_ids`, pending comp assignment ids, and missing `_refresh_comp_table`.
- Restored Mac Home parity for paid received-sheet archiving after 14 days and Home right-click `Delete Sheet`; delete removes the sheet file and matching inventory ledger rows.
- Added a regression test for the Card Ladder extension parser so an 8-line chunk containing the next sale row cannot assign row B's price to row A.
- Confirmed assignment tests still cover unlicensed, DNB-over-threshold, GOAT Bonus 1, sport aliases, seller policies, and team half-profit payouts.

Last full Mac result after recovery in this Windows Codex workspace: `120 tests OK`.
Last Mac extension parser result: `extension parser regression ok`.
Last Mac startup smoke: `mac app startup ok`.

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
