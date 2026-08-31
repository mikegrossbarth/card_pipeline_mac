# Sync UI Fixes

## Windows Audit - 2026-08-31

- Windows local-only `sync UI fixes.md` was removed from `mikegrossbarth/card_pipeline`; this universal tracker is now the sync source of truth.
- Windows parity commit `cd4e2c7 Sync Windows receive and photo picker safety` completes the Mac-origin Receive row-ref cert hydration and sold-photo attach/unattached-picker safety items.
- Verified stale checklist item: Windows already has `425537c Add eBay connect account model`. Do not keep that item marked pending for Windows.
- Confirmed scope guard: Windows still intentionally skips Instagram Inventory Sync work, and this audit ignored inventory photo scan performance work.

## Sync Tracker Location - 2026-08-31

- The repo-local Windows `sync UI fixes.md` note is retired. Use this tracked file, `docs/SYNC_UI_FIXES.md`, as the parity source of truth.
- When a change lands only in Mac or only in Windows, update this file in the repo where the change lands so the other platform's follow-up is visible during the next pull/audit.
- If an old local note conflicts with this file, treat the old local note as stale unless Michael explicitly says otherwise.

## Payout Refresh Button

- Origin: Windows
- Implemented on origin: Yes
- Mirrored to other platform: Mac pending
- Source commit/repo: `64f69bd Add payout refresh button` and `9c400e0 Document payout refresh button` in `mikegrossbarth/card_pipeline`
- Target to mirror: Mac LUCAS `mikegrossbarth/card_pipeline_mac`
- Summary: Payouts/Tabs now has a `Refresh Payouts` button that reloads payout markers and sheet summaries from disk before rebuilding payout rows, so newly sold cards or payout-marker changes show without closing and reopening L.U.C.A.S.
- Notes / avoid porting: Keep this as a lightweight payout-tab refresh. Do not use it to recalculate unrelated inventory/company assignments.

## Sold Photo Attach Picker Safety

- Origin: Mac
- Implemented on origin: Yes
- Mirrored to other platform: Windows complete
- Source commit/repo: Current Mac sold-photo lifecycle behavior in `mikegrossbarth/card_pipeline_mac`
- Windows source commit/repo: `cd4e2c7 Sync Windows receive and photo picker safety` in `mikegrossbarth/card_pipeline`
- Summary: Attach Photo on Windows now uses an unattached-photo picker backed by inventory photo state, active inventory photo references, and sold profit photo references. Photos marked `sold_inventory`, photos whose saved cert belongs to sold inventory, and photo paths/hashes preserved on sold profit rows are hidden from the unattached picker so sold-card photos are not accidentally reattached as available inventory photos.
- Notes / avoid porting: This is not Inventory Photo Scan performance work and does not port Instagram Inventory Sync. Do not expose sold/refund-preserved photos as unattached attach candidates.

## Receive Row Ref Cert Hydration

- Origin: Mac
- Implemented on origin: Yes
- Mirrored to other platform: Windows complete
- Source commit/repo: Current Mac `row_ref_certs` receive behavior in `mikegrossbarth/card_pipeline_mac`
- Windows source commit/repo: `cd4e2c7 Sync Windows receive and photo picker safety` in `mikegrossbarth/card_pipeline`
- Summary: `mark_received_in_workbooks()` now returns `row_ref_certs` for workbook rows selected by row reference that actually contain a cert. Receive hydrates those certs back into matching review rows before creating inventory records, clearing temporary `RAW-*` IDs so certed workbook rows cannot become generated raw inventory rows.
- Notes / avoid porting: Preserve row-ref marking for true blank-cert raw rows, but hydrate certed row-ref rows before inventory/company record creation.

## Receive Reconcile Counts Raw Rows

- Origin: Windows
- Implemented on origin: Yes
- Mirrored to other platform: Mac pending
- Source commit/repo: `7050379 Fix received raw sheet inventory sync` in `mikegrossbarth/card_pipeline`
- Target to mirror: Mac LUCAS `mikegrossbarth/card_pipeline_mac`
- Summary: Home receive reconciliation now counts full row identities instead of cert numbers only. A row is accounted by cert when present, raw `Item ID` when present, or normalized card title as a fallback. This prevents a mixed raw sheet like `drose514_8_21_26.xlsx` from being moved to Received just because the Curry Gold row is in inventory while an Edwards Optic raw row is still unaccounted; it should remain partial until every raw/cert row is accounted. Manual/context-menu Received moves and automatic fully-received moves now pre-sync inventory before the hard block checks whether any workbook row is still missing from inventory, company sheets, or sold ledgers.
- Notes / avoid porting: Do not return `_reconcile_accounted_home_sheets` to cert-only completeness checks. Raw rows must participate in both the partial notice count and the fully received move decision. When a raw row has an `Item ID`, accounting/candidate checks must match that `Item ID` before title fallback; title fallback is only for rows with no cert and no raw `Item ID`, otherwise generated duplicate raw rows can hide the real workbook row. Received raw rows that are missing `Item ID` must have stable IDs written back into the workbook before inventory candidates are created; make sure `_ensure_raw_item_ids_in_sheet_paths` honors later header aliases like `Cert` and `Description`, not only the first alias in each list.

## People Rules Balance Share / Person Semantics

- Origin: Mac
- Implemented on origin: Yes
- Mirrored to other platform: Windows complete
- Source commit/repo: `691eee9 Add blank people rule sheet type option` in `mikegrossbarth/card_pipeline_mac`
- Target to mirror: Windows LUCAS `mikegrossbarth/card_pipeline`
- Windows source commit/repo: `ce14467 Sync Mac people rules balance share` in `mikegrossbarth/card_pipeline`
- Summary: People Rules now uses `Person`, supports blank `Sheet Type` rows with `Balance Share %` for team profit-share payouts, keeps Seller Rate/Deduction rows tied to active Sheet Types, disables/clears Balance Share when Sheet Type is filled, adds a blank Sheet Type dropdown option to reset rows, defaults blank seller min/max bounds to `$0` and `$1,000,000,000`, and prevents balance-share-only rows from making someone a seller source. Team payouts use each person's Balance Share rule when present and otherwise default to 50%.
- Notes / avoid porting: Do not port superseded semantics where balance-share-only rows classify a person as a seller, where team members are hard-coded only as half profit in the UI wording, or where the People Rules first column is shown only as `Seller`.

## Team Payout Rows / Expense Edit Stability

- Origin: Windows
- Implemented on origin: Yes
- Mirrored to other platform: Mac complete
- Source commit/repo: `24daa6e Fix team payout rows and expense edits` in `mikegrossbarth/card_pipeline`
- Target to mirror: Mac LUCAS `mikegrossbarth/card_pipeline_mac`
- Mac source commit/repo: `786ecd8 Sync team payout row fixes` in `mikegrossbarth/card_pipeline_mac`
- Summary: Team-member payouts now generate per sold card / per expense payout rows from the profit ledger, while network sellers remain sheet based. Legacy paid sheet markers no longer hide later General Sold card payouts without timestamps. Edit Expense now preserves previous ledger keys and stamps a stable expense id so older/no-id expenses can be edited repeatedly.
- Notes / avoid porting: Do not revert team payouts back to sheet-level grouping. Team member payout audits must be per profit-ledger item; seller/network payouts can remain sheet based.

## Profit Generated Metric

- Origin: Windows
- Implemented on origin: Yes
- Mirrored to other platform: Mac complete
- Source commit/repo: `94b31c8 Add desktop trade portal` in `mikegrossbarth/card_pipeline`
- Target to mirror: Mac LUCAS `mikegrossbarth/card_pipeline_mac`
- Mac source commit/repo: `9a180b3 Add generated profit metric` in `mikegrossbarth/card_pipeline_mac`
- Summary: Profit Metric dropdown now includes `Generated Profit`, which shows profit generated in each selected breakdown bucket rather than cumulative overall profit. Day/Month behavior is controlled by the existing Breakdown dropdown; `Overall Profit` remains the cumulative running-total view.
- Notes / avoid porting: Do not make `Generated Profit` cumulative. It should behave like per-bucket net profit and respect the same period/person/search filters as existing profit charts.

## Desktop Trade Portal

- Origin: Mac
- Implemented on origin: Yes
- Mirrored to other platform: Windows complete
- Source commit/repo: `6cb1779 Add desktop trade portal` plus layout refinements through `c079a28 Tighten trade incoming column layout` in `mikegrossbarth/card_pipeline_mac`
- Target to mirror: Windows LUCAS `mikegrossbarth/card_pipeline`
- Windows source commit/repo: `94b31c8 Add desktop trade portal` in `mikegrossbarth/card_pipeline`
- Summary: Windows Inventory now has an `Enter Trade Portal` action that opens the desktop trade workflow, lets users select outgoing active inventory cards, enter incoming trade cards, cash paid/received, and saves the trade by marking outgoing cards sold at basis with sale method `Trade` while adding incoming inventory with allocated purchase prices.
- Notes / avoid porting: Port only the desktop/local inventory trade behavior. Do not port mobile routes, offline queue, Cloudflare/tunnel behavior, Instagram sync, or Mac app bundle/launcher mechanics.

## Receive Missing Cert Fast Fail

- Origin: Windows
- Implemented on origin: Yes
- Mirrored to other platform: Mac complete
- Source commit/repo: `3424e17 Avoid receive index refresh on missing scans` in `mikegrossbarth/card_pipeline`
- Target to mirror: Mac LUCAS `mikegrossbarth/card_pipeline_mac`
- Mac source commit/repo: `0986fae Avoid receive index refresh on missing scans` in `mikegrossbarth/card_pipeline_mac`
- Summary: Receive barcode/title/raw lookup no longer synchronously rebuilds the full incoming index when a cert or search text is not found. Missing rows now fail fast against the current index and tell the user to click `Refresh Incoming Index` if they expected a match. Stale matched rows with missing assignment values can still refresh once to repair best company/payout.
- Notes / avoid porting: Do not bring back automatic full incoming/working sheet scans on every receive miss; that path caused timeouts when the card simply was not in the sheets.

## Receive Missing Cert Background Retry

- Origin: Windows
- Implemented on origin: Yes
- Mirrored to other platform: Mac complete
- Source commit/repo: `f2943ed Retry receive index misses in background` in `mikegrossbarth/card_pipeline`
- Target to mirror: Mac LUCAS `mikegrossbarth/card_pipeline_mac`
- Mac source commit/repo: `6057121 Retry receive index misses in background` in `mikegrossbarth/card_pipeline_mac`
- Summary: Receive rows that miss the current incoming index now show `CHECKING INDEX` and start a background incoming/working sheet index refresh. When the refresh returns, LUCAS retries unmatched rows on the main event loop and either attaches the real sheet or marks `NO SHEET FOUND` only after the retry also misses.
- Notes / avoid porting: Keep the retry off the immediate scan/Enter path. Do not return to blocking full index rebuilds before adding the row.

## Inventory Photo Refresh Should Stay Lightweight

- Origin: Windows
- Implemented on origin: Yes
- Mirrored to other platform: Mac complete
- Source commit/repo: `8606129 Keep photo refresh lightweight` in `mikegrossbarth/card_pipeline`
- Target to mirror: Mac LUCAS `mikegrossbarth/card_pipeline_mac`
- Mac source commit/repo: `1355e80 Keep inventory photo refresh lightweight` in `mikegrossbarth/card_pipeline_mac`
- Summary: Inventory photo scans that link photos now queue a lightweight inventory redraw instead of forcing `refresh_inventory_tab(enrich=True)` across every active inventory row. The event handler still defaults to enrichment for legacy callers, but accepts an `{"enrich": false}` payload for redraw-only refreshes.
- Notes / avoid porting: Do not make photo linking recompute best company / estimated payout for all inventory rows. Explicit `Update Payouts` and assignment refresh actions should remain the path for heavy enrichment.

## Card Ladder Helper Version Drift

- Origin: Windows
- Implemented on origin: Yes
- Mirrored to other platform: Mac verified; no change needed
- Source commit/repo: `aa0a145 Align Card Ladder helper version` in `mikegrossbarth/card_pipeline`
- Target to mirror: Mac LUCAS `mikegrossbarth/card_pipeline_mac`
- Mac source commit/repo: Mac already internally consistent with bundled helper version `2026-07-21-visible-cert-partial-v25`
- Summary: Windows bridge now expects the same bundled/loaded Card Ladder helper version it ships with, `2026-07-21-visible-cert-result-v24`, so extension polling is allowed to receive comp commands again. A regression test compares the bridge expectation to the bundled extension background helper constant.
- Notes / avoid porting: Do not bump Windows expected helper to `visible-cert-partial-v25` unless the Windows bundled Chrome extension and users' loaded helper are also bumped/reloaded. Do not reintroduce blind grader-option coordinate clicks.

## Card Ladder Generic Title Review Keeps Values

- Origin: Windows
- Implemented on origin: Yes
- Mirrored to other platform: Mac complete
- Source commit/repo: `9f78192 Preserve Card Ladder values for title review` in `mikegrossbarth/card_pipeline`
- Target to mirror: Mac LUCAS `mikegrossbarth/card_pipeline_mac`
- Mac source commit/repo: `9a2e2d4 Preserve Card Ladder values for title review` in `mikegrossbarth/card_pipeline_mac`
- Summary: Generic/broad Card Ladder profile titles such as `2018 Topps PSA 9` are still blocked from auto-filling the card title, but successful CL value and filtered comps are now preserved on the row with status `Card Ladder title review`. This avoids rows looking uncomped when Card Ladder returned money data but not a trustworthy card title.
- Notes / avoid porting: Do not go back to saving broad profile titles as card descriptions. The title should stay blank/manual-review unless the captured profile title has enough detail.

## Card Ladder Generic Title Settle Retry

- Origin: Windows
- Implemented on origin: Yes
- Mirrored to other platform: Mac complete
- Source commit/repo: `05d0d39 Retry generic Card Ladder title capture` in `mikegrossbarth/card_pipeline`
- Target to mirror: Mac LUCAS `mikegrossbarth/card_pipeline_mac`
- Mac source commit/repo: `30ecd32 Retry generic Card Ladder title capture` in `mikegrossbarth/card_pipeline_mac`
- Summary: The Chrome Card Ladder helper now waits about one extra second and recaptures DOM results when the first complete result has a generic/broad profile title. If the title resolves to a detailed profile, LUCAS uses the resolved result; if it stays generic, the existing `Card Ladder title review` behavior preserves values/comps but leaves the title blank.
- Notes / avoid porting: Keep this as a targeted recapture only for generic profile titles. Do not add long global sleeps to every row, and do not save broad titles as card descriptions.

## Inventory Sale Duplicate Key Uses Inventory Key

- Origin: Windows
- Implemented on origin: Yes
- Mirrored to other platform: Mac complete
- Source commit/repo: `34f2b6f Fix inventory sale duplicate matching` in `mikegrossbarth/card_pipeline`
- Target to mirror: Mac LUCAS `mikegrossbarth/card_pipeline_mac`
- Mac source commit/repo: Verified present on Mac `main` at `2fe9837 Integrate inventory person selection UI updates`
- Summary: Sold-from-inventory profit rows now preserve the hidden `inventory_key` and use it as the true de-dupe identity. This prevents a new inventory sale from being swallowed when a user fixes/swaps cert numbers and the corrected cert already exists on an older sold row. Company-sheet recovery still uses the weaker company/source/cert match only to avoid double-counting an inventory sale recovered from weekly company sheets.
- Notes / avoid porting: Do not return to using only `cert + company + source sheet` as the identity for sold-from-inventory rows. That weak key is only safe for recovery/backfill de-dupe, not fresh inventory sales.

## Team/Personal Card Ladder Bridge Coexistence

- Origin: Windows
- Implemented on origin: Yes
- Mirrored to other platform: Mac complete
- Source commit/repo: `e3beebf Let team and personal bridge coexist` in `mikegrossbarth/card_pipeline`
- Target to mirror: Mac LUCAS `mikegrossbarth/card_pipeline_mac`
- Mac source commit/repo: Verified present on Mac `main` at `2fe9837 Integrate inventory person selection UI updates`
- Summary: Windows Team LUCAS keeps default local bridge port `8765`, while Personal/Michael LUCAS defaults to `8766` when no explicit `LUCAS_MOBILE_PORT` or `mobile_port` setting is supplied. The desktop bridge now allows fallback across the helper's known port range, so launching both profiles no longer forces one bridge to fail just because the other profile is open.
- Notes / avoid porting: Explicit port settings must still win. Do not make both Team and Personal compete for only `8765`; Chrome's helper already polls the known bridge port range.

## Manual Paid Payout History Without Profit Ledger Rows

- Origin: Windows
- Implemented on origin: Yes
- Mirrored to other platform: Mac complete
- Source commit/repo: `a1c7cb5 Keep manual payout history out of profit` in `mikegrossbarth/card_pipeline`
- Target to mirror: Mac LUCAS `mikegrossbarth/card_pipeline_mac`
- Mac source commit/repo: Verified present on Mac `main` at `2fe9837 Integrate inventory person selection UI updates`
- Summary: Payout history can include explicit manual paid-payout markers using `manual_paid_adjustment` and `manual_paid_amount`, allowing a historical paid amount to show in the payout history popup without creating a fake sale/profit row in `profit_ledger.json`.
- Notes / avoid porting: Do not model manual historical payouts as sold cards, sale price, or profit. They should affect paid-history display only, not profit totals or generated profit graphs.

## Create Tab Value Totals Row

- Origin: Windows/Mac shared
- Implemented on origin: Yes
- Mirrored to other platform: Complete
- Source commit/repo: Windows `6f43460 Add create tab totals row` in `mikegrossbarth/card_pipeline`; Mac `16d3e3c Add create tab totals row` in `mikegrossbarth/card_pipeline_mac`
- Target to mirror: N/A
- Summary: Create now uses the same value-total row helper as Comp, adding a non-editable bottom `TOTAL` row that sums Purchase, Card Ladder, Comps, CY Estimate, and Est. Payout for visible columns.
- Notes / avoid porting: Keep this as a display-only tree row. Do not write the totals row into source workbooks or let it participate in row edits/deletes.
## Inventory Bulk Edit Toggle Restored

- Origin: Windows
- Implemented on origin: Yes
- Mirrored to other platform: Mac note only unless explicitly requested
- Source commit/repo: `03dc502 Restore inventory bulk edit toggle` in `mikegrossbarth/card_pipeline`
- Target to mirror: Mac LUCAS `mikegrossbarth/card_pipeline_mac` only if a human explicitly asks for the Mac implementation
- Summary: Inventory bulk editing is intentionally opt-in again. The visible `Bulk Edit` toggle controls arrow-key cell navigation and Enter/F2 editing. When Bulk Edit is off, normal Inventory editing uses double-click-to-edit cells, preserving the older per-cell edit flow. Person edits remain restricted to known people through a readonly dropdown and validation.
- Notes / avoid porting: The earlier direct-cell-edit cleanup note is superseded. Do not remove the visible Bulk Edit toggle from Windows.
