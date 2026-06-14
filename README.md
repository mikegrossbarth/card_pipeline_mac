# L.U.C.A.S for macOS

Lot Upload, Comping & Assignment System.

This is the macOS project copy of L.U.C.A.S. It is intentionally separate from the original Windows project so the Mac launch, install, and setup flow can evolve without disturbing the production Windows app.

L.U.C.A.S is a desktop workflow app for card intake, receiving, working-sheet tracking, Card Ladder comping, assignment routing, payouts, and profit review.

For a click-by-click setup walkthrough for a brand-new Mac, start with [FIRST_RUN_SETUP.md](FIRST_RUN_SETUP.md).

## Install

1. Install Google Chrome.
2. Install Python 3.11 or newer for macOS. The python.org installer is the easiest path because it includes Tkinter.
3. Download or clone this project.
4. Open Terminal in the project folder.
5. Run:

```bash
chmod +x install_dependencies.sh run_card_pipeline.sh "Run Card Pipeline.command" create_macos_app.sh
./install_dependencies.sh
```

6. Open `.env`, which the installer creates from `.env.example`.
7. Add `GOOGLE_API_KEY`, `GOOGLE_SHEETS_OAUTH_CLIENT_ID`, `GOOGLE_SHEETS_OAUTH_CLIENT_SECRET`, and `LUCAS_WORKING_SHEETS_DIR`.
8. Launch with:

```bash
./run_card_pipeline.sh
```

You can also double-click `Run Card Pipeline.command` after the file has execute permission. To create a Finder-friendly launcher, run:

```bash
./create_macos_app.sh
```

That creates `LUCAS.app` in the project folder. You can keep it there or copy it to your Desktop; it launches this project path and local `.venv`. If launch fails, check `~/Desktop/LUCAS-launch.log`.

## Local Configuration

`.env` is intentionally local and should not be committed. A typical Mac setup with Google Drive for desktop looks like:

```env
GOOGLE_API_KEY=your_google_ai_studio_key
GOOGLE_SHEETS_OAUTH_CLIENT_ID=your_desktop_oauth_client_id
GOOGLE_SHEETS_OAUTH_CLIENT_SECRET=your_desktop_oauth_client_secret
LUCAS_WORKING_SHEETS_DIR=/Users/yourname/Library/CloudStorage/GoogleDrive-your.email@gmail.com/My Drive/CARD_PIPELINE/WORKING SHEETS
```

`LUCAS_WORKING_SHEETS_DIR` should point at the `WORKING SHEETS` folder. The app uses that folder's parent as the pipeline root, so `INCOMING SHEETS`, `RECEIVED SHEETS`, `COMPANY SHEETS`, `ASSIGNMENT RULES`, markers, locks, and profit records live beside it.

If you keep the pipeline outside Google Drive, use the same folder shape anywhere macOS can read and write.

## Data Folder

The expected shared data folder shape is:

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

On first run, click `Working Folder` in the top-right header and choose the actual `WORKING SHEETS` folder. The choice is saved locally in `lucas_settings.json`.

## Card Ladder Extension

Card Ladder comping requires a Card Ladder account and an active Chrome login session.

1. Open Chrome and go to `chrome://extensions`.
2. Turn on `Developer mode`.
3. Click `Load unpacked`.
4. Select `cardladder-autocomp/extension` from this project.
5. Log into Card Ladder in Chrome before running comps.

The app starts the local Card Ladder bridge automatically when L.U.C.A.S opens. The extension talks to `127.0.0.1` ports `8765` through `8772`, which works the same on macOS.

## Input Modes

Use the `Create` tab for all card entry.

- `Barcode Scanner`: scanning station mode for continuous cert entry.
- `Photo OCR`: add photos or a folder, scan them in the app, and append detected card rows.
- `Existing Spreadsheet`: load a simple workbook where column 1 is cert number, column 2 is card description, and column 3 is purchase price.

Enter a title, then click `Save as Working Sheet`.

## Comping

Use the `Comp` tab for comping. Select a saved sheet, choose whether to run `Card Ladder + CY`, `Card Ladder`, or `CY`, choose the comp method and run scope, then click `Run All Comps`.

The app stores Card Ladder value, comps, assignment, payout, and status in the active workbook output. Rows marked `invalid_cert` are skipped by empty-comps-only runs.

On macOS, L.U.C.A.S can submit certs to the local CourtYard app and fill `CY value`. This uses the bundled macOS CourtYard automation in `comp_engine/cy_automation` plus `scripts/macos/cgscroll`. It expects `CYCardScanner`, `cliclick`, and `tesseract` to be installed locally. When a CY batch finishes, L.U.C.A.S quits CourtYard. Disable the lookup with `LUCAS_DISABLE_CY_LOOKUP=1`.

## Assignment

Use the `Receive` tab for physically receiving cards and source matching. Use the `Assignment` tab for pure assignment review and fallback assignment work.

Assignment companies are local in `assignment_companies.json`. The manager supports manual rules, local files, Google Keep exports, workbook/CSV files, Google Sheets through OAuth, manual payout tiers, payout files, and linked `Payouts` tabs.

The company list in Assignment Rules can be filtered by name and by All, Active, or Inactive status.

Example Mac source paths:

```json
{
  "name": "Arena Club",
  "value_source": "comps",
  "rules": "/Users/yourname/Library/CloudStorage/GoogleDrive-your.email@gmail.com/My Drive/CARD_PIPELINE/ASSIGNMENT RULES/arena-club-rules.xlsx",
  "payout": "/Users/yourname/Library/CloudStorage/GoogleDrive-your.email@gmail.com/My Drive/CARD_PIPELINE/ASSIGNMENT RULES/arena-club-payout.xlsx"
}
```

If no company can take a priced row, `Best Company` shows `NOBODY TAKES`.

## Payouts And Profit

Use `Payouts/Tabs` to track active balances by assigned person and mark person-level balances paid.

Use `Profit` to review sold cards and sold sheets. The Profit tab can filter by assigned person, shows a daily profit line chart, and can toggle between individual sold-card rows and grouped sold-sheet summaries.

## Tests

Run the committed offline test suite with:

```bash
.venv/bin/python -m unittest discover -s tests -v
```

Useful sanity checks:

```bash
.venv/bin/python -m compileall -q .
.venv/bin/python -c "import app; root = app.CardPipelineApp(); root.update_idletasks(); root.destroy(); print('app startup ok')"
```

## Local Files

Do not commit these:

```text
.env
.venv
lucas_settings.json
lucas_user_identity.json
lucas_google_sheets_token.json
assignment_companies.json
work/
outputs/
```

## Troubleshooting

If the app does not open, run `./run_card_pipeline.sh` from Terminal so macOS keeps the error visible.

If `import tkinter` fails, install a Python build with Tkinter. The python.org macOS installer is usually the simplest fix. With Homebrew Python, install `python-tk`.

If Chrome blocks the extension or it does not check in, make sure the unpacked extension is loaded from this Mac project copy, Chrome is open, and Card Ladder is logged in.
