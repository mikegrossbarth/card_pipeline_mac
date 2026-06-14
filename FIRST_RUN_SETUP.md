# L.U.C.A.S First Run Setup for macOS

This guide assumes you are setting up L.U.C.A.S on a Mac for the first time.

L.U.C.A.S stands for Lot Upload, Comping & Assignment System. It helps you create card sheets, comp cards through Card Ladder, receive cards, assign cards to buying companies, and track payouts.

## Quick Answer

For a complete Mac setup, the computer needs:

- macOS with Terminal access
- Google Chrome
- Python 3.11 or newer with Tkinter
- this L.U.C.A.S Mac project folder
- a local `.venv` created by `install_dependencies.sh`
- a data folder with `WORKING SHEETS`, `INCOMING SHEETS`, and `RECEIVED SHEETS`
- a Card Ladder account and the unpacked Chrome extension
- a Google AI Studio API key for Photo OCR and screenshot OCR fallback
- Google billing and spend controls for Google API calls
- Google OAuth credentials for Google Sheets rules or payout sheets
- Google Drive for desktop if the team uses a shared Drive folder
- active assignment companies with rule and payout sources

## Important Setup URLs

| Need | URL |
| --- | --- |
| L.U.C.A.S repository | `https://github.com/mikegrossbarth/card_pipeline` |
| Google Chrome download | `https://www.google.com/chrome/` |
| Python for macOS | `https://www.python.org/downloads/macos/` |
| Homebrew | `https://brew.sh/` |
| Google Drive for desktop | `https://www.google.com/drive/download/` |
| Card Ladder | `https://app.cardladder.com/` |
| Chrome extensions page | `chrome://extensions` |
| Google AI Studio API keys | `https://aistudio.google.com/app/apikey` |
| Gemini API billing | `https://ai.google.dev/gemini-api/docs/billing` |
| Gemini API key security | `https://ai.google.dev/gemini-api/docs/api-key` |
| Google AI Studio spend cap | `https://aistudio.google.com/app/spend` |
| Google Cloud Console | `https://console.cloud.google.com/` |
| Google Cloud Billing | `https://console.cloud.google.com/billing` |
| Google Sheets API | `https://console.cloud.google.com/apis/library/sheets.googleapis.com` |
| Google Auth Platform | `https://console.cloud.google.com/auth` |
| Google OAuth clients | `https://console.cloud.google.com/auth/clients` |

## Step 1: Get The App Folder

Put the L.U.C.A.S Mac project folder somewhere normal, for example:

```text
/Users/yourname/Applications/card-pipeline-mac
```

or:

```text
/Users/yourname/Documents/card-pipeline-mac
```

The folder should contain:

```text
app.py
install_dependencies.sh
run_card_pipeline.sh
Run Card Pipeline.command
README.md
cardladder-autocomp
```

## Step 2: Install Google Chrome

Install Google Chrome from:

```text
https://www.google.com/chrome/
```

Chrome is required for Card Ladder automation because the app talks to a local Chrome extension.

## Step 3: Install Python

Install Python 3.11 or newer for macOS.

Recommended path:

```text
https://www.python.org/downloads/macos/
```

The python.org installer is recommended because it normally includes Tkinter, which L.U.C.A.S needs for the desktop UI.

Homebrew can also work:

```bash
brew install python
brew install python-tk
```

## Step 4: Run The Installer

Open Terminal in the project folder and run:

```bash
chmod +x install_dependencies.sh run_card_pipeline.sh "Run Card Pipeline.command" create_macos_app.sh
./install_dependencies.sh
```

The installer will:

1. find Python 3.11 or newer
2. create `.venv`
3. install dependencies from `requirements.txt`
4. verify Tkinter is available
5. create `.env` from `.env.example` if `.env` does not exist

If Tkinter is missing, install the python.org Python build or install tkinter support for your Homebrew Python.

## Step 5: Create The Sheet Folders

Create a main folder for pipeline data.

For Google Drive for desktop, Mac paths usually look like:

```text
/Users/yourname/Library/CloudStorage/GoogleDrive-your.email@gmail.com/My Drive/CARD_PIPELINE
```

Inside it, create:

```text
WORKING SHEETS
INCOMING SHEETS
RECEIVED SHEETS
```

L.U.C.A.S will also use or create:

```text
COMPANY SHEETS
ASSIGNMENT RULES
sheet_markers.json
profit_ledger.json
unassigned_players.json
assignment_player_overrides.json
.locks
```

## Step 6: Launch The App

From Terminal:

```bash
./run_card_pipeline.sh
```

After permissions are set, you can also double-click:

```text
Run Card Pipeline.command
```

To create a Finder app launcher:

```bash
./create_macos_app.sh
```

Then double-click `LUCAS.app`. Keep `LUCAS.app` beside the project folder; it launches the local project and `.venv`.

## Step 7: Choose The Working Folder

When L.U.C.A.S opens:

1. click `Working Folder`
2. choose the `WORKING SHEETS` folder
3. confirm the app sees your working sheets

Choose `WORKING SHEETS`, not the parent `CARD_PIPELINE` folder. The app remembers this in `lucas_settings.json`.

## Step 8: Configure `.env`

Open `.env` in the project folder.

A typical full setup looks like:

```env
GOOGLE_API_KEY=your_google_ai_studio_key
GOOGLE_SHEETS_OAUTH_CLIENT_ID=your_desktop_oauth_client_id
GOOGLE_SHEETS_OAUTH_CLIENT_SECRET=your_desktop_oauth_client_secret
LUCAS_WORKING_SHEETS_DIR=/Users/yourname/Library/CloudStorage/GoogleDrive-your.email@gmail.com/My Drive/CARD_PIPELINE/WORKING SHEETS
```

Do not commit `.env`.

## Step 9: Set Up Photo OCR And OCR Fallback

Photo OCR and Card Ladder screenshot OCR fallback need:

```env
GOOGLE_API_KEY=...
```

Create or copy the key at:

```text
https://aistudio.google.com/app/apikey
```

Then save `.env` and restart L.U.C.A.S.

## Step 10: Set Up Google Billing And Spend Controls

Use the same Google account and project that created the `GOOGLE_API_KEY`.

Set up billing, payment method, spend cap, and billing budget alerts through:

```text
https://aistudio.google.com/app/spend
https://console.cloud.google.com/billing
```

Keep the API key only in `.env`. If the key may have leaked, rotate it in Google AI Studio and update `.env`.

## Step 11: Set Up Card Ladder

1. Open Chrome.
2. Go to `chrome://extensions`.
3. Turn on `Developer mode`.
4. Click `Load unpacked`.
5. Select:

```text
cardladder-autocomp/extension
```

6. Log into Card Ladder in the same Chrome profile.

The desktop app starts the local Card Ladder bridge automatically. If comping does not start, check that Chrome is open, the extension is enabled, Card Ladder is logged in, and L.U.C.A.S says the bridge is running.

## Step 12: Set Up Google Sheets Access

Google Sheets access is required for Google Sheet rule or payout sources.

1. Open Google Cloud Console.
2. Enable the Google Sheets API.
3. Open Google Auth Platform.
4. Create an OAuth Client ID.
5. Choose application type `Desktop app`.
6. Copy the client ID and secret into `.env`.
7. Restart L.U.C.A.S.
8. Open Assignment Rules.
9. Click `Connect Google`.
10. Sign in with the Google account that can access the sheets.

The app creates `lucas_google_sheets_token.json`. Do not commit that file.

## Step 13: Set Up Google Drive For Desktop

Install Google Drive for desktop if the team uses a shared Drive folder.

Download:

```text
https://www.google.com/drive/download/
```

On modern macOS, Drive usually appears under:

```text
/Users/yourname/Library/CloudStorage/
```

Use Finder to locate the exact folder, then copy the `WORKING SHEETS` path into `.env` or choose it with the app's `Working Folder` button.

## Step 14: Set Up Assignment Companies

In L.U.C.A.S:

1. open the `Assignment` tab
2. click `Assignment Rules`
3. create a company
4. choose whether the company is active
5. choose a rule source
6. choose a payout source
7. save the company

The app saves company setup locally in `assignment_companies.json`.

## Step 15: Check The Full Workflow

Test in this order:

1. Create a small working sheet in `Create`.
2. Save it to `WORKING SHEETS`.
3. Open `Comp` and make sure the sheet appears.
4. Run one Card Ladder comp as a test.
5. Open Assignment Rules and make sure companies load.
6. Open `Assignment` and confirm best company and estimated payout can populate.
7. Open `Receive` and test marking a row received.
8. Open `Payouts/Tabs`.
9. Open `Profit` after company sheets have sold rows.

Do not start with a giant sheet until this small test works.

## Setup Checklist

- [ ] Google Chrome is installed
- [ ] Python 3.11+ with Tkinter is installed
- [ ] `./install_dependencies.sh` completed successfully
- [ ] `.env` exists
- [ ] `WORKING SHEETS`, `INCOMING SHEETS`, and `RECEIVED SHEETS` exist
- [ ] L.U.C.A.S opens
- [ ] `Working Folder` points to `WORKING SHEETS`
- [ ] `GOOGLE_API_KEY` is added
- [ ] Google billing, spend cap, and alerts are set
- [ ] Card Ladder extension is loaded
- [ ] user is logged into Card Ladder in Chrome
- [ ] Google OAuth credentials are added
- [ ] `Connect Google` has been completed
- [ ] assignment companies are created and active
- [ ] one small test sheet works end to end
