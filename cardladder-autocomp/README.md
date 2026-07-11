# Card Ladder Auto-Comp Helper

This folder contains the Chrome extension used by L.U.C.A.S for Card Ladder comping.

Most users should run comps from the L.U.C.A.S `Comp` tab. The app starts the local bridge and sends queued rows to the extension.

## Chrome Setup

1. Open Chrome and go to `chrome://extensions`.
2. Turn on `Developer mode`.
3. Click `Load unpacked`.
4. Select this folder: `cardladder-autocomp/extension`.
5. Log into Card Ladder in the same Chrome profile.

Card Ladder requires an active account session. If Chrome is not logged in, the extension can load but the run will not complete.

The main L.U.C.A.S app uses the project `.env` for OCR fallback. The Card Ladder extension itself does not need the Google API key, but screenshot OCR fallback in the app does.

The extension talks to the desktop bridge on `127.0.0.1` ports `8765` through `8772`. No Node/npm install or build step is required for normal use.

Current helper version expected by the app bridge:

```text
2026-07-11-stale-first-row-retry-v23
```

For non-PSA graders, the helper first tries normal page/DOM clicks. If Card Ladder ignores those synthetic clicks, it briefly uses Chrome's `debugger` API to send a trusted click to the visible Grader bar. Chrome may show a browser-level debugger warning during that fallback. The debugger attach is scoped to the Card Ladder tab and is detached after the click.

## Files

- `extension/background.js`: opens and controls the Card Ladder run window.
- `extension/content.js`: interacts with Card Ladder pages and reads values/comps.
- `extension/popup.*`: provides manual queue loading, run status, and result download tools.
