# LUCAS Inventory Photo Album

This is the LUCAS version of the earlier automatic inventory update flow. It uses only the phone album and synced-folder half. It does not use Instagram, Cloudflare R2, or public posting.

## Flow

1. Keep active inventory photos in an iPhone Photos album named `LUCAS Inventory`.
2. A Shortcut exports that album to a synced folder on the computer.
3. LUCAS scans the synced folder every three hours and when you click `Inventory -> Scan Photos`.
4. LUCAS OCRs cert numbers from new/changed photos and links matching certs to active rows in `inventory_ledger.json`.
5. Inventory shows a `Photos` count and exports `Photos` plus `Photo Paths`.
6. When a card leaves active inventory by being sold, deleted, or moved to a company sheet, LUCAS deletes its linked photo from the active photo folder unless another active inventory row still uses that same photo.

## Folder

By default LUCAS uses:

```text
CARD_PIPELINE/INVENTORY PHOTOS
```

Create a phone-side Shortcut that exports the `LUCAS Inventory` album into that folder through iCloud Drive, Google Drive, or another synced folder path. If you need a different folder, set `inventory_photo_folder` in `lucas_settings.json`.

## Shortcut Shape

Suggested iPhone Shortcut:

1. `Get Contents of Folder`
   - Folder: the synced `LUCAS Inventory` export folder.
2. `Delete Files`
   - Delete the files returned by the previous action.
   - Turn off confirmation if iOS allows it.
3. `Find Photos`
   - Album is `LUCAS Inventory`.
4. `Repeat with Each`.
5. `Save File`
   - File: repeat item.
   - Destination: the synced export folder.
   - Ask Where to Save: off.
   - Overwrite If File Exists: on if available.

The delete-then-save start matters because it makes the synced folder match the phone album. Photos removed from the album should disappear from the synced folder on the next Shortcut run.

## Notes

- LUCAS only links photos to active inventory rows with cert numbers.
- Raw cards without cert numbers are not auto-linked by this scanner.
- HEIC files are supported only if the local Python environment can open them with `pillow-heif`. JPEG export from the Shortcut is the safest option.
- The scan uses `GOOGLE_API_KEY` and the existing LUCAS photo OCR dependencies.
