# LUCAS Inventory Photo Album

This is the LUCAS version of the earlier automatic inventory update flow. It uses only the phone album and synced-folder half. It does not use Instagram, Cloudflare R2, or public posting.

## Flow

1. Keep active inventory photos in an iPhone Photos album named `LUCAS Inventory`.
2. A Shortcut exports that album to a synced folder on the computer.
3. In LUCAS, click `Inventory -> Photo Folder` and choose that exact synced folder.
4. LUCAS scans the synced folder every three hours and when you click `Inventory -> Scan Photos`.
5. LUCAS OCRs cert numbers from new/changed photos and links matching certs to active rows in `inventory_ledger.json`.
6. Inventory shows a `Photos` count and exports `Photos` plus `Photo Paths`.
7. When a card leaves active inventory by being sold, deleted, or moved to a company sheet, LUCAS deletes its linked photo from the active photo folder unless another active inventory row still uses that same photo.

## Viewing Photos In LUCAS

After a scan links a photo, the matching Inventory row shows a number in the `Photos` column.

To open it:

1. Right-click the Inventory row.
2. Click `Open Photo`.

You can also click `Open Photo Folder` from the same right-click menu to jump to the synced album folder.

While scanning, the Inventory status line updates with progress like `Inventory photo scan: 3/42 IMG_1234.jpg`.

## Folder

By default LUCAS uses:

```text
CARD_PIPELINE/INVENTORY PHOTOS
```

Create a phone-side Shortcut that exports the `LUCAS Inventory` album into that folder through iCloud Drive, Google Drive, or another synced folder path. If you need a different folder, set `inventory_photo_folder` in `lucas_settings.json`.

The easier way is to click `Inventory -> Photo Folder` and pick the folder visually. LUCAS saves that choice in `lucas_settings.json` and shows how many photo files it found.

## Shortcut Shape

Safe iPhone Shortcut:

1. `Find Photos`
   - Album is `LUCAS Inventory`.
2. `Save File`
   - File: photos from the previous action.
   - Destination: the synced export folder.
   - Ask Where to Save: off.
   - Overwrite If File Exists: on if available.

Do not add `Delete Photos`; that can delete from the iPhone photo library. If cleanup is needed later, delete files from the export folder only, not Photos results.

## Notes

- LUCAS only links photos to active inventory rows with cert numbers.
- Raw cards without cert numbers are not auto-linked by this scanner.
- HEIC files are supported only if the local Python environment can open them with `pillow-heif`. JPEG export from the Shortcut is the safest option.
- The scan uses `GOOGLE_API_KEY` and the existing LUCAS photo OCR dependencies.
