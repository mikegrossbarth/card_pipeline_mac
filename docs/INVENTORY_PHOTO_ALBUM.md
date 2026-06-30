# LUCAS Inventory Photo Album

This is the LUCAS version of the earlier automatic inventory update flow. It uses only phone capture plus synced folders. It does not use Instagram, Cloudflare R2, public posting, or the private iPhone Photos library as the source of truth.

## Flow

1. Use a Shortcut that takes inventory photos and saves the files directly into a synced folder such as iCloud Drive, Google Drive, or `Shortcuts/LUCAS Inventory Photos`.
2. In LUCAS, click `Inventory -> Photo Folder` and choose that exact synced folder.
3. Click `Inventory -> Scan Photos`. If the selected folder is private/iCloud/phone-synced, LUCAS first copies new or changed images into shared `CARD_PIPELINE/INVENTORY PHOTOS`, then scans that shared folder.
4. When cards are sold/deleted/moved, LUCAS deletes the linked synced files it can see. With the file-based Shortcut, that deletion can sync back to the phone-side folder.
5. LUCAS also repeats that same mirror-and-scan flow every three hours.
6. LUCAS OCRs cert numbers from new/changed photos and links matching certs to active rows in `inventory_ledger.json`.
7. Inventory shows a `Photos` count and exports `Photos` plus `Photo Paths`.
8. When a card leaves active inventory by being sold, deleted, or moved to a company sheet, LUCAS deletes every linked source/shared photo file it can safely resolve unless another active inventory row still uses that same photo.

## Viewing Photos In LUCAS

After a scan links a photo, the matching Inventory row shows a number in the `Photos` column.

To open it:

1. Right-click the Inventory row.
2. Click `Open Photo`.

You can also click `Open Photo Folder` from the same right-click menu to jump to the synced album folder.

While copying and scanning, the Inventory status line updates with progress like `Mirroring inventory photos: 3/42 IMG_1234.jpg` and `Inventory photo scan: 3/42 IMG_1234.jpg`. Photos already linked to active inventory rows are skipped during later scans.

## Folder

By default LUCAS uses:

```text
CARD_PIPELINE/INVENTORY PHOTOS
```

Create a phone-side Shortcut that saves inventory photos as files into that folder through iCloud Drive, Google Drive, or another synced folder path. If you need a different folder, set `inventory_photo_folder` in `lucas_settings.json`.

The easier way is to click `Inventory -> Photo Folder` and pick the folder visually. LUCAS saves that choice in `lucas_settings.json` and shows how many photo files it found.

## Team Sharing

Use the phone/iCloud/Drive folder as your private source folder, then click `Inventory -> Scan Photos`. LUCAS copies the current source photos into:

```text
CARD_PIPELINE/INVENTORY PHOTOS
```

That folder lives under the shared pipeline, so Google Drive can sync it to everyone else. Scan Photos scans the shared folder after copying, which means linked inventory rows point at shared photo paths instead of a private iCloud folder.

New links are saved as portable paths relative to `CARD_PIPELINE/INVENTORY PHOTOS` whenever possible. LUCAS still reads older absolute links and remaps them into the current shared folder if the path contains an `INVENTORY PHOTOS` folder segment.

Other users only need to pull the latest LUCAS, point their shared pipeline/working folder at the same team `CARD_PIPELINE`, then use `Inventory -> Scan Photos` or right-click linked rows with `Open Photo`.

## Preferred Shortcut Shape

Use this shape when you want LUCAS to be able to delete the photo file later:

1. `Take Photo`
   - Show Camera Preview: on.
   - Use Front Camera: off.
2. `Save File`
   - File: photo from the previous action.
   - Destination: the synced export folder.
   - Ask Where to Save: off.
   - Overwrite If File Exists: on if available.

This keeps the workflow file-based. The photo file lives in the synced folder, and LUCAS can delete that file when the card leaves inventory.

## Legacy Photos Album Shape

The older shape still works, but it leaves originals in the iPhone Photos library:

1. `Find Photos`
   - Album is `LUCAS Inventory`.
2. `Save File`
   - File: photos from the previous action.
   - Destination: the synced export folder.
   - Ask Where to Save: off.
   - Overwrite If File Exists: on if available.

LUCAS delete/sold/move cleanup can remove files from the synced source folder and shared `CARD_PIPELINE/INVENTORY PHOTOS` folder. It cannot directly delete a picture from the private iPhone Photos library unless your Shortcut is built around Files/synced folders instead of keeping the original in Photos.

## Notes

- LUCAS only links photos to active inventory rows with cert numbers.
- Raw cards without cert numbers are not auto-linked by this scanner.
- HEIC files are supported through `pillow-heif`, which is installed by `install_dependencies`. If HEIC scans fail on an older install, rerun `install_dependencies` or change the Shortcut to export JPEG files.
- The scan uses `GOOGLE_API_KEY` and the existing LUCAS photo OCR dependencies.
