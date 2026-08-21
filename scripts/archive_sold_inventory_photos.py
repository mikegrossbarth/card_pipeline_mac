#!/usr/bin/env python3
"""Archive sold inventory photos that are still sitting in INVENTORY PHOTOS.

Default mode is a read-only dry run:
    python scripts/archive_sold_inventory_photos.py --profile personal

Apply mode moves files and updates inventory_photo_state.json:
    python scripts/archive_sold_inventory_photos.py --profile personal --apply
"""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import shutil
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any


PERSONAL_ROOT = Path(
    "/Users/michaelgrossbarth/Library/CloudStorage/GoogleDrive-mikegrossbarth@gmail.com/My Drive/LUCAS_PERSONAL"
)
TEAM_ROOT = Path(
    "/Users/michaelgrossbarth/Library/CloudStorage/GoogleDrive-mikegrossbarth@gmail.com/My Drive/CARD_PIPELINE"
)
IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".webp", ".heic", ".heif"}
RETENTION_DAYS = 14


def load_json(path: Path, fallback: Any) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return fallback


def save_json(path: Path, data: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(path.suffix + ".tmp")
    tmp.write_text(json.dumps(data, indent=2, sort_keys=True), encoding="utf-8")
    tmp.replace(path)


def rows_from_ledger(raw: Any) -> list[dict[str, Any]]:
    if isinstance(raw, list):
        return [row for row in raw if isinstance(row, dict)]
    if isinstance(raw, dict):
        for key in ("items", "rows", "records", "ledger", "profit", "inventory"):
            value = raw.get(key)
            if isinstance(value, list):
                return [row for row in value if isinstance(row, dict)]
    return []


def cert(value: Any) -> str:
    return re.sub(r"\D", "", str(value or ""))


def file_hash(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def path_keys(value: Any, root: Path) -> set[str]:
    text = str(value or "").strip()
    if not text:
        return set()
    photo_root = root / "INVENTORY PHOTOS"
    path = Path(text).expanduser()
    keys = {text, path.name}
    if path.is_absolute():
        keys.add(str(path))
        try:
            keys.add(str(path.resolve()))
        except Exception:
            pass
        try:
            keys.add(path.relative_to(photo_root).as_posix())
        except Exception:
            pass
    else:
        keys.add(str(photo_root / path))
    return {key for key in keys if key}


def photo_keys(path: Path, root: Path) -> set[str]:
    photo_root = root / "INVENTORY PHOTOS"
    keys = {str(path), path.name}
    try:
        keys.add(str(path.resolve()))
    except Exception:
        pass
    try:
        keys.add(path.relative_to(photo_root).as_posix())
    except Exception:
        pass
    return {key for key in keys if key}


def unique_archive_path(dest_dir: Path, filename: str) -> Path:
    candidate = dest_dir / filename
    if not candidate.exists():
        return candidate
    stem = candidate.stem
    suffix = candidate.suffix
    index = 2
    while True:
        candidate = dest_dir / f"{stem}-{index}{suffix}"
        if not candidate.exists():
            return candidate
        index += 1


def state_record_exists_in_album(record: dict[str, Any], root: Path) -> bool:
    photo_root = root / "INVENTORY PHOTOS"
    for key in ("path", "relative_path", "filename"):
        value = str(record.get(key) or "").strip()
        if not value:
            continue
        path = Path(value).expanduser()
        candidates = [path]
        if not path.is_absolute():
            candidates.append(photo_root / path)
        for candidate in candidates:
            if candidate.exists() and candidate.is_file():
                return True
    return False


def archive_profile(root: Path, apply: bool) -> dict[str, Any]:
    inventory = rows_from_ledger(load_json(root / "inventory_ledger.json", []))
    profit = rows_from_ledger(load_json(root / "profit_ledger.json", []))
    state_path = root / "inventory_photo_state.json"
    state = load_json(state_path, {"version": 1, "photos": {}})
    if not isinstance(state, dict):
        state = {"version": 1, "photos": {}}
    photos = state.setdefault("photos", {})
    if not isinstance(photos, dict):
        photos = {}
        state["photos"] = photos

    active_rows = [row for row in inventory if str(row.get("status") or "").strip().lower() == "active"]
    active_certs = {cert(row.get("cert_number")) for row in active_rows if cert(row.get("cert_number"))}
    sold_certs = {cert(row.get("cert_number")) for row in profit if cert(row.get("cert_number"))}

    active_photo_keys: set[str] = set()
    for row in active_rows:
        for value in row.get("photo_paths") or []:
            active_photo_keys.update(path_keys(value, root))

    sold_photo_keys: set[str] = set()
    sold_photo_details: dict[str, dict[str, Any]] = {}
    for row in profit:
        for value in row.get("photo_paths") or []:
            keys = path_keys(value, root)
            sold_photo_keys.update(keys)
            for key in keys:
                sold_photo_details[key] = row

    photo_root = root / "INVENTORY PHOTOS"
    files = [
        path
        for path in photo_root.rglob("*")
        if path.is_file() and path.suffix.lower() in IMAGE_EXTENSIONS
    ] if photo_root.exists() else []

    candidates: dict[Path, dict[str, Any]] = {}
    for path in files:
        keys = photo_keys(path, root)
        if keys & active_photo_keys:
            continue
        matched_sold_path = keys & sold_photo_keys
        if matched_sold_path:
            detail_key = next(iter(matched_sold_path))
            candidates[path] = {
                "reason": "sold_photo_path",
                "profit_row": sold_photo_details.get(detail_key, {}),
            }
            continue
        try:
            sha = file_hash(path)
        except Exception:
            continue
        record = photos.get(sha)
        if not isinstance(record, dict):
            continue
        record_certs = {cert(value) for value in record.get("certs") or [] if cert(value)}
        if record_certs and record_certs & sold_certs and not record_certs & active_certs:
            candidates[path] = {
                "reason": "sold_photo_state_cert",
                "state_record": record,
            }
        elif str(record.get("status") or "") == "sold_inventory":
            candidates[path] = {
                "reason": "sold_inventory_state",
                "state_record": record,
            }

    sold_state_updates = 0
    for sha, record in photos.items():
        if not isinstance(record, dict):
            continue
        status = str(record.get("status") or "")
        if status in {"sold_inventory", "archived_from_album"}:
            continue
        record_certs = {cert(value) for value in record.get("certs") or [] if cert(value)}
        if record_certs and record_certs & sold_certs and not record_certs & active_certs and state_record_exists_in_album(record, root):
            sold_state_updates += 1
            if apply:
                record["status"] = "sold_inventory"
                record["sold_state_marked_at"] = datetime.now().isoformat(timespec="seconds")

    moved: list[dict[str, str]] = []
    if apply and candidates:
        archive_day = datetime.now().strftime("%Y-%m-%d")
        archive_dir = root / "DELETED ARCHIVE" / "INVENTORY PHOTOS" / archive_day
        archive_dir.mkdir(parents=True, exist_ok=True)
        expires_at = (datetime.now() + timedelta(days=RETENTION_DAYS)).isoformat(timespec="seconds")
        for path, info in sorted(candidates.items(), key=lambda item: str(item[0]).lower()):
            try:
                sha = file_hash(path)
            except Exception:
                sha = ""
            archive_path = unique_archive_path(archive_dir, path.name)
            shutil.move(str(path), str(archive_path))
            metadata = {
                "reason": "sold_inventory_photo_archive",
                "original_path": str(path),
                "archive_path": str(archive_path),
                "archived_at": datetime.now().isoformat(timespec="seconds"),
                "archive_expires_at": expires_at,
                "details": {
                    "reason": info.get("reason", ""),
                    "card_title": (info.get("profit_row") or {}).get("card_title", ""),
                    "cert_number": (info.get("profit_row") or {}).get("cert_number", ""),
                    "inventory_key": (info.get("profit_row") or {}).get("inventory_key", ""),
                    "source_sheet": (info.get("profit_row") or {}).get("source_sheet", ""),
                },
            }
            save_json(archive_path.with_name(archive_path.name + ".archive.json"), metadata)
            record = photos.get(sha) if sha and isinstance(photos.get(sha), dict) else {}
            if not isinstance(record, dict):
                record = {}
            record.update(
                {
                    "path": str(path),
                    "relative_path": path.relative_to(photo_root).as_posix(),
                    "filename": path.name,
                    "sha256": sha,
                    "status": "archived_from_album",
                    "archived_at": metadata["archived_at"],
                    "archive_expires_at": expires_at,
                    "archive_path": str(archive_path),
                    "last_seen": metadata["archived_at"],
                }
            )
            if sha:
                photos[sha] = record
            moved.append({"from": str(path), "to": str(archive_path), "reason": str(info.get("reason", ""))})

        state["updated_at"] = datetime.now().isoformat(timespec="seconds")
        save_json(state_path, state)
    elif apply and sold_state_updates:
        state["updated_at"] = datetime.now().isoformat(timespec="seconds")
        save_json(state_path, state)

    return {
        "root": str(root),
        "active_rows": len(active_rows),
        "profit_rows": len(profit),
        "photo_files": len(files),
        "archive_candidates": len(candidates),
        "sold_state_updates": sold_state_updates,
        "applied": apply,
        "moved": moved[:20],
        "sample_candidates": [
            {"path": str(path), "reason": str(info.get("reason", ""))}
            for path, info in sorted(candidates.items(), key=lambda item: str(item[0]).lower())[:20]
        ],
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--profile", choices=["personal", "team", "both"], default="both")
    parser.add_argument("--apply", action="store_true")
    args = parser.parse_args()

    roots = []
    if args.profile in {"personal", "both"}:
        roots.append(("personal", PERSONAL_ROOT))
    if args.profile in {"team", "both"}:
        roots.append(("team", TEAM_ROOT))

    for label, root in roots:
        result = archive_profile(root, args.apply)
        print(f"## {label}")
        print(json.dumps(result, indent=2))
    if not args.apply:
        print("Dry run only. Re-run with --apply to move files and update inventory_photo_state.json.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
