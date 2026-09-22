from __future__ import annotations

from pathlib import Path

from intake_io import scan_to_cert


def inventory_ledger_identity(record: dict[str, object]) -> str:
    inventory_key = str(record.get("inventory_key") or "").strip()
    if inventory_key:
        return inventory_key
    stable_id = scan_to_cert(record.get("cert_number")) or str(record.get("item_id") or "").strip().lower()
    source = Path(str(record.get("source_sheet") or "")).name.strip().lower()
    person = str(record.get("assigned_person") or "").strip().lower()
    return "|".join(part for part in (stable_id, source, person) if part)
