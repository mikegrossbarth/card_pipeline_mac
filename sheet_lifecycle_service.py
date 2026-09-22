from __future__ import annotations

from pathlib import Path

from intake_io import scan_to_cert


def accounted_source_key(value: object) -> str:
    return Path(str(value or "")).name.strip().lower()


def accounted_identity_key(cert: object = "", item_id: object = "") -> str:
    cert_key = scan_to_cert(cert)
    if cert_key:
        return f"cert:{cert_key}"
    item_key = str(item_id or "").strip().lower()
    if item_key:
        return f"item:{item_key}"
    return ""
