from __future__ import annotations

import re

from intake_io import scan_to_cert


def instagram_inventory_identity(record: dict[str, object] | None) -> str:
    if not isinstance(record, dict):
        return ""
    cert = scan_to_cert(record.get("cert_number"))
    if cert and len(cert) >= 5:
        return f"cert:{cert}"
    item_id = str(record.get("item_id") or "").strip().lower()
    if item_id:
        return f"item:{re.sub(r'[^a-z0-9]+', '', item_id)}"
    title = str(record.get("card_title") or record.get("caption") or "").strip().lower()
    title = re.sub(r"[^a-z0-9]+", " ", title)
    title = re.sub(r"\s+", " ", title).strip()
    return f"title:{title}" if title else ""


def instagram_post_entry_identity(entry: dict[str, object]) -> str:
    explicit = str(entry.get("inventory_identity") or "").strip()
    if explicit:
        return explicit
    return instagram_inventory_identity(
        {
            "cert_number": entry.get("cert_number"),
            "item_id": entry.get("item_id"),
            "card_title": entry.get("card_title") or entry.get("caption"),
        }
    )


def instagram_match_text_tokens(value: object) -> set[str]:
    text = str(value or "").lower()
    tokens = set(re.findall(r"[a-z0-9]+", text))
    stop_words = {
        "the", "and", "with", "for", "card", "cards", "auto", "rc", "rookie", "psa", "bgs",
        "cgc", "sgc", "gem", "mint", "auto", "autograph", "number", "serial", "refractor",
    }
    return {token for token in tokens if len(token) >= 3 and token not in stop_words}
