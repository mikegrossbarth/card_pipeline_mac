from __future__ import annotations

from intake_io import scan_to_cert


def photo_state_matches_sold_cert(existing: dict[str, object], sold_certs: set[str]) -> bool:
    if not existing or not sold_certs:
        return False
    photo_certs = {scan_to_cert(cert) for cert in (existing.get("certs") or []) if scan_to_cert(cert)}
    return bool(photo_certs & sold_certs)
