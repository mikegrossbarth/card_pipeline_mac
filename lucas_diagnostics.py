from __future__ import annotations

import json
import os
import subprocess
import sys
import time
from pathlib import Path
from typing import Any

from google_sheets_import import TOKEN_PATH, load_local_env_files, load_token, oauth_client_config, token_matches_client


ROOT = Path(__file__).resolve().parent


def lucas_version_label(platform_label: str = "") -> str:
    commit = git_short_commit()
    suffix = f" {platform_label.strip()}" if platform_label.strip() else ""
    return f"LUCAS{suffix} {commit or 'unknown'}"


def git_short_commit() -> str:
    try:
        result = subprocess.run(
            ["git", "-C", str(ROOT), "rev-parse", "--short", "HEAD"],
            capture_output=True,
            text=True,
            timeout=3,
            check=False,
        )
    except Exception:
        return ""
    return result.stdout.strip() if result.returncode == 0 else ""


def google_token_status() -> dict[str, Any]:
    load_local_env_files()
    token = load_token()
    token_exists = TOKEN_PATH.exists()
    client_id = ""
    client_secret = ""
    config_error = ""
    try:
        client_id, client_secret = oauth_client_config()
    except Exception as error:
        config_error = str(error)
    expires_at = float(token.get("expires_at") or 0) if isinstance(token, dict) else 0
    expires_in = int(expires_at - time.time()) if expires_at else None
    account = str(token.get("email") or token.get("account") or "").strip() if isinstance(token, dict) else ""
    if not account:
        account = "Unavailable from Sheets token"
    return {
        "token_exists": token_exists,
        "token_path": str(TOKEN_PATH),
        "token_has_access_token": bool(token.get("access_token")) if isinstance(token, dict) else False,
        "token_has_refresh_token": bool(token.get("refresh_token")) if isinstance(token, dict) else False,
        "token_matches_client": token_matches_client(token, client_id) if token and client_id else False,
        "token_expires_in_seconds": expires_in,
        "account": account,
        "client_id_present": bool(client_id),
        "client_secret_present": bool(client_secret),
        "config_error": config_error,
    }


def google_status_lines(sheet_status: str = "", keep_status: str = "") -> list[str]:
    status = google_token_status()
    lines = [
        f"OAuth token exists: {'yes' if status['token_exists'] else 'no'}",
        f"Token account: {status['account']}",
        f"Token path: {status['token_path']}",
        f"Access token saved: {'yes' if status['token_has_access_token'] else 'no'}",
        f"Refresh token saved: {'yes' if status['token_has_refresh_token'] else 'no'}",
        f"OAuth client ID present: {'yes' if status['client_id_present'] else 'no'}",
        f"OAuth client secret present: {'yes' if status['client_secret_present'] else 'no'}",
    ]
    if status["token_expires_in_seconds"] is not None:
        lines.append(f"Token expires in: {max(0, status['token_expires_in_seconds'])} seconds")
    if sheet_status:
        lines.append(f"Source Sheet can be opened: {sheet_status}")
    if keep_status:
        lines.append(f"Keep note last synced: {keep_status}")
    if status["config_error"]:
        lines.append(f"Google config error: {status['config_error']}")
    return lines


def setup_doctor_results(pipeline_root: Path, bridge_snapshot: dict[str, Any] | None = None, platform_label: str = "") -> list[dict[str, str]]:
    load_local_env_files()
    bridge_snapshot = bridge_snapshot or {}
    token = google_token_status()
    required_dirs = [
        pipeline_root / "WORKING SHEETS",
        pipeline_root / "INCOMING SHEETS",
        pipeline_root / "RECEIVED SHEETS",
        pipeline_root / "ASSIGNMENT RULES",
        pipeline_root / "COMPANY SHEETS",
    ]
    extension_version = str(bridge_snapshot.get("extensionVersion") or "")
    expected_extension = str(bridge_snapshot.get("expectedExtensionVersion") or "")
    rows = [
        check_row("App version", bool(git_short_commit()), lucas_version_label(platform_label)),
        check_row("Python/dependencies", True, f"Python {sys.version.split()[0]}"),
        check_row("GOOGLE_API_KEY in .env", bool(os.environ.get("GOOGLE_API_KEY", "").strip()), "Used by Photo OCR and Card Ladder OCR fallback."),
        check_row("Sheets OAuth client ID", token["client_id_present"], token["config_error"] or "GOOGLE_SHEETS_OAUTH_CLIENT_ID"),
        check_row("Sheets OAuth client secret", token["client_secret_present"], "GOOGLE_SHEETS_OAUTH_CLIENT_SECRET"),
        check_row("Sheets OAuth token", bool(token["token_exists"] and token["token_has_access_token"]), token["token_path"]),
        check_row("Shared pipeline folders", all(path.exists() for path in required_dirs), ", ".join(path.name for path in required_dirs if not path.exists()) or str(pipeline_root)),
        check_row("Chrome extension reachable", bool(extension_version), f"Seen version: {extension_version or 'not seen'}"),
        check_row("Card Ladder helper version", bool(extension_version and expected_extension and extension_version == expected_extension), f"Expected {expected_extension or 'unknown'}, seen {extension_version or 'not seen'}"),
        check_row("Company Rules config", (ROOT / "assignment_companies.json").exists(), str(ROOT / "assignment_companies.json")),
    ]
    return rows


def check_row(name: str, ok: bool, detail: str = "") -> dict[str, str]:
    return {"name": name, "status": "OK" if ok else "Needs attention", "detail": str(detail or "")}


def diagnostic_json(payload: dict[str, Any]) -> str:
    return json.dumps(payload, indent=2, default=str)
