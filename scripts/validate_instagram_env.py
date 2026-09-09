#!/usr/bin/env python3
"""Validate the LUCAS Instagram env without printing secrets."""

from __future__ import annotations

import json
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from datetime import datetime, timezone
from pathlib import Path


EXPECTED_APP_ID = "1043156651570082"
EXPECTED_PAGE_ID = "1090039820868692"
EXPECTED_INSTAGRAM_ID = "17841465322546974"
BAD_INSTAGRAM_ID = "17841415167583312"
GRAPH_BASE = "https://graph.facebook.com/v26.0"
MIN_TOKEN_SECONDS = 7 * 24 * 60 * 60


def load_env(path: Path) -> dict[str, str]:
    values: dict[str, str] = {}
    for raw_line in path.read_text().splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        values[key.strip()] = value.strip().strip('"').strip("'")
    return values


def graph_get(path: str, token: str) -> tuple[int, dict[str, object]]:
    separator = "&" if "?" in path else "?"
    url = f"{GRAPH_BASE}{path}{separator}{urllib.parse.urlencode({'access_token': token})}"
    try:
        with urllib.request.urlopen(url, timeout=20) as response:
            return response.status, json.loads(response.read().decode())
    except urllib.error.HTTPError as exc:
        body = exc.read().decode(errors="replace")
        try:
            return exc.code, json.loads(body)
        except json.JSONDecodeError:
            return exc.code, {"raw": body[:300]}


def describe_error(payload: dict[str, object]) -> str:
    error = payload.get("error")
    if isinstance(error, dict):
        return f"code={error.get('code')} message={error.get('message')}"
    return str(payload)


def debug_token(app_id: str, app_secret: str, token: str) -> tuple[int, dict[str, object]]:
    app_access_token = f"{app_id}|{app_secret}"
    query = urllib.parse.urlencode({"input_token": token, "access_token": app_access_token})
    url = f"{GRAPH_BASE}/debug_token?{query}"
    try:
        with urllib.request.urlopen(url, timeout=20) as response:
            return response.status, json.loads(response.read().decode())
    except urllib.error.HTTPError as exc:
        body = exc.read().decode(errors="replace")
        try:
            return exc.code, json.loads(body)
        except json.JSONDecodeError:
            return exc.code, {"raw": body[:300]}


def format_expires_at(value: object) -> str:
    try:
        timestamp = int(value or 0)
    except (TypeError, ValueError):
        timestamp = 0
    if not timestamp:
        return "never/unknown"
    return datetime.fromtimestamp(timestamp, timezone.utc).isoformat()


def main() -> int:
    env_path = Path(__file__).resolve().parents[1] / ".env"
    values = load_env(env_path)
    app_id = values.get("LUCAS_INSTAGRAM_APP_ID", "")
    instagram_id = values.get("LUCAS_INSTAGRAM_USER_ID", "")
    token = values.get("LUCAS_INSTAGRAM_ACCESS_TOKEN", "")
    app_secret = values.get("LUCAS_INSTAGRAM_APP_SECRET", "")
    failures: list[str] = []

    if app_id != EXPECTED_APP_ID:
        failures.append(f"LUCAS_INSTAGRAM_APP_ID is {app_id or '<missing>'}; expected {EXPECTED_APP_ID}.")
    if instagram_id != EXPECTED_INSTAGRAM_ID:
        failures.append(
            f"LUCAS_INSTAGRAM_USER_ID is {instagram_id or '<missing>'}; expected {EXPECTED_INSTAGRAM_ID}."
        )
        if instagram_id == BAD_INSTAGRAM_ID:
            failures.append(f"{BAD_INSTAGRAM_ID} is the known failing ID for media/quota calls.")
    if not token:
        failures.append("LUCAS_INSTAGRAM_ACCESS_TOKEN is missing.")

    print(f"App ID: {app_id or '<missing>'}")
    print(f"Instagram ID: {instagram_id or '<missing>'}")
    print(f"Token present: {'yes' if token else 'no'}")
    print(f"App secret present: {'yes' if app_secret else 'no'}")

    if token and app_id and app_secret:
        status, payload = debug_token(app_id, app_secret, token)
        data = payload.get("data") if isinstance(payload, dict) else {}
        data = data if isinstance(data, dict) else {}
        print(f"token type: {data.get('type') or '<unknown>'}")
        print(f"token valid: {data.get('is_valid')}")
        print(f"token profile/page id: {data.get('profile_id') or '<unknown>'}")
        print(f"token expires: {format_expires_at(data.get('expires_at'))}")
        if status != 200:
            failures.append(f"debug_token failed: {describe_error(payload)}")
        if data.get("type") != "PAGE":
            failures.append(f"token type is {data.get('type') or '<unknown>'}; expected PAGE.")
        if str(data.get("profile_id") or "") != EXPECTED_PAGE_ID:
            failures.append(f"token profile/page id is {data.get('profile_id') or '<missing>'}; expected {EXPECTED_PAGE_ID}.")
        if data.get("is_valid") is not True:
            failures.append("token is not valid according to Meta debug_token.")
        try:
            expires_at = int(data.get("expires_at") or 0)
        except (TypeError, ValueError):
            expires_at = 0
        if expires_at:
            remaining = expires_at - int(time.time())
            if remaining < MIN_TOKEN_SECONDS:
                failures.append("token expires in under 7 days; this is probably a short-lived token chain.")
    elif token:
        print("token debug: skipped because LUCAS_INSTAGRAM_APP_SECRET is missing.")

    if token:
        checks = [
            ("profile", f"/{EXPECTED_INSTAGRAM_ID}?fields=id,username,name"),
            ("media", f"/{EXPECTED_INSTAGRAM_ID}/media?fields=id,permalink&limit=1"),
            ("quota", f"/{EXPECTED_INSTAGRAM_ID}/content_publishing_limit"),
            (
                "page link",
                f"/{EXPECTED_PAGE_ID}?fields=name,id,connected_page_backed_instagram_account",
            ),
        ]
        for label, path in checks:
            status, payload = graph_get(path, token)
            if status == 200:
                print(f"{label}: OK")
            else:
                message = describe_error(payload)
                print(f"{label}: ERROR {status} {message}")
                failures.append(f"{label} check failed: {message}")

    if failures:
        print("\nFAILED")
        for failure in failures:
            print(f"- {failure}")
        return 1

    print("\nOK: Instagram env matches the known-good LUCAS setup.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
