from __future__ import annotations

import base64
import json
import os
import urllib.error
import urllib.parse
import urllib.request
from dataclasses import dataclass
from pathlib import Path


PRODUCTION_TOKEN_URL = "https://api.ebay.com/identity/v1/oauth2/token"
SANDBOX_TOKEN_URL = "https://api.sandbox.ebay.com/identity/v1/oauth2/token"


class EbayOAuthError(RuntimeError):
    pass


@dataclass(frozen=True)
class EbayConfig:
    env: str
    client_id: str
    client_secret: str
    runame: str
    scopes: tuple[str, ...] = ()

    @property
    def token_url(self) -> str:
        return SANDBOX_TOKEN_URL if self.env.lower() == "sandbox" else PRODUCTION_TOKEN_URL

    @classmethod
    def from_env(cls) -> "EbayConfig":
        scopes = tuple(str(os.environ.get("EBAY_OAUTH_SCOPES") or "").split())
        return cls(
            env=str(os.environ.get("EBAY_ENV") or "production").strip().lower() or "production",
            client_id=str(os.environ.get("EBAY_CLIENT_ID") or "").strip(),
            client_secret=str(os.environ.get("EBAY_CLIENT_SECRET") or "").strip(),
            runame=str(os.environ.get("EBAY_RUNAME") or "").strip(),
            scopes=scopes,
        )

    def validate(self) -> None:
        missing = []
        if not self.client_id:
            missing.append("EBAY_CLIENT_ID")
        if not self.client_secret:
            missing.append("EBAY_CLIENT_SECRET")
        if not self.runame:
            missing.append("EBAY_RUNAME")
        if missing:
            raise EbayOAuthError(f"Missing eBay config value(s): {', '.join(missing)}")


def mask_token(value: object) -> str:
    token = str(value or "").strip()
    if len(token) <= 16:
        return "*" * len(token)
    return f"{token[:8]}...{token[-6:]}"


def _basic_auth_header(config: EbayConfig) -> str:
    raw = f"{config.client_id}:{config.client_secret}".encode("utf-8")
    return "Basic " + base64.b64encode(raw).decode("ascii")


def _oauth_post(config: EbayConfig, payload: dict[str, object], timeout: int = 45) -> dict[str, object]:
    config.validate()
    data = urllib.parse.urlencode(payload).encode("utf-8")
    request = urllib.request.Request(
        config.token_url,
        data=data,
        headers={
            "Authorization": _basic_auth_header(config),
            "Content-Type": "application/x-www-form-urlencoded",
        },
        method="POST",
    )
    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            raw = response.read().decode("utf-8")
    except urllib.error.HTTPError as error:
        body = error.read().decode("utf-8", errors="replace")
        raise EbayOAuthError(body or str(error)) from error
    except urllib.error.URLError as error:
        raise EbayOAuthError(str(error)) from error
    try:
        return json.loads(raw) if raw else {}
    except json.JSONDecodeError as error:
        raise EbayOAuthError(f"eBay returned invalid JSON: {raw[:200]}") from error


def exchange_authorization_code(config: EbayConfig, authorization_code: str) -> dict[str, object]:
    code = str(authorization_code or "").strip()
    if not code:
        raise EbayOAuthError("Missing eBay authorization code.")
    return _oauth_post(
        config,
        {
            "grant_type": "authorization_code",
            "code": code,
            "redirect_uri": config.runame,
        },
    )


def refresh_access_token(config: EbayConfig, refresh_token: str, scopes: tuple[str, ...] | None = None) -> dict[str, object]:
    token = str(refresh_token or "").strip()
    if not token:
        raise EbayOAuthError("Missing EBAY_REFRESH_TOKEN.")
    payload: dict[str, object] = {
        "grant_type": "refresh_token",
        "refresh_token": token,
    }
    scope_values = scopes if scopes is not None else config.scopes
    if scope_values:
        payload["scope"] = " ".join(scope_values)
    return _oauth_post(config, payload)


def update_env_values(env_path: Path, values: dict[str, object]) -> None:
    existing = env_path.read_text(encoding="utf-8").splitlines() if env_path.exists() else []
    remaining = {key: str(value or "") for key, value in values.items() if str(value or "").strip()}
    updated: list[str] = []
    for line in existing:
        stripped = line.strip()
        key = stripped.split("=", 1)[0] if "=" in stripped and not stripped.startswith("#") else ""
        if key in remaining:
            updated.append(f"{key}={remaining.pop(key)}")
        else:
            updated.append(line)
    for key, value in remaining.items():
        updated.append(f"{key}={value}")
    env_path.write_text("\n".join(updated).rstrip() + "\n", encoding="utf-8")
