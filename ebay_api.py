from __future__ import annotations

import base64
import json
import os
import secrets
import time
import urllib.error
import urllib.parse
import urllib.request
from dataclasses import dataclass
from pathlib import Path


PRODUCTION_TOKEN_URL = "https://api.ebay.com/identity/v1/oauth2/token"
SANDBOX_TOKEN_URL = "https://api.sandbox.ebay.com/identity/v1/oauth2/token"
PRODUCTION_AUTHORIZE_URL = "https://auth.ebay.com/oauth2/authorize"
SANDBOX_AUTHORIZE_URL = "https://auth.sandbox.ebay.com/oauth2/authorize"
CONNECT_STATE_KIND = "lucas_ebay_connect"
DEFAULT_SCOPES = (
    "https://api.ebay.com/oauth/api_scope",
    "https://api.ebay.com/oauth/api_scope/sell.account",
    "https://api.ebay.com/oauth/api_scope/sell.inventory",
)


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

    @property
    def authorize_url(self) -> str:
        return SANDBOX_AUTHORIZE_URL if self.env.lower() == "sandbox" else PRODUCTION_AUTHORIZE_URL

    @classmethod
    def from_env(cls) -> "EbayConfig":
        scopes = tuple(str(os.environ.get("EBAY_OAUTH_SCOPES") or "").split()) or DEFAULT_SCOPES
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


def encode_connect_state(account: str = "", profile: str = "") -> str:
    payload = {
        "kind": CONNECT_STATE_KIND,
        "account": str(account or "default").strip() or "default",
        "profile": str(profile or "").strip().lower(),
        "nonce": secrets.token_urlsafe(16),
        "created_at": int(time.time()),
    }
    raw = json.dumps(payload, separators=(",", ":"), sort_keys=True).encode("utf-8")
    return base64.urlsafe_b64encode(raw).decode("ascii").rstrip("=")


def decode_connect_state(value: str) -> dict[str, object]:
    token = str(value or "").strip()
    if not token:
        return {}
    padding = "=" * (-len(token) % 4)
    try:
        payload = json.loads(base64.urlsafe_b64decode((token + padding).encode("ascii")).decode("utf-8"))
    except Exception:
        return {}
    if not isinstance(payload, dict) or payload.get("kind") != CONNECT_STATE_KIND:
        return {}
    return payload


def build_authorization_url(config: EbayConfig, state: str, scopes: tuple[str, ...] | None = None) -> str:
    config.validate()
    scope_values = scopes if scopes is not None else config.scopes
    if not scope_values:
        raise EbayOAuthError("Missing eBay OAuth scopes.")
    return config.authorize_url + "?" + urllib.parse.urlencode(
        {
            "client_id": config.client_id,
            "redirect_uri": config.runame,
            "response_type": "code",
            "scope": " ".join(scope_values),
            "state": state,
        }
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


def ebay_token_store_path(data_root: object = None) -> Path:
    configured = str(os.environ.get("EBAY_TOKEN_STORE_PATH") or "").strip()
    if configured:
        return Path(configured).expanduser()
    root = Path(str(data_root or "")).expanduser() if str(data_root or "").strip() else Path(__file__).resolve().parent / "work"
    return root / "ebay_accounts.json"


def load_ebay_accounts(path: Path) -> dict[str, object]:
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {"accounts": {}}
    if not isinstance(data, dict):
        return {"accounts": {}}
    accounts = data.get("accounts")
    if not isinstance(accounts, dict):
        data["accounts"] = {}
    return data


def save_ebay_account_token(
    path: Path,
    account: str,
    config: EbayConfig,
    token_result: dict[str, object],
    scopes: tuple[str, ...] | None = None,
) -> dict[str, object]:
    refresh_token = str(token_result.get("refresh_token") or "").strip()
    if not refresh_token:
        raise EbayOAuthError("eBay accepted the authorization code but did not return a refresh_token.")
    access_token = str(token_result.get("access_token") or "").strip()
    account_key = str(account or "default").strip() or "default"
    now = int(time.time())
    data = load_ebay_accounts(path)
    accounts = data.setdefault("accounts", {})
    if not isinstance(accounts, dict):
        accounts = {}
        data["accounts"] = accounts
    previous = accounts.get(account_key) if isinstance(accounts.get(account_key), dict) else {}
    record = {
        **previous,
        "account": account_key,
        "env": config.env,
        "client_id": config.client_id,
        "runame": config.runame,
        "scopes": list(scopes if scopes is not None else config.scopes),
        "refresh_token": refresh_token,
        "refresh_token_expires_in": token_result.get("refresh_token_expires_in"),
        "access_token": access_token,
        "access_token_expires_in": token_result.get("expires_in"),
        "connected_at": previous.get("connected_at") or now,
        "updated_at": now,
    }
    accounts[account_key] = record
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(path.suffix + ".tmp")
    tmp.write_text(json.dumps(data, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    tmp.replace(path)
    return record


def ebay_account_status(path: Path) -> dict[str, object]:
    data = load_ebay_accounts(path)
    accounts = data.get("accounts") if isinstance(data, dict) else {}
    result: list[dict[str, object]] = []
    if isinstance(accounts, dict):
        for key, record in sorted(accounts.items()):
            if not isinstance(record, dict):
                continue
            result.append(
                {
                    "account": key,
                    "env": record.get("env", ""),
                    "client_id": record.get("client_id", ""),
                    "connected_at": record.get("connected_at"),
                    "updated_at": record.get("updated_at"),
                    "refresh_token": mask_token(record.get("refresh_token", "")),
                    "scopes": record.get("scopes", []),
                }
            )
    return {"ok": True, "accounts": result}


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
