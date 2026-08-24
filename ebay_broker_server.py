from __future__ import annotations

import json
import base64
import hashlib
import hmac
import os
import secrets
import time
import urllib.parse
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

from ebay_api import (
    EbayConfig,
    EbayOAuthError,
    build_authorization_url,
    exchange_authorization_code,
    refresh_access_token,
)


BROKER_STATE_KIND = "lucas_ebay_broker"
DEFAULT_STORE_PATH = Path(__file__).resolve().parent / "work" / "ebay_broker_connections.json"


def broker_store_path() -> Path:
    configured = str(os.environ.get("LUCAS_EBAY_BROKER_STORE_PATH") or "").strip()
    return Path(configured).expanduser() if configured else DEFAULT_STORE_PATH


def broker_public_url() -> str:
    return str(os.environ.get("LUCAS_EBAY_BROKER_PUBLIC_URL") or "https://lucas.mikeyscards.com/ebay").strip().rstrip("/")


def broker_state_secret() -> str:
    return str(os.environ.get("LUCAS_EBAY_BROKER_STATE_SECRET") or os.environ.get("EBAY_CLIENT_SECRET") or "").strip()


def _load_store() -> dict[str, object]:
    path = broker_store_path()
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {"version": 1, "connections": {}}
    if not isinstance(data, dict):
        return {"version": 1, "connections": {}}
    if not isinstance(data.get("connections"), dict):
        data["connections"] = {}
    data.setdefault("version", 1)
    return data


def _save_store(data: dict[str, object]) -> None:
    path = broker_store_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(path.suffix + ".tmp")
    tmp.write_text(json.dumps(data, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    tmp.replace(path)


def _encode_state(payload: dict[str, object]) -> str:
    raw = json.dumps(payload, separators=(",", ":"), sort_keys=True).encode("utf-8")
    secret = broker_state_secret()
    if not secret:
        raise EbayOAuthError("Missing LUCAS_EBAY_BROKER_STATE_SECRET.")
    signature = hmac.new(secret.encode("utf-8"), raw, hashlib.sha256).digest()
    token = base64.urlsafe_b64encode(raw).decode("ascii").rstrip("=")
    sig = base64.urlsafe_b64encode(signature).decode("ascii").rstrip("=")
    return f"{token}.{sig}"


def _decode_state(value: str) -> dict[str, object]:
    try:
        token, sig = str(value or "").split(".", 1)
        raw = base64.urlsafe_b64decode((token + "=" * (-len(token) % 4)).encode("ascii"))
        supplied = base64.urlsafe_b64decode((sig + "=" * (-len(sig) % 4)).encode("ascii"))
    except Exception:
        return {}
    secret = broker_state_secret()
    if not secret:
        return {}
    expected = hmac.new(secret.encode("utf-8"), raw, hashlib.sha256).digest()
    if not hmac.compare_digest(supplied, expected):
        return {}
    try:
        payload = json.loads(raw.decode("utf-8"))
    except json.JSONDecodeError:
        return {}
    return payload if isinstance(payload, dict) and payload.get("kind") == BROKER_STATE_KIND else {}


def _callback_allowed(callback: str) -> bool:
    parsed = urllib.parse.urlparse(str(callback or "").strip())
    return parsed.scheme == "http" and parsed.hostname in {"127.0.0.1", "localhost"} and parsed.path == "/ebay/broker/callback"


def _redirect(handler: BaseHTTPRequestHandler, target: str) -> None:
    handler.send_response(302)
    handler.send_header("location", target)
    handler.send_header("cache-control", "no-store")
    handler.send_header("content-length", "0")
    handler.end_headers()


def _send_json(handler: BaseHTTPRequestHandler, payload: dict[str, object], status: int = 200) -> None:
    body = json.dumps(payload).encode("utf-8")
    handler.send_response(status)
    handler.send_header("content-type", "application/json")
    handler.send_header("cache-control", "no-store")
    handler.send_header("x-content-type-options", "nosniff")
    handler.send_header("content-length", str(len(body)))
    handler.end_headers()
    handler.wfile.write(body)


def _send_page(handler: BaseHTTPRequestHandler, title: str, body: str, status: int = 200) -> None:
    html = f"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{title}</title>
<style>
body {{
  margin: 0;
  min-height: 100vh;
  display: grid;
  place-items: center;
  background: #0b141b;
  color: #f3f4f6;
  font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
}}
main {{ max-width: 680px; padding: 32px; }}
h1 {{ font-size: 40px; margin: 0 0 16px; }}
p {{ color: #cbd5e1; font-size: 18px; line-height: 1.5; }}
</style>
</head>
<body><main>{body}</main></body>
</html>
""".encode("utf-8")
    handler.send_response(status)
    handler.send_header("content-type", "text/html; charset=utf-8")
    handler.send_header("cache-control", "no-store")
    handler.send_header("x-content-type-options", "nosniff")
    handler.send_header("content-length", str(len(html)))
    handler.end_headers()
    handler.wfile.write(html)


class EbayBrokerHandler(BaseHTTPRequestHandler):
    server_version = "LUCAS-eBay-Broker/1"

    def log_message(self, _format: str, *_args) -> None:
        return

    def _route_path(self, path: str) -> str:
        if path == "/ebay":
            return "/"
        if path.startswith("/ebay/"):
            return path[5:] or "/"
        return path

    def do_GET(self) -> None:
        parsed = urllib.parse.urlparse(self.path)
        route_path = self._route_path(parsed.path)
        if route_path in {"/", ""}:
            _send_page(
                self,
                "LUCAS eBay",
                "<h1>LUCAS eBay</h1><p>This service securely connects eBay seller accounts to LUCAS.</p>",
            )
            return
        if route_path in {"/health", "/status"}:
            _send_json(
                self,
                {
                    "ok": True,
                    "service": "lucas-ebay-broker",
                    "public_url": broker_public_url(),
                    "mode": EbayConfig.from_env().env,
                },
            )
            return
        if route_path == "/connect":
            self._send_connect(parsed)
            return
        if route_path == "/callback":
            self._send_callback(parsed)
            return
        _send_json(self, {"ok": False, "error": "not found"}, status=404)

    def do_POST(self) -> None:
        parsed = urllib.parse.urlparse(self.path)
        route_path = self._route_path(parsed.path)
        if route_path == "/token":
            self._send_token()
            return
        _send_json(self, {"ok": False, "error": "not found"}, status=404)

    def _send_connect(self, parsed) -> None:
        query = urllib.parse.parse_qs(parsed.query)
        callback = str(query.get("callback", [""])[0] or "").strip()
        if not _callback_allowed(callback):
            _send_json(self, {"ok": False, "error": "invalid desktop callback"}, status=400)
            return
        account = str(query.get("account", ["default"])[0] or "default").strip() or "default"
        profile = str(query.get("profile", [""])[0] or "").strip().lower()
        state = _encode_state(
            {
                "kind": BROKER_STATE_KIND,
                "nonce": secrets.token_urlsafe(16),
                "created_at": int(time.time()),
                "account": account,
                "profile": profile,
                "callback": callback,
            }
        )
        try:
            config = EbayConfig.from_env()
            target = build_authorization_url(config, state)
        except EbayOAuthError as error:
            _send_json(self, {"ok": False, "error": str(error)}, status=500)
            return
        _redirect(self, target)

    def _send_callback(self, parsed) -> None:
        query = urllib.parse.parse_qs(parsed.query)
        error = str(query.get("error", [""])[0] or "").strip()
        state = _decode_state(str(query.get("state", [""])[0] or ""))
        callback = str(state.get("callback") or "").strip()
        if not _callback_allowed(callback):
            _send_json(self, {"ok": False, "error": "invalid saved desktop callback"}, status=400)
            return
        if error:
            target = callback + "?" + urllib.parse.urlencode({"error": error, "error_description": query.get("error_description", [""])[0]})
            _redirect(self, target)
            return
        code = str(query.get("code", [""])[0] or "").strip()
        if not code:
            target = callback + "?" + urllib.parse.urlencode({"error": "missing_code"})
            _redirect(self, target)
            return
        try:
            config = EbayConfig.from_env()
            token_result = exchange_authorization_code(config, code)
        except EbayOAuthError as error:
            target = callback + "?" + urllib.parse.urlencode({"error": "exchange_failed", "error_description": str(error)})
            _redirect(self, target)
            return
        refresh_token = str(token_result.get("refresh_token") or "").strip()
        if not refresh_token:
            target = callback + "?" + urllib.parse.urlencode({"error": "missing_refresh_token"})
            _redirect(self, target)
            return
        connection_token = secrets.token_urlsafe(32)
        data = _load_store()
        connections = data.setdefault("connections", {})
        if not isinstance(connections, dict):
            connections = {}
            data["connections"] = connections
        account = str(state.get("account") or "default").strip() or "default"
        connections[connection_token] = {
            "account": account,
            "profile": str(state.get("profile") or "").strip(),
            "refresh_token": refresh_token,
            "scopes": list(config.scopes),
            "env": config.env,
            "created_at": int(time.time()),
            "updated_at": int(time.time()),
        }
        _save_store(data)
        target = callback + "?" + urllib.parse.urlencode(
            {
                "connection_token": connection_token,
                "account": account,
                "broker_url": broker_public_url(),
                "marketplace_id": "EBAY_US",
            }
        )
        _redirect(self, target)

    def _send_token(self) -> None:
        length = int(self.headers.get("content-length") or "0")
        try:
            payload = json.loads(self.rfile.read(length).decode("utf-8")) if length else {}
        except json.JSONDecodeError:
            _send_json(self, {"ok": False, "error": "invalid json"}, status=400)
            return
        connection_token = str(payload.get("connection_token") or "").strip()
        data = _load_store()
        connections = data.get("connections") if isinstance(data.get("connections"), dict) else {}
        record = connections.get(connection_token) if isinstance(connections, dict) else {}
        if not isinstance(record, dict):
            _send_json(self, {"ok": False, "error": "unknown connection"}, status=404)
            return
        try:
            token_result = refresh_access_token(EbayConfig.from_env(), str(record.get("refresh_token") or ""))
        except EbayOAuthError as error:
            _send_json(self, {"ok": False, "error": str(error)}, status=502)
            return
        record["updated_at"] = int(time.time())
        try:
            _save_store(data)
        except OSError:
            pass
        _send_json(
            self,
            {
                "ok": True,
                "access_token": token_result.get("access_token", ""),
                "expires_in": token_result.get("expires_in"),
            },
        )


def run(host: str = "127.0.0.1", port: int = 8788) -> None:
    server = ThreadingHTTPServer((host, port), EbayBrokerHandler)
    print(f"LUCAS eBay broker listening on http://{host}:{port}")
    server.serve_forever()


if __name__ == "__main__":
    run(os.environ.get("HOST", "127.0.0.1"), int(os.environ.get("PORT", "8788")))
