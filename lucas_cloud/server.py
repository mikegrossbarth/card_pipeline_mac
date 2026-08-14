from __future__ import annotations

import json
import os
import re
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse

from .service import LucasCloudService


def header_token(handler: BaseHTTPRequestHandler, query: dict[str, list[str]]) -> str:
    auth = handler.headers.get("Authorization", "")
    if auth.lower().startswith("bearer "):
        return auth[7:].strip()
    return handler.headers.get("X-LUCAS-Token", "").strip() or (query.get("token") or [""])[0]


class LucasCloudHandler(BaseHTTPRequestHandler):
    service: LucasCloudService

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        query = parse_qs(parsed.query)
        if parsed.path == "/health":
            self.send_json({"ok": True, "service": "lucas-cloud"})
            return
        match = re.fullmatch(r"/api/v1/workspaces/([^/]+)/photos/pending", parsed.path)
        if match:
            self.guard(lambda: self.service.pending_photos(match.group(1), header_token(self, query), int((query.get("limit") or ["100"])[0])))
            return
        self.send_json({"ok": False, "error": "Not found."}, status=404)

    def do_POST(self) -> None:
        parsed = urlparse(self.path)
        query = parse_qs(parsed.query)
        body = self.read_json()
        if parsed.path == "/api/v1/workspaces":
            admin_token = os.environ.get("LUCAS_CLOUD_ADMIN_TOKEN", "")
            if admin_token and header_token(self, query) != admin_token:
                self.send_json({"ok": False, "error": "Admin token required."}, status=403)
                return
            self.guard(lambda: self.service.create_workspace(str(body.get("slug") or ""), str(body.get("display_name") or body.get("name") or "")))
            return
        match = re.fullmatch(r"/api/v1/workspaces/([^/]+)/photos/upload", parsed.path)
        if match:
            self.guard(lambda: self.service.upload_mobile_photos(match.group(1), header_token(self, query), body))
            return
        match = re.fullmatch(r"/api/v1/workspaces/([^/]+)/photos/([^/]+)/status", parsed.path)
        if match:
            self.guard(lambda: self.service.mark_photo_status(match.group(1), header_token(self, query), match.group(2), body))
            return
        match = re.fullmatch(r"/api/v1/workspaces/([^/]+)/inventory/upsert", parsed.path)
        if match:
            self.guard(lambda: self.service.upsert_inventory(match.group(1), header_token(self, query), body))
            return
        self.send_json({"ok": False, "error": "Not found."}, status=404)

    def read_json(self) -> dict:
        length = int(self.headers.get("Content-Length") or 0)
        if length <= 0:
            return {}
        raw = self.rfile.read(length)
        try:
            value = json.loads(raw.decode("utf-8"))
        except json.JSONDecodeError:
            return {}
        return value if isinstance(value, dict) else {}

    def guard(self, callback) -> None:
        try:
            result = callback()
            self.send_json(result)
        except PermissionError as error:
            self.send_json({"ok": False, "error": str(error)}, status=403)
        except KeyError as error:
            self.send_json({"ok": False, "error": str(error).strip("'")}, status=404)
        except ValueError as error:
            self.send_json({"ok": False, "error": str(error)}, status=400)
        except Exception as error:
            self.send_json({"ok": False, "error": str(error)}, status=500)

    def send_json(self, payload: dict, status: int = 200) -> None:
        body = json.dumps(payload, sort_keys=True).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def log_message(self, format: str, *args) -> None:
        return


def make_server(data_root: Path | str, host: str = "127.0.0.1", port: int = 8780) -> ThreadingHTTPServer:
    class Handler(LucasCloudHandler):
        service = LucasCloudService(data_root)

    return ThreadingHTTPServer((host, port), Handler)


def main() -> None:
    data_root = Path(os.environ.get("LUCAS_CLOUD_DATA_ROOT") or "work/lucas_cloud_data").expanduser()
    host = os.environ.get("LUCAS_CLOUD_HOST", "127.0.0.1")
    port = int(os.environ.get("LUCAS_CLOUD_PORT", "8780"))
    server = make_server(data_root, host=host, port=port)
    print(f"LUCAS cloud server listening on http://{host}:{port}")
    server.serve_forever()


if __name__ == "__main__":
    main()
