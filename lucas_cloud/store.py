from __future__ import annotations

import base64
import hashlib
import json
import re
import secrets
import sqlite3
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from .models import InventoryItem, PhotoRecord, Workspace


SCHEMA_VERSION = 1


def utc_now() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")


def safe_slug(value: str) -> str:
    slug = re.sub(r"[^a-z0-9]+", "-", str(value or "").strip().lower()).strip("-")
    return slug or "workspace"


def token_hash(token: str) -> str:
    return hashlib.sha256(token.encode("utf-8")).hexdigest()


class LucasCloudStore:
    """SQLite-backed storage for the hosted LUCAS platform.

    The store keeps durable metadata in SQLite and writes photo bytes to the
    configured object root. It is designed to be lifted into Postgres/S3 later
    without changing the higher-level service contract.
    """

    def __init__(self, data_root: Path | str):
        self.data_root = Path(data_root).expanduser()
        self.db_path = self.data_root / "lucas_cloud.sqlite3"
        self.object_root = self.data_root / "objects"

    def connect(self) -> sqlite3.Connection:
        self.data_root.mkdir(parents=True, exist_ok=True)
        self.object_root.mkdir(parents=True, exist_ok=True)
        connection = sqlite3.connect(self.db_path)
        connection.row_factory = sqlite3.Row
        connection.execute("PRAGMA foreign_keys = ON")
        return connection

    def initialize(self) -> None:
        with self.connect() as db:
            db.executescript(
                """
                CREATE TABLE IF NOT EXISTS schema_meta (
                    key TEXT PRIMARY KEY,
                    value TEXT NOT NULL
                );

                CREATE TABLE IF NOT EXISTS workspaces (
                    id TEXT PRIMARY KEY,
                    slug TEXT NOT NULL UNIQUE,
                    display_name TEXT NOT NULL,
                    status TEXT NOT NULL DEFAULT 'active',
                    created_at TEXT NOT NULL,
                    updated_at TEXT NOT NULL
                );

                CREATE TABLE IF NOT EXISTS workspace_tokens (
                    id TEXT PRIMARY KEY,
                    workspace_id TEXT NOT NULL REFERENCES workspaces(id) ON DELETE CASCADE,
                    label TEXT NOT NULL,
                    token_hash TEXT NOT NULL UNIQUE,
                    created_at TEXT NOT NULL,
                    revoked_at TEXT NOT NULL DEFAULT ''
                );

                CREATE TABLE IF NOT EXISTS inventory_items (
                    id TEXT PRIMARY KEY,
                    workspace_id TEXT NOT NULL REFERENCES workspaces(id) ON DELETE CASCADE,
                    external_key TEXT NOT NULL,
                    status TEXT NOT NULL,
                    assigned_person TEXT NOT NULL DEFAULT '',
                    cert_number TEXT NOT NULL DEFAULT '',
                    item_id TEXT NOT NULL DEFAULT '',
                    card_title TEXT NOT NULL DEFAULT '',
                    purchase_price REAL,
                    raw_json TEXT NOT NULL DEFAULT '{}',
                    created_at TEXT NOT NULL,
                    updated_at TEXT NOT NULL,
                    UNIQUE(workspace_id, external_key)
                );

                CREATE TABLE IF NOT EXISTS photos (
                    id TEXT PRIMARY KEY,
                    workspace_id TEXT NOT NULL REFERENCES workspaces(id) ON DELETE CASCADE,
                    uploader_name TEXT NOT NULL DEFAULT '',
                    assigned_person TEXT NOT NULL DEFAULT '',
                    original_filename TEXT NOT NULL DEFAULT '',
                    content_type TEXT NOT NULL DEFAULT 'application/octet-stream',
                    sha256 TEXT NOT NULL,
                    size_bytes INTEGER NOT NULL,
                    storage_path TEXT NOT NULL,
                    status TEXT NOT NULL,
                    linked_inventory_id TEXT NOT NULL DEFAULT '',
                    uploaded_at TEXT NOT NULL,
                    updated_at TEXT NOT NULL,
                    raw_json TEXT NOT NULL DEFAULT '{}'
                );

                CREATE INDEX IF NOT EXISTS idx_photos_workspace_status
                    ON photos(workspace_id, status, uploaded_at);

                CREATE INDEX IF NOT EXISTS idx_photos_workspace_hash
                    ON photos(workspace_id, sha256);

                CREATE TABLE IF NOT EXISTS audit_events (
                    id TEXT PRIMARY KEY,
                    workspace_id TEXT NOT NULL DEFAULT '',
                    actor TEXT NOT NULL DEFAULT '',
                    action TEXT NOT NULL,
                    entity_type TEXT NOT NULL DEFAULT '',
                    entity_id TEXT NOT NULL DEFAULT '',
                    details_json TEXT NOT NULL DEFAULT '{}',
                    created_at TEXT NOT NULL
                );
                """
            )
            db.execute(
                "INSERT OR REPLACE INTO schema_meta(key, value) VALUES('schema_version', ?)",
                (str(SCHEMA_VERSION),),
            )

    def create_workspace(self, slug: str, display_name: str, token_label: str = "mobile") -> tuple[Workspace, str]:
        self.initialize()
        now = utc_now()
        workspace_id = str(uuid.uuid4())
        normalized_slug = safe_slug(slug or display_name)
        token = secrets.token_urlsafe(32)
        with self.connect() as db:
            db.execute(
                """
                INSERT INTO workspaces(id, slug, display_name, status, created_at, updated_at)
                VALUES (?, ?, ?, 'active', ?, ?)
                """,
                (workspace_id, normalized_slug, display_name.strip() or normalized_slug, now, now),
            )
            db.execute(
                """
                INSERT INTO workspace_tokens(id, workspace_id, label, token_hash, created_at)
                VALUES (?, ?, ?, ?, ?)
                """,
                (str(uuid.uuid4()), workspace_id, token_label, token_hash(token), now),
            )
            self.record_audit(db, workspace_id, "system", "workspace.create", "workspace", workspace_id, {"slug": normalized_slug})
        return self.get_workspace(normalized_slug), token

    def get_workspace(self, slug: str) -> Workspace:
        self.initialize()
        with self.connect() as db:
            row = db.execute("SELECT * FROM workspaces WHERE slug = ?", (safe_slug(slug),)).fetchone()
        if row is None:
            raise KeyError(f"Unknown workspace: {slug}")
        return Workspace(
            id=str(row["id"]),
            slug=str(row["slug"]),
            display_name=str(row["display_name"]),
            status=str(row["status"]),
            created_at=str(row["created_at"]),
        )

    def verify_workspace_token(self, slug: str, token: str) -> Workspace:
        workspace = self.get_workspace(slug)
        if not token:
            raise PermissionError("Missing workspace token.")
        hashed = token_hash(token)
        with self.connect() as db:
            row = db.execute(
                """
                SELECT 1 FROM workspace_tokens
                WHERE workspace_id = ? AND token_hash = ? AND revoked_at = ''
                """,
                (workspace.id, hashed),
            ).fetchone()
        if row is None:
            raise PermissionError("Invalid workspace token.")
        return workspace

    def store_photo(
        self,
        workspace: Workspace,
        image_bytes: bytes,
        original_filename: str,
        content_type: str,
        uploader_name: str = "",
        assigned_person: str = "",
        raw: dict[str, Any] | None = None,
    ) -> PhotoRecord:
        self.initialize()
        if not image_bytes:
            raise ValueError("Photo is empty.")
        sha = hashlib.sha256(image_bytes).hexdigest()
        now = utc_now()
        photo_id = str(uuid.uuid4())
        extension = self._photo_extension(original_filename, content_type)
        relative = Path("workspaces") / workspace.slug / "photos" / now[:4] / now[5:7] / f"{photo_id}{extension}"
        destination = self.object_root / relative
        destination.parent.mkdir(parents=True, exist_ok=True)
        destination.write_bytes(image_bytes)
        record = PhotoRecord(
            id=photo_id,
            workspace_id=workspace.id,
            uploader_name=str(uploader_name or "").strip(),
            assigned_person=str(assigned_person or "").strip(),
            original_filename=str(original_filename or "").strip() or f"{photo_id}{extension}",
            content_type=str(content_type or "application/octet-stream").strip(),
            sha256=sha,
            size_bytes=len(image_bytes),
            storage_path=relative.as_posix(),
            status="pending_scan",
            linked_inventory_id="",
            uploaded_at=now,
            updated_at=now,
        )
        with self.connect() as db:
            db.execute(
                """
                INSERT INTO photos(
                    id, workspace_id, uploader_name, assigned_person, original_filename,
                    content_type, sha256, size_bytes, storage_path, status,
                    linked_inventory_id, uploaded_at, updated_at, raw_json
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    record.id,
                    record.workspace_id,
                    record.uploader_name,
                    record.assigned_person,
                    record.original_filename,
                    record.content_type,
                    record.sha256,
                    record.size_bytes,
                    record.storage_path,
                    record.status,
                    record.linked_inventory_id,
                    record.uploaded_at,
                    record.updated_at,
                    json.dumps(raw or {}, sort_keys=True),
                ),
            )
            self.record_audit(db, workspace.id, uploader_name, "photo.upload", "photo", photo_id, {"assigned_person": assigned_person})
        return record

    def list_photos(self, workspace: Workspace, status: str = "", limit: int = 100) -> list[PhotoRecord]:
        self.initialize()
        limit = max(1, min(int(limit or 100), 500))
        with self.connect() as db:
            if status:
                rows = db.execute(
                    "SELECT * FROM photos WHERE workspace_id = ? AND status = ? ORDER BY uploaded_at ASC LIMIT ?",
                    (workspace.id, status, limit),
                ).fetchall()
            else:
                rows = db.execute(
                    "SELECT * FROM photos WHERE workspace_id = ? ORDER BY uploaded_at DESC LIMIT ?",
                    (workspace.id, limit),
                ).fetchall()
        return [self._photo_from_row(row) for row in rows]

    def update_photo_status(
        self,
        workspace: Workspace,
        photo_id: str,
        status: str,
        linked_inventory_id: str = "",
        actor: str = "",
    ) -> PhotoRecord:
        allowed = {"pending_scan", "scan_claimed", "linked", "no_matching_inventory", "sold_inventory", "archived", "deleted"}
        if status not in allowed:
            raise ValueError(f"Invalid photo status: {status}")
        now = utc_now()
        with self.connect() as db:
            existing = db.execute(
                "SELECT * FROM photos WHERE workspace_id = ? AND id = ?",
                (workspace.id, photo_id),
            ).fetchone()
            if existing is None:
                raise KeyError(f"Unknown photo: {photo_id}")
            db.execute(
                """
                UPDATE photos
                SET status = ?, linked_inventory_id = COALESCE(NULLIF(?, ''), linked_inventory_id), updated_at = ?
                WHERE workspace_id = ? AND id = ?
                """,
                (status, linked_inventory_id, now, workspace.id, photo_id),
            )
            self.record_audit(db, workspace.id, actor, "photo.status", "photo", photo_id, {"status": status, "linked_inventory_id": linked_inventory_id})
            row = db.execute("SELECT * FROM photos WHERE workspace_id = ? AND id = ?", (workspace.id, photo_id)).fetchone()
        return self._photo_from_row(row)

    def upsert_inventory_item(self, workspace: Workspace, payload: dict[str, Any], actor: str = "") -> InventoryItem:
        self.initialize()
        external_key = str(payload.get("external_key") or payload.get("inventory_key") or payload.get("cert_number") or payload.get("item_id") or "").strip()
        if not external_key:
            external_key = str(uuid.uuid4())
        now = utc_now()
        with self.connect() as db:
            row = db.execute(
                "SELECT * FROM inventory_items WHERE workspace_id = ? AND external_key = ?",
                (workspace.id, external_key),
            ).fetchone()
            item_id = str(row["id"]) if row else str(uuid.uuid4())
            created_at = str(row["created_at"]) if row else now
            db.execute(
                """
                INSERT INTO inventory_items(
                    id, workspace_id, external_key, status, assigned_person, cert_number,
                    item_id, card_title, purchase_price, raw_json, created_at, updated_at
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(workspace_id, external_key) DO UPDATE SET
                    status = excluded.status,
                    assigned_person = excluded.assigned_person,
                    cert_number = excluded.cert_number,
                    item_id = excluded.item_id,
                    card_title = excluded.card_title,
                    purchase_price = excluded.purchase_price,
                    raw_json = excluded.raw_json,
                    updated_at = excluded.updated_at
                """,
                (
                    item_id,
                    workspace.id,
                    external_key,
                    str(payload.get("status") or "active").strip().lower(),
                    str(payload.get("assigned_person") or payload.get("person") or "").strip(),
                    str(payload.get("cert_number") or "").strip(),
                    str(payload.get("item_id") or "").strip(),
                    str(payload.get("card_title") or payload.get("title") or "").strip(),
                    self._float_or_none(payload.get("purchase_price")),
                    json.dumps(payload, sort_keys=True),
                    created_at,
                    now,
                ),
            )
            self.record_audit(db, workspace.id, actor, "inventory.upsert", "inventory", item_id, {"external_key": external_key})
            saved = db.execute("SELECT * FROM inventory_items WHERE workspace_id = ? AND external_key = ?", (workspace.id, external_key)).fetchone()
        return self._inventory_from_row(saved)

    def record_audit(
        self,
        db: sqlite3.Connection,
        workspace_id: str,
        actor: str,
        action: str,
        entity_type: str = "",
        entity_id: str = "",
        details: dict[str, Any] | None = None,
    ) -> None:
        db.execute(
            """
            INSERT INTO audit_events(id, workspace_id, actor, action, entity_type, entity_id, details_json, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                str(uuid.uuid4()),
                workspace_id,
                str(actor or ""),
                action,
                entity_type,
                entity_id,
                json.dumps(details or {}, sort_keys=True),
                utc_now(),
            ),
        )

    def photo_bytes(self, record: PhotoRecord) -> bytes:
        path = self.object_root / record.storage_path
        if not path.exists() or not path.is_file():
            raise FileNotFoundError(record.storage_path)
        return path.read_bytes()

    def _photo_extension(self, filename: str, content_type: str) -> str:
        suffix = Path(str(filename or "")).suffix.lower()
        if suffix in {".jpg", ".jpeg", ".png", ".webp", ".heic", ".heif"}:
            return suffix
        content = str(content_type or "").lower()
        if "png" in content:
            return ".png"
        if "webp" in content:
            return ".webp"
        if "heic" in content or "heif" in content:
            return ".heic"
        return ".jpg"

    def _photo_from_row(self, row: sqlite3.Row) -> PhotoRecord:
        return PhotoRecord(
            id=str(row["id"]),
            workspace_id=str(row["workspace_id"]),
            uploader_name=str(row["uploader_name"]),
            assigned_person=str(row["assigned_person"]),
            original_filename=str(row["original_filename"]),
            content_type=str(row["content_type"]),
            sha256=str(row["sha256"]),
            size_bytes=int(row["size_bytes"]),
            storage_path=str(row["storage_path"]),
            status=str(row["status"]),
            linked_inventory_id=str(row["linked_inventory_id"]),
            uploaded_at=str(row["uploaded_at"]),
            updated_at=str(row["updated_at"]),
        )

    def _inventory_from_row(self, row: sqlite3.Row) -> InventoryItem:
        return InventoryItem(
            id=str(row["id"]),
            workspace_id=str(row["workspace_id"]),
            external_key=str(row["external_key"]),
            status=str(row["status"]),
            assigned_person=str(row["assigned_person"]),
            cert_number=str(row["cert_number"]),
            item_id=str(row["item_id"]),
            card_title=str(row["card_title"]),
            purchase_price=self._float_or_none(row["purchase_price"]),
            created_at=str(row["created_at"]),
            updated_at=str(row["updated_at"]),
        )

    def _float_or_none(self, value: Any) -> float | None:
        if value in ("", None):
            return None
        try:
            return float(value)
        except (TypeError, ValueError):
            return None


def decode_data_url_image(value: str) -> tuple[str, bytes]:
    text = str(value or "").strip()
    if "," not in text or not text.startswith("data:"):
        return "application/octet-stream", base64.b64decode(text)
    header, encoded = text.split(",", 1)
    content_type = header[5:].split(";", 1)[0] or "application/octet-stream"
    return content_type, base64.b64decode(encoded)
