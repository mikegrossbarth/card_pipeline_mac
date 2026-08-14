from __future__ import annotations

from dataclasses import asdict
from pathlib import Path
from typing import Any

from .models import PhotoRecord, Workspace
from .store import LucasCloudStore, decode_data_url_image


class LucasCloudService:
    """Application service for hosted LUCAS workflows."""

    def __init__(self, data_root: Path | str):
        self.store = LucasCloudStore(data_root)
        self.store.initialize()

    def create_workspace(self, slug: str, display_name: str) -> dict[str, Any]:
        workspace, token = self.store.create_workspace(slug, display_name)
        return {"workspace": asdict(workspace), "token": token}

    def upload_mobile_photos(self, workspace_slug: str, token: str, payload: dict[str, Any]) -> dict[str, Any]:
        workspace = self.store.verify_workspace_token(workspace_slug, token)
        images = payload.get("images")
        if not isinstance(images, list):
            image = payload.get("image")
            images = [{"image": image, "name": payload.get("name") or payload.get("filename") or "mobile-photo.jpg"}] if image else []
        if not images:
            return {"ok": False, "error": "Take or choose at least one photo."}
        if len(images) > 24:
            return {"ok": False, "error": "Upload 24 photos or fewer at a time."}
        uploaded: list[PhotoRecord] = []
        for index, item in enumerate(images, start=1):
            if not isinstance(item, dict):
                continue
            image_value = str(item.get("image") or "").strip()
            if not image_value:
                continue
            content_type, image_bytes = decode_data_url_image(image_value)
            if not str(content_type or "").startswith("image/"):
                return {"ok": False, "error": f"Photo {index} is not an image."}
            if len(image_bytes) > 16 * 1024 * 1024:
                return {"ok": False, "error": f"Photo {index} is too large."}
            uploaded.append(
                self.store.store_photo(
                    workspace,
                    image_bytes,
                    str(item.get("name") or item.get("filename") or f"photo-{index}.jpg"),
                    content_type,
                    uploader_name=str(payload.get("uploader") or payload.get("uploader_name") or ""),
                    assigned_person=str(payload.get("assigned_person") or payload.get("person") or ""),
                    raw={"client_id": payload.get("client_id") or payload.get("clientId") or ""},
                )
            )
        if not uploaded:
            return {"ok": False, "error": "No usable photos were uploaded."}
        return {
            "ok": True,
            "saved": len(uploaded),
            "photos": [asdict(photo) for photo in uploaded],
            "workspace": workspace.slug,
        }

    def pending_photos(self, workspace_slug: str, token: str, limit: int = 100) -> dict[str, Any]:
        workspace = self.store.verify_workspace_token(workspace_slug, token)
        photos = self.store.list_photos(workspace, status="pending_scan", limit=limit)
        return {"ok": True, "workspace": workspace.slug, "photos": [asdict(photo) for photo in photos]}

    def mark_photo_status(self, workspace_slug: str, token: str, photo_id: str, payload: dict[str, Any]) -> dict[str, Any]:
        workspace = self.store.verify_workspace_token(workspace_slug, token)
        record = self.store.update_photo_status(
            workspace,
            photo_id,
            str(payload.get("status") or ""),
            linked_inventory_id=str(payload.get("linked_inventory_id") or payload.get("inventory_id") or ""),
            actor=str(payload.get("actor") or payload.get("user") or ""),
        )
        return {"ok": True, "photo": asdict(record)}

    def upsert_inventory(self, workspace_slug: str, token: str, payload: dict[str, Any]) -> dict[str, Any]:
        workspace = self.store.verify_workspace_token(workspace_slug, token)
        item = self.store.upsert_inventory_item(workspace, payload, actor=str(payload.get("actor") or payload.get("user") or ""))
        return {"ok": True, "item": asdict(item)}

    def workspace_from_token(self, workspace_slug: str, token: str) -> Workspace:
        return self.store.verify_workspace_token(workspace_slug, token)
