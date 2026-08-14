from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class Workspace:
    id: str
    slug: str
    display_name: str
    status: str
    created_at: str


@dataclass(frozen=True)
class PhotoRecord:
    id: str
    workspace_id: str
    uploader_name: str
    assigned_person: str
    original_filename: str
    content_type: str
    sha256: str
    size_bytes: int
    storage_path: str
    status: str
    linked_inventory_id: str
    uploaded_at: str
    updated_at: str


@dataclass(frozen=True)
class InventoryItem:
    id: str
    workspace_id: str
    external_key: str
    status: str
    assigned_person: str
    cert_number: str
    item_id: str
    card_title: str
    purchase_price: float | None
    created_at: str
    updated_at: str
