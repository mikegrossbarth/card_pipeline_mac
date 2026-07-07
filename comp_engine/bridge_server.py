from __future__ import annotations

import json
import ipaddress
import mimetypes
import os
import re
import socket
import sys
import threading
import time
import uuid
from dataclasses import asdict
from datetime import datetime, timezone
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Callable
from urllib.parse import parse_qs, unquote, urlparse

from cardladder_ocr import extract_cl_value_from_data_url
from cy_automation.cy_macos import CYMacOSAdapter
from workbook_io import WorkbookRow
import assignment_engine

BRIDGE_VERSION = "2026-06-17-cardladder-no-blind-grader-option-v22"
EXPECTED_CARDLADDER_EXTENSION_VERSION = "2026-06-17-no-blind-grader-option-v22"
EXPECTED_CARDLADDER_MANIFEST_VERSION = "0.1.4"
DEBUG_DIR = Path(__file__).resolve().parent.parent / "work" / "cardladder-bridge"
DEBUG_LOG = DEBUG_DIR / "bridge.log"
MOBILE_APP_DIR = Path(__file__).resolve().parent.parent / "mobile_app"
MOBILE_PROFILE_LABELS = {
    "team": "LUCAS Team",
    "personal": "LUCAS Personal",
}
BRIDGE_LOCAL_ONLY_PATH_PREFIXES = (
    "/command",
    "/status",
    "/ack",
    "/result/cardladder",
    "/ocr/cardladder",
    "/finish/cardladder",
    "/source/google-keep",
)
COMP_STRATEGY_AVERAGE = "average_last_5"
COMP_STRATEGY_HIGH = "highest_last_5"
COMP_STRATEGY_LOW = "lowest_last_5"
COMP_STRATEGY_STALE_NEWEST = "stale_newest_else_average"
COMP_STRATEGY_LABELS = {
    COMP_STRATEGY_AVERAGE: "Average last 5",
    COMP_STRATEGY_HIGH: "Highest of last 5",
    COMP_STRATEGY_LOW: "Lowest of last 5",
    COMP_STRATEGY_STALE_NEWEST: "Date weighted",
}
_CY_ADAPTER: CYMacOSAdapter | None = None
_CY_ADAPTER_LOCK = threading.Lock()
_CY_LOOKUP_LOCK = threading.Lock()


def is_loopback_address(value: str) -> bool:
    host = str(value or "").strip().lower()
    if host in {"localhost", "ip6-localhost"}:
        return True
    try:
        return ipaddress.ip_address(host).is_loopback
    except ValueError:
        return False


def request_origin_allowed(origin: str, host_header: str) -> bool:
    origin = str(origin or "").strip()
    if not origin:
        return True
    parsed = urlparse(origin)
    if parsed.scheme == "chrome-extension":
        return True
    if parsed.scheme not in {"http", "https"}:
        return False
    request_host = str(host_header or "").strip().lower()
    if parsed.netloc.lower() == request_host:
        return True
    return is_loopback_address(parsed.hostname or "")


def fill_missing_category_from_title(row: WorkbookRow) -> None:
    if str(getattr(row, "category", "") or "").strip():
        return
    parsed = assignment_engine.parse_card_for_matching(getattr(row, "card_title", "") or "")
    sport = str(parsed.get("sport") or "").strip()
    if sport:
        row.category = sport


class BridgeState:
    def __init__(self) -> None:
        self.lock = threading.Lock()
        self.instance_id = uuid.uuid4().hex
        self.rows: list[WorkbookRow] = []
        self.command: dict | None = None
        self.command_id = int(time.time() * 1000)
        self.last_seen_extension = ""
        self.extension_version = ""
        self.extension_manifest_version = ""
        self.extension_name = ""
        self.extension_url = ""
        self.last_result_extension_version = ""
        self.cardladder_running = False
        self.cancel_requested = False
        self.comp_strategy = COMP_STRATEGY_AVERAGE
        self.updated_row_ids: set[int] = set()
        self.on_update: Callable[[], None] | None = None
        self.mobile_pin_provider: Callable[[], str] | None = None
        self.mobile_inventory_search: Callable[[dict], dict] | None = None
        self.mobile_inventory_add: Callable[[dict], dict] | None = None
        self.mobile_inventory_mark_sold: Callable[[dict], dict] | None = None
        self.mobile_card_identify: Callable[[dict], dict] | None = None
        self.mobile_profit_summary: Callable[[dict], dict] | None = None
        self.mobile_expense_add: Callable[[dict], dict] | None = None
        self.mobile_payouts: Callable[[dict], dict] | None = None
        self.mobile_queue_sync: Callable[[dict], dict] | None = None
        self.mobile_inventory_photo_resolver: Callable[[str], tuple[bytes, str] | None] | None = None
        self.instagram_media_token = uuid.uuid4().hex
        self.instagram_media_resolver: Callable[[str], tuple[bytes, str] | None] | None = None
        self.keep_note_sources: list[dict[str, str]] = []
        self.last_keep_sync: dict[str, str] = {}
        self.cy_lookup_inflight: set[int] = set()
        self.cy_lookup_pending: set[int] = set()
        self.cy_lookup_generation = 0
        self.cardladder_allows_cy = False
        self.cy_batch_running = False

    def set_rows(self, rows: list[WorkbookRow]) -> None:
        with self.lock:
            self.rows = rows
            self.updated_row_ids = set()

    def set_comp_strategy(self, strategy: str) -> None:
        with self.lock:
            self.comp_strategy = strategy if strategy in COMP_STRATEGY_LABELS else COMP_STRATEGY_AVERAGE

    def register_keep_note_sources(self, sources: list[dict[str, object]]) -> None:
        normalized: list[dict[str, str]] = []
        for source in sources:
            url = str(source.get("url") or "").strip()
            path = str(source.get("path") or source.get("file") or "").strip()
            name = str(source.get("name") or Path(path).stem or "Google Keep note")
            if url and path:
                normalized.append({"url": url, "path": path, "name": name})
        with self.lock:
            self.keep_note_sources = normalized

    def post_google_keep_note(self, payload: dict) -> dict:
        text = str(payload.get("text") or "").strip()
        url = str(payload.get("url") or "").strip()
        title = str(payload.get("title") or "").strip() or "Google Keep note"
        synced_at = str(payload.get("synced_at") or payload.get("syncedAt") or "").strip() or datetime.now(timezone.utc).isoformat()
        if not text:
            return {"ok": False, "saved": 0, "error": "Google Keep note text was empty."}
        with self.lock:
            sources = list(self.keep_note_sources)
        matches = [source for source in sources if keep_urls_match(url, source.get("url", ""))]
        saved_paths: list[str] = []
        for source in matches:
            path = Path(source["path"]).expanduser()
            try:
                path.parent.mkdir(parents=True, exist_ok=True)
                path.write_text(text + "\n", encoding="utf-8")
                saved_paths.append(str(path))
            except OSError:
                continue
        with self.lock:
            self.last_keep_sync = {"url": url, "title": title, "syncedAt": synced_at, "saved": str(len(saved_paths))}
        return {"ok": bool(saved_paths), "saved": len(saved_paths), "paths": saved_paths}

    def mobile_auth_ok(self, payload: dict | None = None, query: dict[str, list[str]] | None = None) -> bool:
        provider = self.mobile_pin_provider
        expected = provider() if provider else ""
        if not expected:
            return False
        candidate = ""
        if payload:
            candidate = str(payload.get("pin") or payload.get("token") or "").strip()
        if not candidate and query:
            candidate = str(query.get("pin", [""])[0] or query.get("token", [""])[0]).strip()
        return bool(candidate and candidate == expected)

    def mobile_config(self) -> dict:
        return {
            "ok": True,
            "service": "lucas-mobile",
            "requiresPin": bool(self.mobile_pin_provider),
            "photoSearch": True,
            "inventorySold": True,
            "profit": True,
            "expenses": True,
            "payouts": True,
        }

    def search_mobile_inventory(self, payload: dict) -> dict:
        if not self.mobile_auth_ok(payload):
            return {"ok": False, "error": "Invalid mobile PIN."}
        if not self.mobile_inventory_search:
            return {"ok": False, "error": "Inventory search is not available."}
        return self.mobile_inventory_search(payload)

    def add_mobile_inventory(self, payload: dict) -> dict:
        if not self.mobile_auth_ok(payload):
            return {"ok": False, "error": "Invalid mobile PIN."}
        if not self.mobile_inventory_add:
            return {"ok": False, "error": "Inventory add is not available."}
        return self.mobile_inventory_add(payload)

    def mark_mobile_inventory_sold(self, payload: dict) -> dict:
        if not self.mobile_auth_ok(payload):
            return {"ok": False, "error": "Invalid mobile PIN."}
        if not self.mobile_inventory_mark_sold:
            return {"ok": False, "error": "Inventory sale is not available."}
        return self.mobile_inventory_mark_sold(payload)

    def identify_mobile_card(self, payload: dict) -> dict:
        if not self.mobile_auth_ok(payload):
            return {"ok": False, "error": "Invalid mobile PIN."}
        if not self.mobile_card_identify:
            return {"ok": False, "error": "Photo card search is not available."}
        return self.mobile_card_identify(payload)

    def get_mobile_profit_summary(self, payload: dict) -> dict:
        if not self.mobile_auth_ok(payload):
            return {"ok": False, "error": "Invalid mobile PIN."}
        if not self.mobile_profit_summary:
            return {"ok": False, "error": "Profit view is not available."}
        return self.mobile_profit_summary(payload)

    def add_mobile_expense(self, payload: dict) -> dict:
        if not self.mobile_auth_ok(payload):
            return {"ok": False, "error": "Invalid mobile PIN."}
        if not self.mobile_expense_add:
            return {"ok": False, "error": "Expense entry is not available."}
        return self.mobile_expense_add(payload)

    def get_mobile_payouts(self, payload: dict) -> dict:
        if not self.mobile_auth_ok(payload):
            return {"ok": False, "error": "Invalid mobile PIN."}
        if not self.mobile_payouts:
            return {"ok": False, "error": "Payout view is not available."}
        return self.mobile_payouts(payload)

    def sync_mobile_queue(self, payload: dict) -> dict:
        if not self.mobile_auth_ok(payload):
            return {"ok": False, "error": "Invalid mobile PIN."}
        if not self.mobile_queue_sync:
            return {"ok": False, "error": "Mobile queue sync is not available."}
        return self.mobile_queue_sync(payload)

    def mobile_inventory_photo_path(self, photo_id: str, filename: str = "photo.jpg") -> str:
        safe_name = re.sub(r"[^A-Za-z0-9._-]+", "-", str(filename or "photo.jpg")).strip("-") or "photo.jpg"
        return f"/mobile/api/inventory/photo/{photo_id}/{safe_name}"

    def get_mobile_inventory_photo(self, payload: dict | None, query: dict[str, list[str]], photo_id: str) -> tuple[bytes, str] | None:
        if not self.mobile_auth_ok(payload, query):
            return None
        resolver = self.mobile_inventory_photo_resolver
        if resolver is None:
            return None
        return resolver(photo_id)

    def instagram_media_path(self, photo_id: str, filename: str = "photo.jpg") -> str:
        token = self.instagram_media_token
        safe_name = re.sub(r"[^A-Za-z0-9._-]+", "-", str(filename or "photo.jpg")).strip("-") or "photo.jpg"
        return f"/instagram/media/{token}/{photo_id}/{safe_name}"

    def get_instagram_media(self, token: str, photo_id: str) -> tuple[bytes, str] | None:
        if not token or token != self.instagram_media_token:
            return None
        resolver = self.instagram_media_resolver
        if resolver is None:
            return None
        return resolver(photo_id)

    def start_all_comps(self, requery_all: bool = False, allow_deferred_cy: bool = False) -> int:
        with self.lock:
            self.command_id += 1
            self.cardladder_allows_cy = bool(allow_deferred_cy)
            eligible_rows = [
                row
                for row in self.rows
                if row.cert_number and row.grader and (requery_all or not row_has_comp_data(row))
            ]
            queue = [
                {
                    "excelRow": row.excel_row,
                    "certNumber": row.cert_number,
                    "grader": row.grader,
                    "cardTitle": row.card_title,
                }
                for row in eligible_rows
            ]
            for row in eligible_rows:
                row.status = "Queued"
            if not queue:
                self.command = None
                self.cardladder_running = False
                self.cardladder_allows_cy = False
                debug_log(f"start_all_comps command={self.command_id} eligible=0 requery_all={requery_all} allow_deferred_cy={allow_deferred_cy}")
                return self.command_id
            self.command = {
                "id": self.command_id,
                "type": "RUN_ALL_COMPS",
                "sources": ["cardladder"],
                "queue": queue,
                "createdAt": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()),
            }
            self.cardladder_running = True
            self.cancel_requested = False
            debug_log(f"start_all_comps command={self.command_id} queued={len(queue)} requery_all={requery_all} allow_deferred_cy={allow_deferred_cy}")
            return self.command_id

    def start_cy_lookups(self, rows: list[WorkbookRow], defer: bool = False) -> int:
        cy_lookups: list[tuple[int, str, str, int]] = []
        with self.lock:
            self.command_id += 1
            self.cy_lookup_generation += 1
            cy_generation = self.cy_lookup_generation
            if defer or self.cardladder_running:
                self.cardladder_allows_cy = True
            for row in rows:
                candidate = self._cy_lookup_candidate(row, force=True)
                if candidate is None:
                    continue
                row.status = "CY queued"
                if defer or self.cardladder_running:
                    self.cy_lookup_pending.add(row.excel_row)
                    continue
                self.cy_lookup_inflight.add(row.excel_row)
                cy_lookups.append((*candidate, cy_generation))
            if cy_lookups or self.cy_lookup_pending:
                self.cy_batch_running = True
            debug_log(
                f"start_cy_lookups command={self.command_id} started={len(cy_lookups)} "
                f"pending={len(self.cy_lookup_pending)} defer={defer}"
            )
        for cy_lookup in cy_lookups:
            threading.Thread(target=self._cy_lookup_worker, args=cy_lookup, daemon=True).start()
        if self.on_update:
            self.on_update()
        return self.command_id

    def request_cancel(self) -> None:
        should_close_cy = False
        with self.lock:
            self.cancel_requested = True
            self.command = None
            self.cardladder_running = False
            self.cardladder_allows_cy = False
            self.cy_lookup_generation += 1
            self.cy_lookup_pending.clear()
            should_close_cy = bool(self.cy_lookup_inflight or self.cy_batch_running)
            self.cy_lookup_inflight.clear()
            self.cy_batch_running = False
            for row in self.rows:
                if row.status == "Queued":
                    row.status = "Card Ladder cancelled"
                elif row.status == "CY queued":
                    row.status = "CY cancelled"
        if should_close_cy:
            close_cy_adapter()
        if self.on_update:
            self.on_update()

    def extension_poll(self, metadata: dict[str, str] | None = None) -> dict:
        with self.lock:
            self.last_seen_extension = time.strftime("%H:%M:%S")
            extension_version = ""
            if metadata:
                extension_version = metadata.get("extensionVersion") or ""
                self.extension_version = metadata.get("extensionVersion") or self.extension_version
                self.extension_manifest_version = metadata.get("manifestVersion") or self.extension_manifest_version
                self.extension_name = metadata.get("extensionName") or self.extension_name
                self.extension_url = metadata.get("extensionUrl") or self.extension_url
                debug_log(
                    "extension_poll "
                    f"version={self.extension_version or 'unknown'} "
                    f"manifest={self.extension_manifest_version or 'unknown'} "
                    f"name={self.extension_name or 'unknown'}"
                )
            command = self.command if extension_version == EXPECTED_CARDLADDER_EXTENSION_VERSION else None
            return {
                "instanceId": self.instance_id,
                "command": command,
                "keepNoteSources": list(self.keep_note_sources),
                "lastKeepSync": dict(self.last_keep_sync),
            }

    def acknowledge_command(self, command_id: int) -> None:
        with self.lock:
            if self.command and self.command.get("id") == command_id:
                self.command = None

    def post_cardladder_result(self, result: dict) -> None:
        DEBUG_DIR.mkdir(parents=True, exist_ok=True)
        debug_stamp = f"{time.strftime('%Y%m%d-%H%M%S')}-{time.time_ns() % 1_000_000_000:09d}"
        (DEBUG_DIR / f"result-{debug_stamp}.json").write_text(
            json.dumps(result, indent=2),
            encoding="utf-8",
        )
        cy_lookup: tuple[int, str, str, int] | None = None
        with self.lock:
            result_extension_version = str(result.get("extensionVersion") or "")
            if result_extension_version:
                self.last_result_extension_version = result_extension_version
                self.extension_version = result_extension_version
            cert = str(result.get("certNumber") or "")
            excel_row = int(result.get("excelRow") or 0)
            target_row = next((row for row in self.rows if excel_row and row.excel_row == excel_row), None)
            if target_row is None and cert:
                target_row = next((row for row in self.rows if row.cert_number == cert), None)
            if target_row is not None:
                self._apply_cardladder_result_to_row(target_row, result)
                self.updated_row_ids.add(id(target_row))
                if self.cardladder_allows_cy or target_row.excel_row in self.cy_lookup_pending:
                    cy_lookup = self._queue_or_prepare_cy_lookup(target_row)
                debug_log(f"cardladder_result row={target_row.excel_row} cert={target_row.cert_number} cy_lookup={bool(cy_lookup)}")
            else:
                debug_log(f"cardladder_result no_target excel_row={excel_row} cert={cert}")
        if cy_lookup:
            threading.Thread(target=self._cy_lookup_worker, args=cy_lookup, daemon=True).start()
        if self.on_update:
            self.on_update()

    def _cy_lookup_candidate(self, row: WorkbookRow, force: bool = False) -> tuple[int, str, str] | None:
        if not cy_lookup_enabled():
            debug_log(f"cy_lookup_skip row={row.excel_row} reason=disabled")
            return None
        if row.cy_value is not None and not force:
            debug_log(f"cy_lookup_skip row={row.excel_row} reason=has_value")
            return None
        if row.excel_row in self.cy_lookup_inflight:
            debug_log(f"cy_lookup_skip row={row.excel_row} reason=inflight")
            return None
        cert_number = str(row.cert_number or "").strip()
        slab_type = clean_grader(row.grader)
        if not cert_number or slab_type not in {"PSA", "BGS", "CGC", "SGC"}:
            debug_log(f"cy_lookup_skip row={row.excel_row} reason=missing_or_unsupported cert={cert_number} slab={slab_type}")
            return None
        return row.excel_row, cert_number, slab_type

    def _queue_or_prepare_cy_lookup(self, row: WorkbookRow) -> tuple[int, str, str, int] | None:
        candidate = self._cy_lookup_candidate(row)
        if candidate is None:
            return None
        if self.cardladder_running:
            self.cy_lookup_pending.add(row.excel_row)
            debug_log(f"cy_lookup_pending row={row.excel_row} cert={row.cert_number}")
            return None
        self.cy_lookup_inflight.add(row.excel_row)
        return (*candidate, self.cy_lookup_generation)

    def _cy_lookup_worker(self, excel_row: int, cert_number: str, slab_type: str, generation: int) -> None:
        value = None
        confidence = None
        message = ""
        should_close = False
        with self.lock:
            if generation != self.cy_lookup_generation or excel_row not in self.cy_lookup_inflight:
                debug_log(f"cy_lookup_cancelled_before_start row={excel_row} cert={cert_number}")
                return
        try:
            result = lookup_cy_buy_price(cert_number, slab_type)
            if len(result) == 3:
                value, confidence, message = result
            else:
                value, message = result
        except Exception as error:
            message = str(error)
            debug_log(f"cy_lookup_error row={excel_row} cert={cert_number} slab={slab_type} error={message}")
        with self.lock:
            if generation != self.cy_lookup_generation or excel_row not in self.cy_lookup_inflight:
                debug_log(f"cy_lookup_cancelled_after_lookup row={excel_row} cert={cert_number}")
                return
            self.cy_lookup_inflight.discard(excel_row)
            row = next((candidate for candidate in self.rows if candidate.excel_row == excel_row), None)
            if row is not None and str(row.cert_number or "").strip() == cert_number:
                self.updated_row_ids.add(id(row))
                existing_status = str(row.status or "").strip()
                is_cy_only_status = existing_status in {"CY queued", "CY unavailable", "CY OK"}
                if value is not None:
                    row.cy_value = value
                    row.cy_confidence = confidence
                    if is_cy_only_status:
                        row.status = "CY OK"
                    row.notes = append_note(row.notes, f"CY value: ${value:,.2f}")
                    debug_log(f"cy_lookup_ok row={excel_row} cert={cert_number} value={value} confidence={confidence}")
                elif message:
                    if is_cy_only_status:
                        row.status = "CY unavailable"
                    row.notes = append_note(row.notes, f"CY lookup: {message}")
                    debug_log(f"cy_lookup_unavailable row={excel_row} cert={cert_number} message={message}")
            should_close = self.cy_batch_running and not self.cy_lookup_inflight and not self.cy_lookup_pending and not self.cardladder_running
            if should_close:
                self.cy_batch_running = False
        if should_close:
            close_cy_adapter()
        if self.on_update:
            self.on_update()

    def _apply_cardladder_result_to_row(self, row: WorkbookRow, result: dict) -> None:
        result_status = str(result.get("status") or "")
        value = parse_value(result.get("value"))
        row.card_ladder_value = value
        ocr = result.get("ocr") if isinstance(result.get("ocr"), dict) else {}
        comps = ocr.get("comps") if isinstance(ocr.get("comps"), list) else []
        profile_title = clean_profile_title(ocr.get("profileTitle") or ocr.get("profile_title") or ocr.get("profile"))
        profile_grader = clean_grader(ocr.get("profileGrader") or ocr.get("profile_grader") or row.grader)
        profile_grade = clean_grade(ocr.get("profileGrade") or ocr.get("profile_grade") or "")
        raw_comp_count = len(comps)
        comps = filter_comps_for_card(comps, profile_title or row.card_title)
        filtered_comp_count = raw_comp_count - len(comps)
        if result_status == "partial_comp_capture":
            if profile_title:
                row.card_title = build_card_title(profile_title, profile_grader, profile_grade)
                fill_missing_category_from_title(row)
            if row_has_comp_data(row):
                row.notes = str(result.get("error") or "Partial Card Ladder capture skipped; kept existing comps.")
                return
            row.card_ladder_value = None
            row.card_ladder_comps_average = None
            row.card_ladder_comps = ""
            row.card_ladder_screenshot = ""
            row.status = "Card Ladder partial capture"
            row.notes = str(result.get("error") or "Card Ladder comp capture was incomplete.")
            return
        if result_status == "invalid_cert":
            row.card_title = ""
            row.card_ladder_value = None
            row.card_ladder_comps_average = None
            row.card_ladder_comps = ""
            row.card_ladder_screenshot = ""
            row.status = "Card Ladder invalid cert"
            row.notes = str(result.get("error") or "Card Ladder showed no information with this cert.")
            return
        if (
            result_status == "no_results"
            and not result.get("extensionVersion")
            and not profile_title
            and not ocr
        ):
            row.status = "Reload Card Ladder extension"
            row.notes = (
                "The Card Ladder result came from an older Chrome extension that cannot capture "
                "profile names on no-result pages. Reload the bundled Card Ladder Auto-Comp extension."
            )
            return
        if result_status == "extension_error":
            row.card_ladder_value = None
            row.card_ladder_comps_average = None
            row.card_ladder_comps = ""
            row.card_ladder_screenshot = str(ocr.get("debugImage") or "")
            row.status = "Card Ladder extension error"
            row.notes = str(result.get("error") or "Card Ladder lookup failed before a result could be captured.")
            return
        if profile_title:
            row.card_title = build_card_title(profile_title, profile_grader, profile_grade)
            fill_missing_category_from_title(row)
        row.card_ladder_comps_average = comp_price(comps, self.comp_strategy)
        row.card_ladder_comps = format_comps(comps, self.comp_strategy)
        row.card_ladder_screenshot = str(ocr.get("debugImage") or "")
        if result_status == "no_results":
            row.status = "Card Ladder no results"
        elif raw_comp_count and not comps:
            row.status = "Card Ladder review"
        else:
            row.status = "Card Ladder OK" if value is not None else "Card Ladder review"
        filter_note = f"Rejected {filtered_comp_count} likely wrong-card comp(s)." if filtered_comp_count else ""
        row.notes = " ".join(part for part in (str(result.get("error") or result.get("status") or ""), filter_note) if part).strip()

    def finish_cardladder(self, payload: dict) -> None:
        cy_lookups: list[tuple[int, str, str, int]] = []
        with self.lock:
            self.cardladder_running = False
            self.cancel_requested = False
            cy_generation = self.cy_lookup_generation
            for row in self.rows:
                if row.status == "Queued":
                    row.status = "Card Ladder not found"
            pending_rows = [
                row
                for row in self.rows
                if row.excel_row in self.cy_lookup_pending
            ]
            self.cy_lookup_pending.clear()
            for row in pending_rows:
                candidate = self._cy_lookup_candidate(row, force=True)
                if candidate is not None:
                    self.cy_lookup_inflight.add(row.excel_row)
                    cy_lookups.append((*candidate, cy_generation))
            if cy_lookups:
                self.cy_batch_running = True
            self.cardladder_allows_cy = False
            debug_log(f"finish_cardladder pending_cy={len(cy_lookups)}")
        for cy_lookup in cy_lookups:
            threading.Thread(target=self._cy_lookup_worker, args=cy_lookup, daemon=True).start()
        if self.on_update:
            self.on_update()

    def snapshot(self) -> dict:
        with self.lock:
            return {
                "bridgeVersion": BRIDGE_VERSION,
                "instanceId": self.instance_id,
                "extensionLastSeen": self.last_seen_extension,
                "extensionVersion": self.extension_version,
                "extensionManifestVersion": self.extension_manifest_version,
                "extensionName": self.extension_name,
                "extensionUrl": self.extension_url,
                "expectedExtensionVersion": EXPECTED_CARDLADDER_EXTENSION_VERSION,
                "expectedManifestVersion": EXPECTED_CARDLADDER_MANIFEST_VERSION,
                "lastResultExtensionVersion": self.last_result_extension_version,
                "cardladderRunning": self.cardladder_running,
                "cancelRequested": self.cancel_requested,
                "compStrategy": self.comp_strategy,
                "keepNoteSources": list(self.keep_note_sources),
                "lastKeepSync": dict(self.last_keep_sync),
                "rows": [asdict(row) for row in self.rows],
            }


def parse_value(value) -> float | None:
    if value is None or value == "":
        return None
    text = str(value).replace("$", "").replace(",", "").strip()
    multiplier = 1
    if text.lower().endswith("k"):
        multiplier = 1000
        text = text[:-1].strip()
    if re.fullmatch(r"-?\d{1,3}\.\d{3}", text):
        text = text.replace(".", "")
    try:
        return float(text) * multiplier
    except ValueError:
        return None


def keep_urls_match(first: str, second: str) -> bool:
    first_key = keep_url_key(first)
    second_key = keep_url_key(second)
    if first_key and second_key:
        return first_key == second_key
    first_norm = normalize_keep_url(first)
    second_norm = normalize_keep_url(second)
    return bool(first_norm and second_norm and (first_norm == second_norm or first_norm.startswith(second_norm) or second_norm.startswith(first_norm)))


def keep_url_key(value: str) -> str:
    raw = str(value or "")
    parsed = urlparse(raw)
    haystack = f"{parsed.path}#{parsed.fragment}"
    match = re.search(r"/notes/([^/?#]+)", haystack)
    if match:
        return unquote(match.group(1)).strip().lower()
    match = re.search(r"#NOTE/([^/?#]+)", haystack, flags=re.I)
    if match:
        return unquote(match.group(1)).strip().lower()
    match = re.search(r"(?:note|id|text)%3D([^&#]+)", haystack, flags=re.I)
    return unquote(match.group(1)).strip().lower() if match else ""


def normalize_keep_url(value: str) -> str:
    parsed = urlparse(str(value or "").strip())
    if not parsed.netloc.lower().endswith("keep.google.com"):
        return ""
    return f"{parsed.scheme}://{parsed.netloc}{parsed.path}#{parsed.fragment}".rstrip("/#").lower()


def cy_lookup_enabled(platform: str | None = None) -> bool:
    if os.environ.get("LUCAS_DISABLE_CY_LOOKUP", "").strip().lower() in {"1", "true", "yes"}:
        return False
    return (platform or sys.platform) == "darwin"


def lookup_cy_buy_price(cert_number: str, slab_type: str) -> tuple[float | None, object | None, str]:
    cert_number = str(cert_number or "").strip()
    slab_type = clean_grader(slab_type)
    if not cert_number:
        return None, None, "missing cert number"
    if slab_type not in {"PSA", "BGS", "CGC", "SGC"}:
        return None, None, f"unsupported slab type {slab_type or 'unknown'}"
    # CourtYard is a single macOS GUI, so only one automation sequence can safely
    # click/type/read at a time even when LUCAS has several CY worker threads.
    with _CY_LOOKUP_LOCK:
        debug_log(f"cy_lookup_gui_start cert={cert_number} slab={slab_type}")
        payload = get_cy_adapter().submit_cert_lookup(cert_number, slab_type)
        debug_log(f"cy_lookup_gui_done cert={cert_number} slab={slab_type}")
    value = parse_value(payload.get("cy_buy_price"))
    confidence = payload.get("cy_confidence")
    if value is not None:
        return value, confidence, str(payload.get("message") or "CY lookup OK")
    status = str(payload.get("status") or "").strip()
    message = str(payload.get("message") or payload.get("detail") or "CY value unavailable").strip()
    return None, confidence, f"{status}: {message}" if status else message


def get_cy_adapter() -> CYMacOSAdapter:
    global _CY_ADAPTER
    with _CY_ADAPTER_LOCK:
        if _CY_ADAPTER is None:
            _CY_ADAPTER = CYMacOSAdapter()
        return _CY_ADAPTER


def close_cy_adapter() -> None:
    with _CY_ADAPTER_LOCK:
        adapter = _CY_ADAPTER
    if adapter is None:
        return
    try:
        adapter.close_app()
        debug_log("cy_app_closed")
    except Exception as error:
        debug_log(f"cy_app_close_error error={error}")


def append_note(existing: str, note: str) -> str:
    existing = str(existing or "").strip()
    note = str(note or "").strip()
    if not note or note in existing:
        return existing
    return f"{existing}; {note}" if existing else note


def debug_log(message: str) -> None:
    try:
        DEBUG_DIR.mkdir(parents=True, exist_ok=True)
        stamp = time.strftime("%Y-%m-%d %H:%M:%S")
        with DEBUG_LOG.open("a", encoding="utf-8") as handle:
            handle.write(f"[{stamp}] {message}\n")
    except Exception:
        pass


def is_blank_card_title(card_title: str, grader: str) -> bool:
    title = str(card_title or "").strip()
    company = str(grader or "").strip()
    if not title:
        return True
    return bool(company and title.upper() == company.upper())


def row_has_comp_data(row: WorkbookRow) -> bool:
    status = str(row.status or "").strip().lower().replace("_", " ")
    notes = str(row.notes or "").strip().lower().replace("_", " ")
    has_terminal_empty_result = any(
        token in f"{status} {notes}"
        for token in (
            "invalid cert",
            "no information with this cert",
            "no results",
        )
    )
    return (
        row.card_ladder_comps_average is not None
        or bool(str(row.card_ladder_comps or "").strip())
        or has_terminal_empty_result
    )


def clean_profile_title(value) -> str:
    text = re.sub(r"\s+", " ", str(value or "")).strip()
    text = re.sub(r"^profile\s*:\s*", "", text, flags=re.I)
    tail_patterns = [
        r"\s+\bclose\s+\$?\d[\d,]*(?:\.\d{1,2})?.*$",
        r"\s+\bclose\s+search[_\s-]*off\b.*$",
        r"\s+\bclose\b\s*$",
        r"\s+\bclose\b\s+(?=\b(?:PSA|BGS|SGC|CGC|BECKETT|BVG)\b|\d+(?:\.\d+)?\b).*$",
        r"\s+[x×]\s*$",
        r"\s+\bthere\s+are\s+no\s+results\b.*$",
        r"\s+\btry\s+searching\b.*$",
        r"\s+\bhelp[_\s-]*outline\b.*$",
        r"\s+\b(?:date\s+sold|type|price)\b.*$",
        r"\s+\$\d[\d,]*(?:\.\d{1,2})?\s+\b(?:help[_\s-]*outline|ebay|fanatics|pwcc|goldin|alt|myslabs|heritage|pristine|auction)\b.*$",
    ]
    for pattern in tail_patterns:
        text = re.sub(pattern, "", text, flags=re.I)
    text = re.sub(r"\s*\(pop\s*[^)]*\)\s*$", "", text, flags=re.I)
    return re.sub(r"\s+", " ", text).strip()


def clean_grader(value) -> str:
    text = re.sub(r"\s+", " ", str(value or "")).strip().upper()
    aliases = {"BECKETT": "BGS", "BVG": "BGS", "PSA": "PSA", "BGS": "BGS", "SGC": "SGC", "CGC": "CGC"}
    return aliases.get(text, text)


def clean_grade(value) -> str:
    text = re.sub(r"\s+", " ", str(value or "")).strip()
    matches = re.findall(r"\d+(?:\.\d+)?", text)
    return matches[-1] if matches else ""


def build_card_title(description: str, grader: str, grade: str) -> str:
    title = clean_profile_title(description)
    parts = [title] if title else []
    if grader and not re.search(rf"\b{re.escape(grader)}\b", title, re.I):
        parts.append(grader)
    if grade and not re.search(rf"(?<!\d){re.escape(grade)}(?!\d)", " ".join(parts)):
        parts.append(grade)
    return re.sub(r"\s+", " ", " ".join(parts)).strip()


def average_comp_prices(comps: list[dict]) -> float | None:
    comps = dedupe_comps(comps)
    values = [parse_value(comp.get("price")) for comp in comps if isinstance(comp, dict)]
    values = [value for value in values if value is not None]
    if not values:
        return None
    return round(sum(values) / len(values), 2)


def filter_comps_for_card(comps: list[dict], card_title: str) -> list[dict]:
    if not comps:
        return []
    target_tokens = comp_match_tokens(card_title)
    target_years = {token for token in target_tokens if re.fullmatch(r"(?:19|20)\d{2}", token)}
    target_numbers = comp_card_number_tokens(target_tokens)
    target_parallel_tokens = target_tokens & {
        "gold",
        "silver",
        "orange",
        "red",
        "blue",
        "green",
        "yellow",
        "black",
        "white",
        "purple",
        "pink",
        "holo",
        "laser",
        "shimmer",
        "fluorescent",
        "optic",
        "mosaic",
        "prizm",
        "refractor",
        "wave",
        "cracked",
        "ice",
        "choice",
    }
    if len(target_tokens) < 4:
        return comps
    filtered: list[dict] = []
    for comp in comps:
        if not isinstance(comp, dict):
            continue
        title = clean_comp_title(comp.get("title"))
        tokens = comp_match_tokens(title)
        if len(tokens) < 3:
            filtered.append(comp)
            continue
        comp_years = {token for token in tokens if re.fullmatch(r"(?:19|20)\d{2}", token)}
        if target_years and comp_years and target_years.isdisjoint(comp_years):
            continue
        comp_numbers = comp_card_number_tokens(tokens)
        if target_numbers and comp_numbers and target_numbers.isdisjoint(comp_numbers):
            continue
        comp_parallel_tokens = tokens & target_parallel_tokens
        if len(target_parallel_tokens) >= 2 and not comp_parallel_tokens:
            continue
        overlap = target_tokens & tokens
        required = 2 if len(target_tokens) < 7 else 3
        if len(overlap) >= required:
            filtered.append(comp)
            continue
        ratio = len(overlap) / max(1, min(len(target_tokens), len(tokens)))
        if ratio >= 0.42:
            filtered.append(comp)
    return filtered


def comp_match_tokens(value: object) -> set[str]:
    text = clean_comp_title(value).lower()
    text = re.sub(r"[^a-z0-9]+", " ", text)
    raw_tokens = re.findall(r"[a-z0-9]+", text)
    stop = {
        "psa",
        "bgs",
        "sgc",
        "cgc",
        "gem",
        "mint",
        "rookie",
        "rc",
        "card",
        "cards",
        "auto",
        "autograph",
        "autographs",
        "refractor",
        "refractors",
        "chrome",
        "panini",
        "topps",
        "donruss",
        "bowman",
        "upper",
        "deck",
    }
    return {token for token in raw_tokens if len(token) >= 3 and token not in stop}


def comp_card_number_tokens(tokens: set[str]) -> set[str]:
    return {token for token in tokens if token.isdigit() and len(token) >= 3 and not re.fullmatch(r"(?:19|20)\d{2}", token)}


def dedupe_comps(comps: list[dict]) -> list[dict]:
    cleaned: list[dict] = []
    for comp in comps:
        if not isinstance(comp, dict):
            continue
        normalized = dict(comp)
        normalized["title"] = clean_comp_title(normalized.get("title"))
        normalized["source"] = clean_comp_source(normalized.get("source"))
        if is_junk_comp_title(normalized.get("title")):
            continue
        cleaned.append(normalized)

    best_by_key: dict[tuple[str, str, str, str], dict] = {}
    order: list[tuple[str, str, str, str]] = []
    for comp in cleaned:
        date = normalized_comp_date_key(comp.get("date_sold"))
        price = str(comp.get("price") or "").replace("$", "").replace(",", "").strip()
        sale_type = re.sub(r"\s+", " ", str(comp.get("sale_type") or "")).strip().lower()
        key_base = (date, price, sale_type)
        title_key = compact_comp_title(comp.get("title"))[:80]
        target_key = None
        for existing_key, existing in best_by_key.items():
            same_price = existing_key[1] == price
            same_sale_type = existing_key[2] == sale_type
            same_date_price_type = existing_key[:3] == key_base
            same_date = existing_key[0] == key_base[0]
            existing_title_key = compact_comp_title(existing.get("title"))[:80]
            same_source = clean_comp_source(existing.get("source")).lower() == clean_comp_source(comp.get("source")).lower()
            similar_title = bool(title_key and existing_title_key and (title_key in existing_title_key or existing_title_key in title_key))
            if same_date and same_source and similar_title:
                target_key = existing_key
                break
            if not same_price:
                continue
            if same_date_price_type and (same_source or similar_title):
                target_key = existing_key
                break
            if same_sale_type and same_source and similar_title:
                target_key = existing_key
                break
        target_key = target_key or (*key_base, title_key or str(len(order)))
        if target_key not in best_by_key:
            order.append(target_key)
            best_by_key[target_key] = comp
            continue
        existing = best_by_key[target_key]
        existing_date = parse_comp_date(existing.get("date_sold"))
        comp_date = parse_comp_date(comp.get("date_sold"))
        if existing_date and comp_date and existing_date != comp_date:
            if comp_date < existing_date:
                best_by_key[target_key] = comp
            continue
        if comp_quality(comp) > comp_quality(existing):
            best_by_key[target_key] = comp
    return [best_by_key[key] for key in order][:5]


def clean_comp_source(value) -> str:
    text = re.sub(r"\s+", " ", str(value or "")).strip()
    text = re.sub(r"\s*\(confirmed paid\)\s*", "", text, flags=re.I)
    return text


def clean_comp_title(value) -> str:
    text = re.sub(r"\s+", " ", str(value or "")).strip()
    text = re.sub(r"\b(?:close|help[_\s-]*outline|Date Sold|Type|Price)\b", " ", text, flags=re.I)
    text = re.sub(r"^\s*[-|:]+\s*", "", text)
    text = re.sub(r"\s*[-|:]+\s*$", "", text)
    return re.sub(r"\s+", " ", text).strip()


def compact_comp_title(value) -> str:
    text = clean_comp_title(value).lower()
    text = re.sub(r"\$\s*[\d,]+(?:\.\d{1,2})?", " ", text)
    text = re.sub(r"\b(psa|bgs|sgc|cgc|gem|mint|mt|pop|rookie|rc)\b", " ", text)
    text = re.sub(r"[^a-z0-9]+", "", text)
    return text


def is_junk_comp_title(value) -> bool:
    text = clean_comp_title(value)
    if not text:
        return True
    alnum = re.sub(r"[^A-Za-z0-9]", "", text)
    if len(alnum) < 8:
        return True
    return not re.search(r"[A-Za-z]{3,}", text)


def comp_quality(comp: dict) -> int:
    title = clean_comp_title(comp.get("title"))
    score = min(len(title), 160)
    if re.search(r"\b\d{4}\b", title):
        score += 20
    if re.search(r"#\s*[A-Za-z0-9-]+|\b[A-Za-z]{1,5}\d{1,4}\b", title):
        score += 10
    if comp_price_conflicts_with_title(comp):
        score -= 120
    elif re.search(r"\$\s*[\d,]+(?:\.\d{1,2})?", title):
        score -= 40
    if is_junk_comp_title(title):
        score -= 200
    return score


def comp_price_conflicts_with_title(comp: dict) -> bool:
    price = parse_value(comp.get("price"))
    if price is None:
        return False
    title_values = [
        parse_value(match.group(0))
        for match in re.finditer(r"\$\s*[\d,]+(?:\.\d{1,2})?", clean_comp_title(comp.get("title")))
    ]
    title_values = [value for value in title_values if value is not None]
    return bool(title_values and all(abs(value - price) > 0.01 for value in title_values))


def comp_price(comps: list[dict], strategy: str) -> float | None:
    comps = dedupe_comps(comps)
    values = comp_values(comps)
    if not values:
        return None
    if strategy == COMP_STRATEGY_HIGH:
        return max(values)
    if strategy == COMP_STRATEGY_LOW:
        return min(values)
    if strategy == COMP_STRATEGY_STALE_NEWEST:
        return stale_newest_else_average(comps, values)
    return round(sum(values) / len(values), 2)


def comp_values(comps: list[dict]) -> list[float]:
    comps = dedupe_comps(comps)
    values = [parse_value(comp.get("price")) for comp in comps[:5] if isinstance(comp, dict)]
    return [value for value in values if value is not None]


def newest_comp_date(comps: list[dict]) -> datetime | None:
    comps = dedupe_comps(comps)
    dates = []
    for comp in comps[:5]:
        if not isinstance(comp, dict):
            continue
        parsed = parse_comp_date(comp.get("date_sold"))
        if parsed:
            dates.append(parsed)
    return max(dates) if dates else None


def stale_newest_else_average(comps: list[dict], values: list[float]) -> float | None:
    comps = dedupe_comps(comps)
    values = comp_values(comps)
    if not values:
        return None
    dated_values: list[tuple[datetime, float]] = []
    for comp in comps[:5]:
        if not isinstance(comp, dict):
            continue
        value = parse_value(comp.get("price"))
        if value is None:
            continue
        sold_date = parse_comp_date(comp.get("date_sold"))
        if sold_date:
            dated_values.append((sold_date, value))
    dated_values.sort(key=lambda item: item[0], reverse=True)
    if dated_values and (datetime.now() - dated_values[0][0]).days > 7:
        return best_value_for_comp_date(comps, dated_values[0][0])
    if len(dated_values) >= 2 and (dated_values[0][0] - dated_values[1][0]).days > 7:
        return best_value_for_comp_date(comps, dated_values[0][0])
    average_values = [value for _sold_date, value in dated_values] if dated_values else values
    return round(sum(average_values) / len(average_values), 2)


def best_value_for_comp_date(comps: list[dict], sold_date: datetime) -> float | None:
    same_day: list[dict] = []
    for comp in comps[:5]:
        if not isinstance(comp, dict):
            continue
        value = parse_value(comp.get("price"))
        comp_date = parse_comp_date(comp.get("date_sold"))
        if value is not None and comp_date == sold_date:
            same_day.append(comp)
    if not same_day:
        return None
    best = max(same_day, key=comp_quality)
    return parse_value(best.get("price"))


def normalized_comp_date_key(value) -> str:
    parsed = parse_comp_date(value)
    if parsed:
        return parsed.strftime("%Y-%m-%d")
    return re.sub(r"\s+", " ", str(value or "")).strip().lower()


def parse_comp_date(value) -> datetime | None:
    text = re.sub(r"\s+", " ", str(value or "")).strip()
    text = re.sub(r"(\d+)(st|nd|rd|th)", r"\1", text, flags=re.I)
    text = re.sub(r"\bSept\.?\b", "Sep", text, flags=re.I)
    text = re.sub(r"\b([A-Za-z]{3,9})\.", r"\1", text)
    date_match = re.search(
        r"\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.?\s+\d{1,2},?\s+\d{4}\b"
        r"|\b\d{1,2}/\d{1,2}/\d{2,4}\b"
        r"|\b\d{4}-\d{1,2}-\d{1,2}\b",
        text,
        flags=re.I,
    )
    if date_match:
        text = date_match.group(0)
    text = re.sub(r"\bSept\.?\b", "Sep", text, flags=re.I)
    text = re.sub(r"\b([A-Za-z]{3,9})\.", r"\1", text)
    for fmt in ("%b %d, %Y", "%B %d, %Y", "%m/%d/%Y", "%m/%d/%y", "%Y-%m-%d"):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            pass
    return None


def format_comps(comps: list[dict], strategy: str = COMP_STRATEGY_AVERAGE) -> str:
    comps = dedupe_comps(comps)
    lines: list[str] = []
    selected_value = comp_price(comps, strategy)
    label = COMP_STRATEGY_LABELS.get(strategy, COMP_STRATEGY_LABELS[COMP_STRATEGY_AVERAGE])
    if selected_value is not None:
        lines.append(f"Comp method: {label} -> ${selected_value:,.2f}")
    for index, comp in enumerate(comps[:5], start=1):
        if not isinstance(comp, dict):
            continue
        date = str(comp.get("date_sold") or "").strip()
        price = str(comp.get("price") or "").strip()
        sale_type = str(comp.get("sale_type") or "").strip()
        source = str(comp.get("source") or "").strip()
        title = str(comp.get("title") or "").strip()
        lines.append(f"{index}. {date} | {price} | {sale_type} | {source} | {title}".strip())
    return "\n".join(lines)


def parse_formatted_comps(text: str) -> list[dict]:
    comps: list[dict] = []
    for line in str(text or "").splitlines():
        if re.match(r"^\s*comp method\s*:", line, flags=re.I):
            continue
        match = re.match(r"^\s*\d+\.\s*(.*)$", line)
        if not match:
            continue
        parts = [part.strip() for part in match.group(1).split("|")]
        if len(parts) < 2:
            continue
        comps.append(
            {
                "date_sold": parts[0],
                "price": parts[1],
                "sale_type": parts[2] if len(parts) > 2 else "",
                "source": parts[3] if len(parts) > 3 else "",
                "title": " | ".join(parts[4:]) if len(parts) > 4 else "",
            }
        )
    return comps


class BridgeServer:
    def __init__(self, state: BridgeState, host: str = "0.0.0.0", port: int = 8765, allow_port_fallback: bool | None = None) -> None:
        self.state = state
        self.host = host
        self.port = port
        self.allow_port_fallback = bool(allow_port_fallback) if allow_port_fallback is not None else os.environ.get("LUCAS_ALLOW_BRIDGE_PORT_FALLBACK", "").strip().lower() in {"1", "true", "yes"}
        self.httpd: ThreadingHTTPServer | None = None
        self.thread: threading.Thread | None = None
        self.started = False
        self.error = ""

    def start(self) -> None:
        state = self.state

        class Handler(BaseHTTPRequestHandler):
            def do_OPTIONS(self):
                if not self._origin_allowed():
                    self._send_json({"ok": False, "error": "origin not allowed"}, status=403)
                    return
                self._send_json({})

            def do_GET(self):
                parsed = urlparse(self.path)
                if not self._request_allowed(parsed.path):
                    self._send_json({"ok": False, "error": "local bridge access only"}, status=403)
                    return
                media_match = re.match(r"^/instagram/media/([^/]+)/([^/]+)(?:/[^/]*)?$", parsed.path)
                if media_match:
                    media = state.get_instagram_media(media_match.group(1), media_match.group(2))
                    if media is None:
                        self._send_json({"ok": False, "error": "not found"}, status=404)
                        return
                    body, content_type = media
                    self._send_bytes(body, content_type, cache_control="public, max-age=3600")
                    return
                mobile_photo_match = re.match(r"^/mobile/api/inventory/photo/([^/]+)(?:/[^/]*)?$", parsed.path)
                if mobile_photo_match:
                    query = parse_qs(parsed.query)
                    media = state.get_mobile_inventory_photo(None, query, mobile_photo_match.group(1))
                    if media is None:
                        self._send_json({"ok": False, "error": "not found"}, status=404)
                        return
                    body, content_type = media
                    self._send_bytes(body, content_type, cache_control="private, max-age=300")
                    return
                mobile_profile = self._mobile_profile(parsed.path)
                if mobile_profile:
                    profile_prefix = f"/mobile/{mobile_profile}"
                    if parsed.path in {profile_prefix, f"{profile_prefix}/"}:
                        self._send_mobile_index(mobile_profile)
                        return
                    relative = parsed.path.removeprefix(f"{profile_prefix}/") or "index.html"
                    if relative == "index.html":
                        self._send_mobile_index(mobile_profile)
                        return
                    if relative == "manifest.webmanifest":
                        self._send_mobile_manifest(mobile_profile)
                        return
                    if relative.startswith("api/"):
                        if relative == "api/config":
                            self._send_json(state.mobile_config())
                            return
                        self._send_json({"ok": False, "error": "unknown endpoint"}, status=404)
                        return
                    self._send_static(MOBILE_APP_DIR / relative)
                    return
                if parsed.path in {"/mobile", "/mobile/"}:
                    self._send_static(MOBILE_APP_DIR / "index.html")
                    return
                if parsed.path.startswith("/mobile/"):
                    relative = parsed.path.removeprefix("/mobile/") or "index.html"
                    if relative.startswith("api/"):
                        if relative == "api/config":
                            self._send_json(state.mobile_config())
                            return
                        self._send_json({"ok": False, "error": "unknown endpoint"}, status=404)
                        return
                    self._send_static(MOBILE_APP_DIR / relative)
                    return
                if self.path.startswith("/command"):
                    query = parse_qs(parsed.query)
                    metadata = {
                        "extensionVersion": query.get("extensionVersion", [""])[0],
                        "manifestVersion": query.get("manifestVersion", [""])[0],
                        "extensionName": query.get("extensionName", [""])[0],
                        "extensionUrl": query.get("extensionUrl", [""])[0],
                    }
                    self._send_json(state.extension_poll(metadata))
                    return
                if self.path.startswith("/status"):
                    self._send_json(state.snapshot())
                    return
                self._send_json({"ok": True, "service": "comp-orchestrator"})

            def do_POST(self):
                parsed = urlparse(self.path)
                if not self._request_allowed(parsed.path):
                    self._send_json({"ok": False, "error": "local bridge access only"}, status=403)
                    return
                payload = self._read_json()
                mobile_api_path = self._mobile_api_path(parsed.path)
                if mobile_api_path.startswith("/mobile/api/inventory/search"):
                    self._send_json(state.search_mobile_inventory(payload))
                    return
                if mobile_api_path.startswith("/mobile/api/inventory/add"):
                    self._send_json(state.add_mobile_inventory(payload))
                    return
                if mobile_api_path.startswith("/mobile/api/inventory/sold"):
                    self._send_json(state.mark_mobile_inventory_sold(payload))
                    return
                if mobile_api_path.startswith("/mobile/api/card/identify"):
                    self._send_json(state.identify_mobile_card(payload))
                    return
                if mobile_api_path.startswith("/mobile/api/profit/summary"):
                    self._send_json(state.get_mobile_profit_summary(payload))
                    return
                if mobile_api_path.startswith("/mobile/api/expenses/add"):
                    self._send_json(state.add_mobile_expense(payload))
                    return
                if mobile_api_path.startswith("/mobile/api/payouts"):
                    self._send_json(state.get_mobile_payouts(payload))
                    return
                if mobile_api_path.startswith("/mobile/api/sync/queue"):
                    self._send_json(state.sync_mobile_queue(payload))
                    return
                if parsed.path.startswith("/ack"):
                    state.acknowledge_command(int(payload.get("id") or 0))
                    self._send_json({"ok": True})
                    return
                if parsed.path.startswith("/result/cardladder"):
                    state.post_cardladder_result(payload)
                    self._send_json({"ok": True})
                    return
                if parsed.path.startswith("/ocr/cardladder"):
                    try:
                        self._send_json(extract_cl_value_from_data_url(str(payload.get("image") or "")))
                    except Exception as error:
                        self._send_json({"ok": False, "value": None, "error": str(error)})
                    return
                if parsed.path.startswith("/finish/cardladder"):
                    state.finish_cardladder(payload)
                    self._send_json({"ok": True})
                    return
                if parsed.path.startswith("/source/google-keep"):
                    self._send_json(state.post_google_keep_note(payload))
                    return
                self._send_json({"ok": False, "error": "unknown endpoint"}, status=404)

            def _read_json(self) -> dict:
                length = int(self.headers.get("content-length") or 0)
                raw = self.rfile.read(length).decode("utf-8") if length else "{}"
                return json.loads(raw or "{}")

            def _send_json(self, payload: dict, status: int = 200) -> None:
                body = json.dumps(payload).encode("utf-8")
                self.send_response(status)
                self.send_header("content-type", "application/json")
                origin = self.headers.get("origin", "")
                if origin and self._origin_allowed():
                    self.send_header("access-control-allow-origin", origin)
                    self.send_header("vary", "Origin")
                self.send_header("access-control-allow-methods", "GET,POST,OPTIONS")
                self.send_header("access-control-allow-headers", "content-type")
                self.send_header("content-length", str(len(body)))
                self.end_headers()
                self.wfile.write(body)

            def _origin_allowed(self) -> bool:
                return request_origin_allowed(self.headers.get("origin", ""), self.headers.get("host", ""))

            def _client_is_loopback(self) -> bool:
                return is_loopback_address(str(self.client_address[0] if self.client_address else ""))

            def _request_allowed(self, path: str) -> bool:
                if not self._origin_allowed():
                    return False
                if any(path.startswith(prefix) for prefix in BRIDGE_LOCAL_ONLY_PATH_PREFIXES):
                    return self._client_is_loopback()
                return True

            def _mobile_profile(self, path: str) -> str:
                match = re.match(r"^/mobile/(team|personal)(?:/|$)", path)
                return match.group(1) if match else ""

            def _mobile_api_path(self, path: str) -> str:
                match = re.match(r"^/mobile/(?:team|personal)/(api/.*)$", path)
                if match:
                    return f"/mobile/{match.group(1)}"
                return path

            def _send_mobile_index(self, profile: str) -> None:
                label = MOBILE_PROFILE_LABELS.get(profile, "LUCAS")
                base = f"/mobile/{profile}"
                try:
                    html = (MOBILE_APP_DIR / "index.html").read_text(encoding="utf-8")
                except OSError:
                    self._send_json({"ok": False, "error": "not found"}, status=404)
                    return
                html = html.replace('content="LUCAS"', f'content="{label}"')
                html = html.replace("<title>LUCAS Mobile</title>", f"<title>{label} Mobile</title>")
                html = html.replace('href="/mobile/manifest.webmanifest"', f'href="{base}/manifest.webmanifest"')
                html = html.replace('href="/mobile/styles.css"', f'href="{base}/styles.css"')
                html = html.replace('src="/mobile/app.js"', f'src="{base}/app.js"')
                self._send_bytes(html.encode("utf-8"), "text/html; charset=utf-8")

            def _send_mobile_manifest(self, profile: str) -> None:
                label = MOBILE_PROFILE_LABELS.get(profile, "LUCAS")
                payload = {
                    "name": f"{label} Inventory",
                    "short_name": label,
                    "start_url": f"/mobile/{profile}",
                    "scope": f"/mobile/{profile}/",
                    "display": "standalone",
                    "background_color": "#101820",
                    "theme_color": "#101820",
                    "description": f"{label} inventory companion.",
                    "icons": [],
                }
                self._send_json(payload)

            def _send_bytes(self, body: bytes, content_type: str, cache_control: str = "no-cache") -> None:
                self.send_response(200)
                self.send_header("content-type", content_type)
                self.send_header("cache-control", cache_control)
                self.send_header("content-length", str(len(body)))
                self.end_headers()
                self.wfile.write(body)

            def _send_static(self, path: Path) -> None:
                try:
                    root = MOBILE_APP_DIR.resolve()
                    resolved = path.resolve()
                    if root not in (resolved, *resolved.parents) or not resolved.is_file():
                        self._send_json({"ok": False, "error": "not found"}, status=404)
                        return
                    body = resolved.read_bytes()
                except OSError:
                    self._send_json({"ok": False, "error": "not found"}, status=404)
                    return
                content_type = mimetypes.guess_type(str(resolved))[0] or "application/octet-stream"
                self._send_bytes(body, content_type)

            def log_message(self, format, *args):
                return

        class ReusableThreadingHTTPServer(ThreadingHTTPServer):
            allow_reuse_address = True

        last_error = ""
        port_count = 8 if self.allow_port_fallback else 1
        for candidate_port in range(self.port, self.port + port_count):
            if self._port_has_listener(candidate_port):
                last_error = f"{self.host}:{candidate_port} already has a listener"
                continue
            try:
                self.httpd = ReusableThreadingHTTPServer((self.host, candidate_port), Handler)
                self.port = candidate_port
                self.error = ""
                break
            except OSError as error:
                last_error = str(error)
                self.httpd = None
        if self.httpd is None:
            self.started = False
            self.error = last_error or "Could not bind local bridge port."
            return
        self.thread = threading.Thread(target=self.httpd.serve_forever, daemon=True)
        self.thread.start()
        self.started = True

    def _port_has_listener(self, port: int) -> bool:
        host = "127.0.0.1" if self.host in {"0.0.0.0", "::"} else self.host
        try:
            with socket.create_connection((host, port), timeout=0.2):
                return True
        except OSError:
            return False

    def stop(self) -> None:
        if self.httpd:
            self.httpd.shutdown()
            self.httpd.server_close()
        self.started = False
