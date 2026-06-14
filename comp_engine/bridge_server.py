from __future__ import annotations

import json
import re
import socket
import threading
import time
import uuid
from dataclasses import asdict
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Callable
from urllib.parse import parse_qs, urlparse

from cardladder_ocr import extract_cl_value_from_data_url
from workbook_io import WorkbookRow

BRIDGE_VERSION = "2026-06-01-cardladder-result-log-v3"
EXPECTED_CARDLADDER_EXTENSION_VERSION = "2026-06-10-no-results-ocr-fallback-v3"
EXPECTED_CARDLADDER_MANIFEST_VERSION = "0.1.4"
DEBUG_DIR = Path(__file__).resolve().parent.parent / "work" / "cardladder-bridge"
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
        self.on_update: Callable[[], None] | None = None

    def set_rows(self, rows: list[WorkbookRow]) -> None:
        with self.lock:
            self.rows = rows

    def set_comp_strategy(self, strategy: str) -> None:
        with self.lock:
            self.comp_strategy = strategy if strategy in COMP_STRATEGY_LABELS else COMP_STRATEGY_AVERAGE

    def start_all_comps(self, requery_all: bool = False) -> int:
        with self.lock:
            self.command_id += 1
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
            return self.command_id

    def request_cancel(self) -> None:
        with self.lock:
            self.cancel_requested = True
            self.command = None
            self.cardladder_running = False
            for row in self.rows:
                if row.status == "Queued":
                    row.status = "Card Ladder cancelled"
        if self.on_update:
            self.on_update()

    def extension_poll(self, metadata: dict[str, str] | None = None) -> dict:
        with self.lock:
            self.last_seen_extension = time.strftime("%H:%M:%S")
            if metadata:
                self.extension_version = metadata.get("extensionVersion") or self.extension_version
                self.extension_manifest_version = metadata.get("manifestVersion") or self.extension_manifest_version
                self.extension_name = metadata.get("extensionName") or self.extension_name
                self.extension_url = metadata.get("extensionUrl") or self.extension_url
            return {"instanceId": self.instance_id, "command": self.command}

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
        if result_status == "partial_comp_capture":
            if profile_title:
                row.card_title = build_card_title(profile_title, profile_grader, profile_grade)
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
        if profile_title:
            row.card_title = build_card_title(profile_title, profile_grader, profile_grade)
        row.card_ladder_comps_average = comp_price(comps, self.comp_strategy)
        row.card_ladder_comps = format_comps(comps, self.comp_strategy)
        row.card_ladder_screenshot = str(ocr.get("debugImage") or "")
        if result_status == "no_results":
            row.status = "Card Ladder no results"
        else:
            row.status = "Card Ladder OK" if value is not None else "Card Ladder review"
        row.notes = str(result.get("error") or result.get("status") or "")

    def finish_cardladder(self, payload: dict) -> None:
        with self.lock:
            self.cardladder_running = False
            self.cancel_requested = False
            for row in self.rows:
                if row.status == "Queued":
                    row.status = "Card Ladder not found"
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
                "rows": [asdict(row) for row in self.rows],
            }


def parse_value(value) -> float | None:
    if value is None or value == "":
        return None
    try:
        return float(str(value).replace("$", "").replace(",", "").strip())
    except ValueError:
        return None


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
        date = re.sub(r"\s+", " ", str(comp.get("date_sold") or "")).strip().lower()
        price = str(comp.get("price") or "").replace("$", "").replace(",", "").strip()
        sale_type = re.sub(r"\s+", " ", str(comp.get("sale_type") or "")).strip().lower()
        key_base = (date, price, sale_type)
        title_key = compact_comp_title(comp.get("title"))[:80]
        target_key = None
        for existing_key, existing in best_by_key.items():
            same_price = existing_key[1] == price
            same_sale_type = existing_key[2] == sale_type
            same_date_price_type = existing_key[:3] == key_base
            if not same_price:
                continue
            existing_title_key = compact_comp_title(existing.get("title"))[:80]
            same_source = clean_comp_source(existing.get("source")).lower() == clean_comp_source(comp.get("source")).lower()
            similar_title = bool(title_key and existing_title_key and (title_key in existing_title_key or existing_title_key in title_key))
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
    if is_junk_comp_title(title):
        score -= 200
    return score


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
    first_date = parse_comp_date(comps[0].get("date_sold")) if comps and isinstance(comps[0], dict) else None
    second_date = parse_comp_date(comps[1].get("date_sold")) if len(comps) > 1 and isinstance(comps[1], dict) else None
    if first_date and second_date and abs((first_date - second_date).days) > 7:
        newest_value = parse_value(comps[0].get("price"))
        return newest_value if newest_value is not None else round(sum(values) / len(values), 2)
    return round(sum(values) / len(values), 2)


def parse_comp_date(value) -> datetime | None:
    text = re.sub(r"\s+", " ", str(value or "")).strip()
    text = re.sub(r"(\d+)(st|nd|rd|th)", r"\1", text, flags=re.I)
    for fmt in ("%b %d, %Y", "%B %d, %Y", "%m/%d/%Y", "%Y-%m-%d"):
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
    def __init__(self, state: BridgeState, host: str = "127.0.0.1", port: int = 8765) -> None:
        self.state = state
        self.host = host
        self.port = port
        self.httpd: ThreadingHTTPServer | None = None
        self.thread: threading.Thread | None = None
        self.started = False
        self.error = ""

    def start(self) -> None:
        state = self.state

        class Handler(BaseHTTPRequestHandler):
            def do_OPTIONS(self):
                self._send_json({})

            def do_GET(self):
                if self.path.startswith("/command"):
                    parsed = urlparse(self.path)
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
                payload = self._read_json()
                if self.path.startswith("/ack"):
                    state.acknowledge_command(int(payload.get("id") or 0))
                    self._send_json({"ok": True})
                    return
                if self.path.startswith("/result/cardladder"):
                    state.post_cardladder_result(payload)
                    self._send_json({"ok": True})
                    return
                if self.path.startswith("/ocr/cardladder"):
                    try:
                        self._send_json(extract_cl_value_from_data_url(str(payload.get("image") or "")))
                    except Exception as error:
                        self._send_json({"ok": False, "value": None, "error": str(error)})
                    return
                if self.path.startswith("/finish/cardladder"):
                    state.finish_cardladder(payload)
                    self._send_json({"ok": True})
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
                self.send_header("access-control-allow-origin", "*")
                self.send_header("access-control-allow-methods", "GET,POST,OPTIONS")
                self.send_header("access-control-allow-headers", "content-type")
                self.send_header("content-length", str(len(body)))
                self.end_headers()
                self.wfile.write(body)

            def log_message(self, format, *args):
                return

        class ReusableThreadingHTTPServer(ThreadingHTTPServer):
            allow_reuse_address = True

        last_error = ""
        for candidate_port in range(self.port, self.port + 8):
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
        try:
            with socket.create_connection((self.host, port), timeout=0.2):
                return True
        except OSError:
            return False

    def stop(self) -> None:
        if self.httpd:
            self.httpd.shutdown()
            self.httpd.server_close()
        self.started = False
