from __future__ import annotations

import csv
import html
import queue
import base64
import json
import os
import re
import secrets
import shutil
import socket
import sys
import threading
import tkinter as tk
import urllib.parse
import urllib.request
import webbrowser
from collections import defaultdict
from datetime import datetime, timedelta
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from openpyxl import Workbook, load_workbook

ROOT = Path(__file__).resolve().parent
APP_DEBUG_LOG = ROOT / "work" / "lucas-debug.log"
ENGINE_DIR = ROOT / "comp_engine"
if str(ENGINE_DIR) not in sys.path:
    sys.path.insert(0, str(ENGINE_DIR))

from bridge_server import (  # noqa: E402
    COMP_STRATEGY_AVERAGE,
    COMP_STRATEGY_HIGH,
    COMP_STRATEGY_LOW,
    COMP_STRATEGY_STALE_NEWEST,
    EXPECTED_CARDLADDER_EXTENSION_VERSION,
    EXPECTED_CARDLADDER_MANIFEST_VERSION,
    BridgeServer,
    BridgeState,
    comp_price,
    format_comps,
    parse_formatted_comps,
    row_has_comp_data,
)
from workbook_io import WorkbookRow  # noqa: E402
import assignment_engine  # noqa: E402
from assignment_engine import AssignmentEngine  # noqa: E402
from assignment_engine import CONFIG_PATH as ASSIGNMENT_CONFIG_PATH  # noqa: E402
from assignment_engine import gsheet_shortcut_url, is_google_keep_url, keep_note_cache_path, load_gsheet_shortcut, normalize_source_value, path_from_source_value, safe_filename  # noqa: E402
from assignment_config_ui import open_assignment_rules_dialog  # noqa: E402
from google_sheets_import import export_google_sheet_to_xlsx  # noqa: E402
from shared_state import atomic_write_json, local_identity, shared_lock  # noqa: E402

from intake_io import (  # noqa: E402
    append_company_sheet_rows,
    build_card_title,
    clear_received_in_workbooks,
    clean_part,
    default_output_path,
    ensure_company_weekly_sheets,
    format_money,
    infer_grader,
    mark_received_in_workbooks,
    normalize_grader,
    read_company_profit_records,
    read_photo_export,
    read_simple_spreadsheet,
    remove_company_sheet_rows_for_source,
    scan_to_cert,
    summarize_workbook,
    working_sheet_path,
    write_working_sheet,
    workbook_sheet_names,
    write_pipeline_output,
)


PHOTO_APP_ROOT = ROOT / "photo_tool"
PHOTO_APP_DIR = PHOTO_APP_ROOT / "app"
if PHOTO_APP_DIR.exists() and str(PHOTO_APP_DIR) not in sys.path:
    sys.path.insert(0, str(PHOTO_APP_DIR))
try:
    from dotenv import load_dotenv
except Exception:
    load_dotenv = None
if load_dotenv:
    load_dotenv(ROOT / ".env", override=False)
    load_dotenv(PHOTO_APP_DIR / ".env", override=False)
try:
    from google import genai
    from google.genai import types as genai_types
    from multi_card_extraction import (
        ModelQuotaExceeded,
        ModelResponseParseError,
        TemporaryModelUnavailable,
        identify_cards_sync,
    )
except Exception:
    genai = None
    genai_types = None
    identify_cards_sync = None
    TemporaryModelUnavailable = ModelQuotaExceeded = ModelResponseParseError = Exception
SETTINGS_PATH = ROOT / "lucas_settings.json"
DEFAULT_CARD_PIPELINE_DIR = ROOT / "CARD_PIPELINE"
CARD_PIPELINE_DIR = Path(os.environ.get("LUCAS_PIPELINE_DIR") or DEFAULT_CARD_PIPELINE_DIR)
WORKING_SHEETS_DIR = Path(os.environ.get("LUCAS_WORKING_SHEETS_DIR") or CARD_PIPELINE_DIR / "WORKING SHEETS")
INCOMING_SHEETS_DIR = CARD_PIPELINE_DIR / "INCOMING SHEETS"
RECEIVED_SHEETS_DIR = CARD_PIPELINE_DIR / "RECEIVED SHEETS"
COMPANY_SHEETS_DIR = CARD_PIPELINE_DIR / "COMPANY SHEETS"
SHEET_MARKERS_PATH = CARD_PIPELINE_DIR / "sheet_markers.json"
WEEKLY_COMPANY_SHEETS_PATH = CARD_PIPELINE_DIR / "weekly_company_sheets.json"
PROFIT_LEDGER_PATH = CARD_PIPELINE_DIR / "profit_ledger.json"
INVENTORY_LEDGER_PATH = CARD_PIPELINE_DIR / "inventory_ledger.json"
UNASSIGNED_PLAYERS_PATH = CARD_PIPELINE_DIR / "unassigned_players.json"
PLAYER_OVERRIDES_PATH = CARD_PIPELINE_DIR / "assignment_player_overrides.json"
SELLER_TERMS_PATH = CARD_PIPELINE_DIR / "ASSIGNMENT RULES" / "seller_terms.csv"
LUCAS_LOGO_PATH = ROOT / "assets" / "lucas.png"
CARDLADDER_EXTENSION_DIR = ROOT / "cardladder-autocomp" / "extension"
APP_TITLE = "L.U.C.A.S"
APP_SUBTITLE = "Lot Upload, Comping & Assignment System"

COMP_STRATEGY_DISPLAY = {
    "Average last 5": COMP_STRATEGY_AVERAGE,
    "Highest of last 5": COMP_STRATEGY_HIGH,
    "Lowest of last 5": COMP_STRATEGY_LOW,
    "Date weighted": COMP_STRATEGY_STALE_NEWEST,
}
COMP_SCOPE_EMPTY = "Empty Comps Only"
COMP_SCOPE_ALL = "Recomp All"
COMP_SOURCE_BOTH = "Card Ladder + CY"
COMP_SOURCE_CARD_LADDER = "Card Ladder"
COMP_SOURCE_CY = "CY"
NO_COMPANY_TAKES_LABEL = "NOBODY TAKES"
PROFIT_PERIOD_OPTIONS = ("5 Days", "Week", "Month", "Year", "YTD", "Total")
PROFIT_GRAPH_OPTIONS = ("Daily Trend", "Overall Profit")
EXPENSE_CATEGORY_OPTIONS = ("Travel", "Supplies", "Travel Meal", "Fees")
EXPENSE_LINK_OPTIONS = ("General", "Card", "Sheet")
ASSIGNMENT_CATEGORY_OPTIONS = (
    "basketball",
    "football",
    "baseball",
    "soccer",
    "hockey",
    "pokemon",
    "one piece",
    "wwe",
    "f1",
    "marvel",
    "disney",
    "star wars",
    "ufc",
)
ASSIGNMENT_CATEGORY_WEB_SIGNALS = {
    "basketball": ("basketball", "nba", "wnba", "ncaa basketball", "point guard", "shooting guard", "small forward", "power forward", "center"),
    "football": ("football", "nfl", "quarterback", "running back", "wide receiver", "linebacker", "cornerback", "defensive end"),
    "baseball": ("baseball", "mlb", "pitcher", "catcher", "shortstop", "outfielder", "first baseman", "second baseman", "third baseman"),
    "soccer": ("soccer", "footballer", "futbol", "fifa", "uefa", "premier league", "la liga", "serie a", "bundesliga"),
    "hockey": ("hockey", "nhl", "goaltender", "defenceman", "defenseman", "left wing", "right wing"),
    "pokemon": ("pokemon", "pokémon", "trading card game", "tcg", "pokédex", "pokedex"),
    "one piece": ("one piece", "straw hat", "manga", "anime", "pirate crew"),
    "wwe": ("wwe", "wwf", "professional wrestler", "pro wrestler", "wrestling"),
    "f1": ("formula 1", "formula one", "f1 driver", "grand prix", "racing driver", "motorsport"),
    "marvel": ("marvel", "marvel comics", "marvel cinematic", "superhero", "supervillain"),
    "disney": ("disney", "pixar", "disney character", "animated character"),
    "star wars": ("star wars", "jedi", "sith", "galactic", "lucasfilm"),
    "ufc": ("ufc", "mma", "mixed martial artist", "mixed martial arts", "ultimate fighting championship"),
}


def load_app_settings() -> dict[str, object]:
    if not SETTINGS_PATH.exists():
        return {}
    try:
        raw = json.loads(SETTINGS_PATH.read_text(encoding="utf-8"))
    except Exception:
        return {}
    return raw if isinstance(raw, dict) else {}


def save_app_settings(settings: dict[str, object]) -> None:
    atomic_write_json(SETTINGS_PATH, settings)


def ensure_mobile_pin(settings: dict[str, object]) -> str:
    pin = re.sub(r"\D", "", str(settings.get("mobile_pin") or ""))
    if len(pin) >= 4:
        return pin
    pin = f"{secrets.randbelow(1_000_000):06d}"
    settings["mobile_pin"] = pin
    save_app_settings(settings)
    return pin


def app_debug_log(message: str) -> None:
    try:
        APP_DEBUG_LOG.parent.mkdir(parents=True, exist_ok=True)
        stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with APP_DEBUG_LOG.open("a", encoding="utf-8") as handle:
            handle.write(f"[{stamp}] {message}\n")
    except Exception:
        pass


def is_google_sheet_url(value: object) -> bool:
    parsed = urllib.parse.urlparse(str(value or "").strip())
    return parsed.scheme in {"http", "https"} and parsed.netloc.lower().endswith("docs.google.com") and "/spreadsheets/" in parsed.path


def set_pipeline_root(path: Path, working_sheets_dir: Path | None = None) -> None:
    global CARD_PIPELINE_DIR, WORKING_SHEETS_DIR, INCOMING_SHEETS_DIR, RECEIVED_SHEETS_DIR, COMPANY_SHEETS_DIR, SHEET_MARKERS_PATH, WEEKLY_COMPANY_SHEETS_PATH, PROFIT_LEDGER_PATH, INVENTORY_LEDGER_PATH, UNASSIGNED_PLAYERS_PATH, PLAYER_OVERRIDES_PATH, SELLER_TERMS_PATH
    CARD_PIPELINE_DIR = Path(path).expanduser()
    WORKING_SHEETS_DIR = Path(working_sheets_dir).expanduser() if working_sheets_dir else CARD_PIPELINE_DIR / "WORKING SHEETS"
    INCOMING_SHEETS_DIR = CARD_PIPELINE_DIR / "INCOMING SHEETS"
    RECEIVED_SHEETS_DIR = CARD_PIPELINE_DIR / "RECEIVED SHEETS"
    COMPANY_SHEETS_DIR = CARD_PIPELINE_DIR / "COMPANY SHEETS"
    SHEET_MARKERS_PATH = CARD_PIPELINE_DIR / "sheet_markers.json"
    WEEKLY_COMPANY_SHEETS_PATH = CARD_PIPELINE_DIR / "weekly_company_sheets.json"
    PROFIT_LEDGER_PATH = CARD_PIPELINE_DIR / "profit_ledger.json"
    INVENTORY_LEDGER_PATH = CARD_PIPELINE_DIR / "inventory_ledger.json"
    UNASSIGNED_PLAYERS_PATH = CARD_PIPELINE_DIR / "unassigned_players.json"
    PLAYER_OVERRIDES_PATH = CARD_PIPELINE_DIR / "assignment_player_overrides.json"
    SELLER_TERMS_PATH = CARD_PIPELINE_DIR / "ASSIGNMENT RULES" / "seller_terms.csv"


def set_pipeline_from_working_dir(path: Path) -> None:
    working_dir = normalize_working_dir_selection(Path(path).expanduser())
    set_pipeline_root(working_dir.parent, working_dir)


def normalize_working_dir_selection(path: Path) -> Path:
    child = path / "WORKING SHEETS"
    if path.name.upper() != "WORKING SHEETS" and child.exists() and child.is_dir():
        return child
    return path


def initialize_pipeline_root() -> None:
    settings = load_app_settings()
    configured_working = str(settings.get("working_sheets_dir") or os.environ.get("LUCAS_WORKING_SHEETS_DIR") or "").strip()
    if configured_working:
        set_pipeline_from_working_dir(Path(configured_working))
        return
    configured = str(settings.get("pipeline_root") or "").strip()
    if configured:
        set_pipeline_root(Path(configured))


initialize_pipeline_root()


def company_sheet_week_start_for_time(moment: datetime) -> datetime.date:
    current_week_start = (moment - timedelta(days=moment.weekday())).date()
    if moment.weekday() == 6:
        return (moment + timedelta(days=1)).date()
    return current_week_start

DISPLAY_COLUMNS = (
    "excel_row",
    "source",
    "sheet_source",
    "cert_number",
    "grader",
    "card_title",
    "purchase_price",
    "card_ladder_value",
    "card_ladder_comps_average",
    "cy_value",
    "cy_confidence",
    "best_company",
    "estimated_payout",
    "status",
)

INTAKE_COLUMNS = (
    "excel_row",
    "source",
    "cert_number",
    "grader",
    "card_title",
    "purchase_price",
    "card_ladder_value",
    "card_ladder_comps_average",
    "cy_value",
    "cy_confidence",
    "status",
)

COMP_COLUMNS = (
    "excel_row",
    "source",
    "cert_number",
    "grader",
    "card_title",
    "purchase_price",
    "card_ladder_value",
    "card_ladder_comps_average",
    "cy_value",
    "cy_confidence",
    "best_company",
    "estimated_payout",
    "status",
    "sheet_source",
)

RECEIVE_COLUMNS = (
    "excel_row",
    "source",
    "sheet_source",
    "cert_number",
    "grader",
    "card_title",
    "purchase_price",
    "card_ladder_value",
    "card_ladder_comps_average",
    "cy_value",
    "cy_confidence",
    "best_company",
    "estimated_payout",
    "status",
    "company_pile",
)

REVIEW_COLUMNS = DISPLAY_COLUMNS

INVENTORY_TABLE_COLUMNS = (
    "date",
    "person",
    "sport",
    "cert",
    "grader",
    "card",
    "purchase",
    "card_ladder",
    "comps",
    "cy_estimate",
    "cy_confidence",
    "company",
    "payout",
    "source",
    "status",
)

ADD_INTAKE_ROW_IID = "__add_intake_row__"
ADD_REVIEW_ROW_IID = "__add_review_row__"

EDITABLE_COLUMNS = {
    "source",
    "cert_number",
    "grader",
    "card_title",
    "purchase_price",
    "card_ladder_value",
    "card_ladder_comps_average",
    "cy_value",
    "cy_confidence",
}

HEADINGS = {
    "excel_row": "Row",
    "source": "Source",
    "sheet_source": "Sheet Source",
    "cert_number": "Cert #",
    "grader": "Company",
    "card_title": "Card",
    "purchase_price": "Purchase",
    "card_ladder_value": "Card Ladder",
    "card_ladder_comps_average": "Comps",
    "cy_value": "CY Estimate",
    "cy_confidence": "CY Confidence",
    "best_company": "Best Company",
    "estimated_payout": "Est. Payout",
    "status": "Status",
    "company_pile": "Company Pile",
}

COLUMN_WIDTHS = {
    "excel_row": 52,
    "source": 130,
    "sheet_source": 150,
    "cert_number": 110,
    "grader": 86,
    "card_title": 390,
    "purchase_price": 90,
    "card_ladder_value": 100,
    "card_ladder_comps_average": 100,
    "cy_value": 100,
    "cy_confidence": 110,
    "best_company": 130,
    "estimated_payout": 100,
    "status": 160,
    "company_pile": 105,
}

INVENTORY_HEADINGS = {
    "date": "Date",
    "person": "Person",
    "sport": "Sport",
    "cert": "Cert",
    "grader": "Grader",
    "card": "Card",
    "purchase": "Purchase",
    "card_ladder": "Card Ladder",
    "comps": "Comps",
    "cy_estimate": "CY Estimate",
    "cy_confidence": "CY Confidence",
    "company": "Best Company",
    "payout": "Est. Payout",
    "source": "Source Sheet",
    "status": "Status",
}

INVENTORY_COLUMN_WIDTHS = {
    "date": 95,
    "person": 130,
    "sport": 95,
    "cert": 110,
    "grader": 80,
    "card": 320,
    "purchase": 100,
    "card_ladder": 100,
    "comps": 100,
    "cy_estimate": 100,
    "cy_confidence": 110,
    "company": 140,
    "payout": 100,
    "source": 170,
    "status": 110,
}


class CardPipelineApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        app_debug_log("app_start")
        self.title(f"{APP_TITLE} - {APP_SUBTITLE}")
        self.geometry("1420x820")
        self.minsize(1120, 680)
        self.logo_image: tk.PhotoImage | None = None

        self.events: queue.Queue[str] = queue.Queue()
        self.intake_rows: list[WorkbookRow] = []
        self.intake_sources: dict[int, str] = {}
        self.intake_sheet_sources: dict[int, str] = {}
        self.row_sources: dict[int, str] = {}
        self.comp_sheet_sources: dict[int, str] = {}
        self.review_rows: list[WorkbookRow] = []
        self.review_sources: dict[int, str] = {}
        self.review_sheet_sources: dict[int, str] = {}
        self.incoming_cert_index: dict[str, dict[str, object]] = {}
        self.comp_output_saved = True
        self.lucas_identity = local_identity(SETTINGS_PATH)
        self.app_settings = load_app_settings()
        self.mobile_pin = ensure_mobile_pin(self.app_settings)
        self.state = BridgeState()
        self.state.on_update = lambda: self.events.put("comp_refresh")
        self.state.mobile_pin_provider = lambda: self.mobile_pin
        self.state.mobile_inventory_search = self.mobile_inventory_search
        self.state.mobile_inventory_add = self.mobile_inventory_add
        self.state.mobile_inventory_mark_sold = self.mobile_inventory_mark_sold
        self.state.mobile_card_identify = self.mobile_card_identify
        self.state.mobile_profit_summary = self.mobile_profit_summary
        self.state.mobile_expense_add = self.mobile_expense_add
        self.state.mobile_payouts = self.mobile_payouts
        self.bridge = BridgeServer(self.state)
        self.bridge.start()
        self._refresh_keep_source_registry()
        app_debug_log(f"bridge_started started={self.bridge.started} port={self.bridge.port} error={self.bridge.error}")
        mobile_url = self._mobile_app_url()
        self.bridge_status_text = (
            f"Card Ladder bridge running at http://127.0.0.1:{self.bridge.port} | Mobile: {mobile_url} PIN {self.mobile_pin}"
            if self.bridge.started
            else f"Card Ladder bridge failed to start: {self.bridge.error}"
        )

        self.input_mode = tk.StringVar(value="Barcode Scanner")
        self.review_mode = tk.StringVar(value="Automatic Receive")
        self.review_input_mode = tk.StringVar(value="Barcode Scanner")
        self.comp_strategy_label = tk.StringVar(value="Average last 5")
        self.comp_scope_label = tk.StringVar(value=COMP_SCOPE_EMPTY)
        self.comp_source_label = tk.StringVar(value=COMP_SOURCE_BOTH)
        self.working_sheet_title = tk.StringVar()
        self.create_network_mode_var = tk.BooleanVar(value=bool(self.app_settings.get("network_mode")))
        self.seller_terms_seller_var = tk.StringVar()
        self.seller_terms_sheet_type_var = tk.StringVar()
        self.selected_working_sheet = tk.StringVar()
        self.summary_var = tk.StringVar(value="Choose a create mode to begin.")
        self.status_var = tk.StringVar(value="Card Ladder bridge starting...")
        self.bridge_status_var = tk.StringVar(value=self.bridge_status_text)
        self.pipeline_root_var = tk.StringVar(value=str(CARD_PIPELINE_DIR))

        self.scan_cert = tk.StringVar()
        self.scan_grader = tk.StringVar(value="PSA")
        self.scan_card = tk.StringVar()
        self.scan_status = tk.StringVar(value="Scanning station is off.")
        self.scan_entry: ttk.Entry | None = None
        self.cell_editor: ttk.Entry | None = None
        self.cell_edit: tuple[ttk.Treeview, str, str] | None = None
        self.column_widths_by_tree: dict[int, dict[str, int]] = {}
        self.scanning_station_active = False

        self.file_path = tk.StringVar()
        self.sheet_name = tk.StringVar()
        self.photo_paths: list[Path] = []
        self.photo_status = tk.StringVar(value="No photos selected.")
        self.photo_worker: threading.Thread | None = None
        self.photo_client = None
        self.review_scan_cert = tk.StringVar()
        self.review_scan_entry: ttk.Entry | None = None
        self.review_scanning_active = False
        self.review_status = tk.StringVar(value="Receive station is off.")
        self.assignment_progress_value = tk.DoubleVar(value=0)
        self.review_photo_paths: list[Path] = []
        self.review_photo_status = tk.StringVar(value="No receive photos selected.")
        self.review_photo_worker: threading.Thread | None = None
        self._load_player_overrides()
        self.assignment_engine = AssignmentEngine.load()
        self.assignment_recommendation_job = 0
        self.assignment_recommendation_running = False
        self.assignment_recommendation_after_id: str | None = None
        self.assignment_config_status = tk.StringVar(value=self._assignment_config_status())
        self._ensure_company_sheet_folders()
        self._ensure_weekly_company_sheets_due()
        self.received_sheet_paths: dict[str, Path] = {}
        self.selected_received_sheet = tk.StringVar()
        self.working_sheet_paths: dict[str, Path] = {}
        self.home_sheet_kind = tk.StringVar(value="Incoming")
        self.home_sheet_paths: dict[str, dict[str, Path]] = {"Incoming": {}, "Working": {}, "Received": {}}
        self.home_sheet_summaries: dict[str, dict[str, object]] = {}
        self.home_sheet_markers: dict[str, dict[str, object]] = self._load_sheet_markers()
        self.deleted_sheet_marker_keys: set[str] = set()
        self.home_selected_sheet_key = ""
        self.home_person_var = tk.StringVar()
        self.payout_person_var = tk.StringVar()
        self.payout_status_var = tk.StringVar(value="No unpaid sheets loaded.")
        self.payout_summary_people: dict[str, str] = {}
        self.payout_detail_keys: dict[str, str] = {}
        self.inventory_status_var = tk.StringVar(value="No inventory loaded.")
        self.inventory_metric_var = tk.StringVar(value="")
        self.inventory_person_var = tk.StringVar()
        self.inventory_sport_var = tk.StringVar()
        self.inventory_search_var = tk.StringVar()
        self.inventory_min_var = tk.StringVar()
        self.inventory_max_var = tk.StringVar()
        self.inventory_rows: list[dict[str, object]] = []
        self.filtered_inventory_rows: list[dict[str, object]] = []
        self.inventory_tree_records: dict[str, dict[str, object]] = {}
        self.profit_status_var = tk.StringVar(value="No profit ledger loaded.")
        self.profit_metric_var = tk.StringVar(value="")
        self.profit_person_var = tk.StringVar()
        self.profit_period_var = tk.StringVar(value="Total")
        self.profit_graph_var = tk.StringVar(value="Daily Trend")
        self.profit_view_mode = tk.StringVar(value="Sold Cards")
        self.profit_rows: list[dict[str, object]] = []
        self.filtered_profit_rows: list[dict[str, object]] = []
        self.profit_tree_records: dict[str, dict[str, object]] = {}

        self._build_ui()
        self._show_mode()
        self.refresh_profit_tab()
        self._poll_events()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.status_var.set(self.bridge_status_text)
        self.after(100, self._start_startup_refresh)
        self.after(5 * 60 * 1000, self._weekly_company_sheet_timer)

    def _on_close(self) -> None:
        self.destroy()

    def _build_ui(self) -> None:
        palette = {
            "bg": "#121212",
            "surface": "#181818",
            "panel": "#1f1f1f",
            "panel_high": "#242424",
            "field": "#2a2a2a",
            "border": "#333333",
            "muted": "#b3b3b3",
            "button": "#1ed760",
            "button_hover": "#1fdf64",
            "button_pressed": "#169c46",
            "soft_button": "#2a2a2a",
            "soft_button_hover": "#3a3a3a",
            "text": "#ffffff",
            "subtle_text": "#d9d9d9",
            "selection": "#1db954",
            "warning": "#5a4a14",
            "danger": "#5a1f1f",
        }
        self.configure(bg=palette["bg"])
        self.option_add("*TCombobox*Listbox.background", palette["field"])
        self.option_add("*TCombobox*Listbox.foreground", palette["text"])
        self.option_add("*TCombobox*Listbox.selectBackground", palette["selection"])
        self.option_add("*TCombobox*Listbox.selectForeground", "#000000")
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure(".", font=("Segoe UI", 10))
        style.configure("App.TFrame", background=palette["bg"])
        style.configure("Panel.TFrame", background=palette["panel"])
        style.configure("Header.TFrame", background=palette["surface"])
        style.configure("Header.TLabel", background=palette["surface"])
        style.configure("HeaderTitle.TLabel", background=palette["surface"], foreground=palette["text"], font=("Segoe UI Semibold", 22))
        style.configure("HeaderSub.TLabel", background=palette["surface"], foreground=palette["muted"])
        style.configure("BridgeBadge.TLabel", background=palette["panel_high"], foreground=palette["button"], font=("Segoe UI Semibold", 9), padding=(12, 7))
        style.configure("Panel.TLabel", background=palette["panel"], foreground=palette["text"])
        style.configure("Muted.TLabel", background=palette["panel"], foreground=palette["muted"])
        style.configure("Status.TLabel", background=palette["bg"], foreground=palette["muted"])
        style.configure("Panel.TCheckbutton", background=palette["panel"], foreground=palette["text"])
        style.map(
            "Panel.TCheckbutton",
            background=[("active", palette["panel"])],
            foreground=[("active", palette["text"]), ("disabled", "#777777")],
        )
        style.configure(
            "ChromeTab.TButton",
            font=("Segoe UI Semibold", 9),
            padding=(12, 6),
            background=palette["soft_button"],
            foreground=palette["muted"],
            borderwidth=0,
            relief=tk.FLAT,
        )
        style.map(
            "ChromeTab.TButton",
            background=[("pressed", palette["border"]), ("active", palette["soft_button_hover"])],
            foreground=[("active", palette["text"])],
        )
        style.configure(
            "ChromeTabActive.TButton",
            font=("Segoe UI Semibold", 9),
            padding=(12, 6),
            background=palette["panel_high"],
            foreground=palette["text"],
            borderwidth=0,
            relief=tk.FLAT,
        )
        style.map(
            "ChromeTabActive.TButton",
            background=[("pressed", palette["panel_high"]), ("active", palette["panel_high"])],
            foreground=[("active", palette["text"])],
        )
        style.configure(
            "Primary.TButton",
            font=("Segoe UI Semibold", 10),
            padding=(18, 9),
            background=palette["button"],
            foreground="#000000",
            borderwidth=0,
            focusthickness=0,
            relief=tk.FLAT,
        )
        style.map(
            "Primary.TButton",
            background=[("pressed", palette["button_pressed"]), ("active", palette["button_hover"]), ("disabled", "#535353")],
            foreground=[("disabled", "#b3b3b3")],
            relief=[("pressed", tk.FLAT), ("!pressed", tk.FLAT)],
        )
        style.configure(
            "Soft.TButton",
            font=("Segoe UI Semibold", 10),
            padding=(16, 9),
            background=palette["soft_button"],
            foreground=palette["text"],
            borderwidth=0,
            focusthickness=0,
            relief=tk.FLAT,
        )
        style.map(
            "Soft.TButton",
            background=[("pressed", palette["border"]), ("active", palette["soft_button_hover"]), ("disabled", "#1a1a1a")],
            foreground=[("disabled", "#777777")],
            relief=[("pressed", tk.FLAT), ("!pressed", tk.FLAT)],
        )
        style.configure(
            "TEntry",
            fieldbackground=palette["field"],
            background=palette["field"],
            foreground=palette["text"],
            insertcolor=palette["text"],
            bordercolor=palette["border"],
            lightcolor=palette["border"],
            darkcolor=palette["border"],
            padding=(8, 7),
        )
        style.map("TEntry", bordercolor=[("focus", palette["selection"])])
        style.configure(
            "TCombobox",
            fieldbackground=palette["field"],
            background=palette["field"],
            foreground=palette["text"],
            arrowcolor=palette["muted"],
            bordercolor=palette["border"],
            lightcolor=palette["border"],
            darkcolor=palette["border"],
            padding=(8, 6),
        )
        style.map(
            "TCombobox",
            fieldbackground=[("readonly", palette["field"])],
            foreground=[("readonly", palette["text"])],
            bordercolor=[("focus", palette["selection"])],
            arrowcolor=[("active", palette["text"])],
        )
        style.configure("TNotebook", background=palette["bg"], borderwidth=0, tabmargins=(0, 0, 0, 0))
        style.configure(
            "TNotebook.Tab",
            background=palette["bg"],
            foreground=palette["muted"],
            padding=(18, 10),
            borderwidth=0,
            font=("Segoe UI Semibold", 10),
        )
        style.map(
            "TNotebook.Tab",
            background=[("selected", palette["panel"]), ("active", palette["panel_high"])],
            foreground=[("selected", palette["text"]), ("active", palette["text"])],
        )
        style.configure("Vertical.TScrollbar", background=palette["field"], troughcolor=palette["panel"], bordercolor=palette["panel"], arrowcolor=palette["muted"])
        style.configure("Horizontal.TScrollbar", background=palette["field"], troughcolor=palette["panel"], bordercolor=palette["panel"], arrowcolor=palette["muted"])
        style.configure(
            "Assignment.Horizontal.TProgressbar",
            background="#16a34a",
            troughcolor="#ffffff",
            bordercolor="#d7dde3",
            lightcolor="#16a34a",
            darkcolor="#15803d",
        )
        style.configure("Treeview", rowheight=34, font=("Segoe UI", 10), background=palette["panel"], fieldbackground=palette["panel"], foreground=palette["subtle_text"], borderwidth=0)
        style.configure("Treeview.Heading", font=("Segoe UI Semibold", 9), background=palette["panel_high"], foreground=palette["muted"], padding=(10, 8), borderwidth=0)
        style.map("Treeview", background=[("selected", palette["selection"])], foreground=[("selected", "#000000")])

        header = ttk.Frame(self, style="Header.TFrame", padding=(18, 16))
        header.pack(fill=tk.X)
        if LUCAS_LOGO_PATH.exists():
            try:
                self.logo_image = tk.PhotoImage(file=str(LUCAS_LOGO_PATH)).subsample(6, 6)
                self.iconphoto(False, self.logo_image)
                ttk.Label(header, image=self.logo_image, style="Header.TLabel").pack(side=tk.LEFT, padx=(0, 14))
            except tk.TclError:
                self.logo_image = None
        title_group = ttk.Frame(header, style="Header.TFrame")
        title_group.pack(side=tk.LEFT)
        ttk.Label(title_group, text=APP_TITLE, style="HeaderTitle.TLabel").pack(anchor=tk.W)
        ttk.Label(title_group, text=APP_SUBTITLE, style="HeaderSub.TLabel").pack(anchor=tk.W, pady=(3, 0))
        ttk.Label(header, textvariable=self.bridge_status_var, style="BridgeBadge.TLabel").pack(side=tk.RIGHT, padx=(16, 0))
        ttk.Button(header, text="Working Folder", command=self.choose_working_folder, style="Soft.TButton").pack(side=tk.RIGHT, padx=(16, 0))

        self.tabs = ttk.Notebook(self)
        self.tabs.pack(fill=tk.BOTH, expand=True, padx=18, pady=(16, 12))
        self.home_tab = ttk.Frame(self.tabs, style="App.TFrame", padding=0)
        self.intake_tab = ttk.Frame(self.tabs, style="App.TFrame", padding=0)
        self.comp_tab = ttk.Frame(self.tabs, style="App.TFrame", padding=0)
        self.receive_tab = ttk.Frame(self.tabs, style="App.TFrame", padding=0)
        self.review_tab = ttk.Frame(self.tabs, style="App.TFrame", padding=0)
        self.payouts_tab = ttk.Frame(self.tabs, style="App.TFrame", padding=0)
        self.inventory_tab = ttk.Frame(self.tabs, style="App.TFrame", padding=0)
        self.profit_tab = ttk.Frame(self.tabs, style="App.TFrame", padding=0)
        self.tabs.add(self.home_tab, text="Home")
        self.tabs.add(self.intake_tab, text="Create")
        self.tabs.add(self.comp_tab, text="Comp")
        self.tabs.add(self.receive_tab, text="Receive")
        self.tabs.add(self.review_tab, text="Assignment")
        self.tabs.add(self.payouts_tab, text="Payouts/Tabs")
        self.tabs.add(self.inventory_tab, text="Inventory")
        self.tabs.add(self.profit_tab, text="Profit")
        self.row_trees: list[ttk.Treeview] = []

        self._build_home_tab(palette)

        intake_controls = ttk.Frame(self.intake_tab, style="Panel.TFrame", padding=(16, 12))
        intake_controls.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(intake_controls, text="Input Mode", style="Panel.TLabel").grid(row=0, column=0, sticky="w")
        mode = ttk.Combobox(
            intake_controls,
            textvariable=self.input_mode,
            state="readonly",
            values=["Barcode Scanner", "Manual Entry", "Photo OCR", "Existing Spreadsheet"],
            width=22,
        )
        mode.grid(row=0, column=1, sticky="w", padx=(8, 16))
        mode.bind("<<ComboboxSelected>>", lambda _event: self._show_mode())
        ttk.Button(intake_controls, text="Delete Selected", command=self.delete_selected_intake_rows, style="Soft.TButton").grid(row=0, column=2, sticky="w", padx=(0, 8))
        ttk.Button(intake_controls, text="Clear Rows", command=self.clear_rows, style="Soft.TButton").grid(row=0, column=3, sticky="w")
        ttk.Checkbutton(
            intake_controls,
            text="Network Mode",
            variable=self.create_network_mode_var,
            command=self._toggle_create_network_mode,
            style="Panel.TCheckbutton",
        ).grid(row=0, column=4, sticky="e", padx=(16, 0))
        intake_controls.columnconfigure(4, weight=1)
        self.network_seller_label = ttk.Label(intake_controls, text="Seller", style="Muted.TLabel")
        self.network_seller_label.grid(row=1, column=0, sticky="w", pady=(10, 0))
        self.seller_terms_seller_combo = ttk.Combobox(intake_controls, textvariable=self.seller_terms_seller_var, width=24)
        self.seller_terms_seller_combo.grid(row=1, column=1, sticky="w", padx=(8, 16), pady=(10, 0))
        self._bind_person_autocomplete(self.seller_terms_seller_combo, refresh_callback=self.apply_create_seller_terms)
        self.seller_terms_seller_combo.bind("<<ComboboxSelected>>", lambda _event: self.apply_create_seller_terms(), add="+")
        self.network_sheet_type_label = ttk.Label(intake_controls, text="Sheet Type", style="Muted.TLabel")
        self.network_sheet_type_label.grid(row=1, column=2, sticky="e", padx=(0, 6), pady=(10, 0))
        self.seller_terms_sheet_type_combo = ttk.Combobox(intake_controls, textvariable=self.seller_terms_sheet_type_var, width=18)
        self.seller_terms_sheet_type_combo.grid(row=1, column=3, sticky="w", padx=(0, 16), pady=(10, 0))
        self.seller_terms_sheet_type_combo.configure(postcommand=self._refresh_seller_terms_dropdowns)
        self.seller_terms_sheet_type_combo.bind("<<ComboboxSelected>>", lambda _event: self.apply_create_seller_terms(), add="+")
        ttk.Label(intake_controls, textvariable=self.summary_var, style="Muted.TLabel").grid(row=2, column=0, columnspan=5, sticky="w", pady=(10, 0))
        self._set_create_network_controls_visible(self._network_mode_enabled())

        self.mode_host = ttk.Frame(self.intake_tab, style="Panel.TFrame", padding=(16, 12))
        self.mode_host.pack(fill=tk.X, pady=(0, 10))
        self.intake_tree = self._build_table(self.intake_tab, editable=True, columns=INTAKE_COLUMNS)
        intake_save = ttk.Frame(self.intake_tab, style="Panel.TFrame", padding=(16, 12))
        intake_save.pack(fill=tk.X, pady=(10, 0))
        ttk.Label(intake_save, text="Working Sheet Title", style="Panel.TLabel").pack(side=tk.LEFT, padx=(0, 8))
        ttk.Entry(intake_save, textvariable=self.working_sheet_title, width=42).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(intake_save, text="Save as Working Sheet", command=self.save_working_sheet, style="Primary.TButton").pack(side=tk.LEFT)

        comp_body = ttk.Frame(self.comp_tab, style="App.TFrame")
        comp_body.pack(fill=tk.BOTH, expand=True)
        sheet_panel = ttk.Frame(comp_body, style="Panel.TFrame", padding=(12, 12))
        sheet_panel.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))
        ttk.Label(sheet_panel, text="Active Sheets", style="Panel.TLabel").pack(anchor=tk.W)
        self.working_sheet_list = tk.Listbox(
            sheet_panel,
            width=34,
            height=24,
            activestyle="none",
            exportselection=False,
            bg=palette["panel"],
            fg=palette["subtle_text"],
            selectbackground=palette["selection"],
            selectforeground="#000000",
            highlightthickness=1,
            highlightbackground=palette["border"],
            highlightcolor=palette["selection"],
            relief=tk.FLAT,
            borderwidth=0,
            font=("Segoe UI", 10),
        )
        self.working_sheet_list.pack(fill=tk.Y, expand=True, pady=(8, 8))
        self.working_sheet_list.bind("<Double-Button-1>", lambda _event: self.load_selected_working_sheet())
        ttk.Button(sheet_panel, text="Load Selected Sheet", command=self.load_selected_working_sheet, style="Primary.TButton").pack(fill=tk.X, pady=(0, 8))
        ttk.Button(sheet_panel, text="Refresh Sheets", command=self.refresh_pipeline, style="Soft.TButton").pack(fill=tk.X)
        comp_main = ttk.Frame(comp_body, style="App.TFrame")
        comp_main.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.comp_tree = self._build_table(comp_main, editable=True, columns=COMP_COLUMNS)
        comp_controls = ttk.Frame(comp_main, style="Panel.TFrame", padding=(16, 12))
        comp_controls.pack(fill=tk.X, pady=(10, 0))
        ttk.Button(comp_controls, text="Save Output", command=self.save_output, style="Soft.TButton").pack(side=tk.RIGHT, padx=(8, 0))
        ttk.Button(comp_controls, text="Run All Comps", command=self.run_all_comps, style="Primary.TButton").pack(side=tk.RIGHT, padx=(8, 0))
        ttk.Button(comp_controls, text="Stop Run", command=self.stop_comp_run, style="Soft.TButton").pack(side=tk.RIGHT, padx=(8, 0))
        ttk.Button(comp_controls, text="Clear Comp Rows", command=self.clear_comp_rows, style="Soft.TButton").pack(side=tk.RIGHT, padx=(8, 0))
        self.comp_scope_combo = ttk.Combobox(
            comp_controls,
            textvariable=self.comp_scope_label,
            state="readonly",
            values=(COMP_SCOPE_EMPTY, COMP_SCOPE_ALL),
            width=17,
        )
        self.comp_scope_combo.pack(side=tk.RIGHT, padx=(8, 0))
        ttk.Label(comp_controls, text="Run Scope", style="Panel.TLabel").pack(side=tk.RIGHT)
        self.comp_source_combo = ttk.Combobox(
            comp_controls,
            textvariable=self.comp_source_label,
            state="readonly",
            values=(COMP_SOURCE_BOTH, COMP_SOURCE_CARD_LADDER, COMP_SOURCE_CY),
            width=18,
        )
        self.comp_source_combo.pack(side=tk.RIGHT, padx=(8, 0))
        ttk.Label(comp_controls, text="Run", style="Panel.TLabel").pack(side=tk.RIGHT)
        self.comp_method_combo = ttk.Combobox(
            comp_controls,
            textvariable=self.comp_strategy_label,
            state="readonly",
            values=list(COMP_STRATEGY_DISPLAY.keys()),
            width=20,
        )
        self.comp_method_combo.pack(side=tk.RIGHT, padx=(8, 0))
        self.comp_method_combo.bind("<<ComboboxSelected>>", self.recalculate_comp_method)
        ttk.Label(comp_controls, text="Comp Method", style="Panel.TLabel").pack(side=tk.RIGHT)

        receive_controls = ttk.Frame(self.receive_tab, style="Panel.TFrame", padding=(16, 12))
        receive_controls.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(receive_controls, text="Receive Mode", style="Panel.TLabel").grid(row=0, column=0, sticky="w")
        review_mode = ttk.Combobox(
            receive_controls,
            textvariable=self.review_mode,
            state="readonly",
            values=["Automatic Receive", "Manual Receive"],
            width=20,
        )
        review_mode.grid(row=0, column=1, sticky="w", padx=(8, 16))
        review_mode.bind("<<ComboboxSelected>>", lambda _event: self._show_review_mode())
        ttk.Label(receive_controls, textvariable=self.review_status, style="Muted.TLabel").grid(row=1, column=0, columnspan=5, sticky="w", pady=(10, 0))
        receive_controls.columnconfigure(4, weight=1)

        self.review_mode_host = ttk.Frame(self.receive_tab, style="Panel.TFrame", padding=(16, 12))
        self.review_mode_host.pack(fill=tk.X, pady=(0, 10))
        self.receive_tree = self._build_table(self.receive_tab, editable=True, columns=RECEIVE_COLUMNS)
        receive_bottom = ttk.Frame(self.receive_tab, style="Panel.TFrame", padding=(16, 12))
        receive_bottom.pack(fill=tk.X, pady=(10, 0))
        ttk.Button(receive_bottom, text="Mark Received in Sheets", command=self.mark_review_received_in_sheets, style="Primary.TButton").pack(side=tk.RIGHT, padx=(8, 0))
        ttk.Button(receive_bottom, text="Refresh Incoming Sheets", command=self.refresh_incoming_index, style="Soft.TButton").pack(side=tk.RIGHT, padx=(8, 0))
        ttk.Button(receive_bottom, text="Delete Selected", command=self.delete_selected_review_rows, style="Soft.TButton").pack(side=tk.RIGHT, padx=(8, 0))
        ttk.Button(receive_bottom, text="Clear Receive Rows", command=self.clear_review_rows, style="Soft.TButton").pack(side=tk.RIGHT)

        review_controls = ttk.Frame(self.review_tab, style="Panel.TFrame", padding=(16, 12))
        review_controls.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(review_controls, text="Received Sheet", style="Panel.TLabel").grid(row=0, column=0, sticky="w")
        self.received_sheet_combo = ttk.Combobox(review_controls, textvariable=self.selected_received_sheet, state="readonly", width=32)
        self.received_sheet_combo.grid(row=0, column=1, sticky="ew", padx=(8, 8))
        ttk.Button(review_controls, text="Load", command=self.load_selected_received_sheet_for_review, style="Primary.TButton").grid(row=0, column=2, sticky="w", padx=(0, 8))
        ttk.Button(review_controls, text="Refresh", command=self.refresh_received_sheets, style="Soft.TButton").grid(row=0, column=3, sticky="w")
        ttk.Button(review_controls, text="Assignment Rules", command=self.open_assignment_rules, style="Soft.TButton").grid(row=0, column=4, sticky="w", padx=(8, 0))
        ttk.Button(review_controls, text="Unassigned Players", command=self.open_unassigned_players_dialog, style="Soft.TButton").grid(row=0, column=5, sticky="w", padx=(8, 0))
        review_controls.columnconfigure(1, weight=1)
        ttk.Label(review_controls, textvariable=self.review_status, style="Muted.TLabel").grid(row=1, column=0, columnspan=6, sticky="w", pady=(10, 0))
        ttk.Label(review_controls, textvariable=self.assignment_config_status, style="Muted.TLabel").grid(row=2, column=0, columnspan=6, sticky="w", pady=(4, 0))
        self.assignment_progress = ttk.Progressbar(
            review_controls,
            style="Assignment.Horizontal.TProgressbar",
            variable=self.assignment_progress_value,
            maximum=100,
            mode="determinate",
        )
        self.assignment_progress.grid(row=3, column=0, columnspan=6, sticky="ew", pady=(8, 0))
        self.review_tree = self._build_table(self.review_tab, editable=True, columns=REVIEW_COLUMNS)
        review_bottom = ttk.Frame(self.review_tab, style="Panel.TFrame", padding=(16, 12))
        review_bottom.pack(fill=tk.X, pady=(10, 0))
        ttk.Button(review_bottom, text="Delete Selected", command=self.delete_selected_review_rows, style="Soft.TButton").pack(side=tk.RIGHT, padx=(8, 0))
        ttk.Button(review_bottom, text="Clear Assignment Rows", command=self.clear_review_rows, style="Soft.TButton").pack(side=tk.RIGHT)
        self._show_review_mode()
        self._build_payouts_tab()
        self._build_inventory_tab()
        self._build_profit_tab()

        bottom = ttk.Frame(self, style="App.TFrame", padding=(16, 0, 16, 14))
        bottom.pack(fill=tk.X)
        ttk.Label(bottom, textvariable=self.status_var, style="Status.TLabel").pack(side=tk.LEFT)

    def _build_table(self, parent: ttk.Frame, editable: bool = False, columns: tuple[str, ...] = DISPLAY_COLUMNS) -> ttk.Treeview:
        content = ttk.Frame(parent, style="Panel.TFrame", padding=(1, 1))
        content.pack(fill=tk.BOTH, expand=True)
        tree = ttk.Treeview(content, columns=columns, show="headings", selectmode="extended")
        setattr(tree, "_display_columns", columns)
        for col in columns:
            tree.heading(col, text=HEADINGS[col], anchor=tk.W)
            tree.column(col, width=COLUMN_WIDTHS[col], minwidth=45, stretch=False)
        tree.grid(row=0, column=0, sticky="nsew")
        y_scroll = ttk.Scrollbar(content, orient=tk.VERTICAL, command=tree.yview)
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll = ttk.Scrollbar(content, orient=tk.HORIZONTAL, command=tree.xview)
        x_scroll.grid(row=1, column=0, sticky="ew")
        tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)
        tree.tag_configure("duplicate_cert", background="#4a3d12", foreground="#fff3b0")
        tree.tag_configure("no_sheet_found", background="#4a1717", foreground="#ffd1d1")
        tree.tag_configure("add_review_row", background="#242424", foreground="#1ed760")
        if editable:
            tree.bind("<Double-1>", self._begin_cell_edit)
            tree.bind("<Button-1>", self._handle_table_click, add="+")
            tree.bind("<Delete>", self._delete_selected_table_rows)
        tree.bind("<ButtonRelease-1>", lambda _event, target=tree: self._remember_column_widths(target), add="+")
        content.columnconfigure(0, weight=1)
        content.rowconfigure(0, weight=1)
        setattr(tree, "_table_frame", content)
        self.row_trees.append(tree)
        self.column_widths_by_tree[id(tree)] = {col: COLUMN_WIDTHS[col] for col in columns}
        return tree

    def _build_home_tab(self, palette: dict[str, str]) -> None:
        body = ttk.Frame(self.home_tab, style="App.TFrame")
        body.pack(fill=tk.BOTH, expand=True)

        sheet_panel = ttk.Frame(body, style="Panel.TFrame", padding=(12, 12))
        sheet_panel.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))
        sheet_panel.configure(width=360)
        sheet_panel.pack_propagate(False)
        person_row = ttk.Frame(sheet_panel, style="Panel.TFrame")
        person_row.pack(fill=tk.X, pady=(0, 8))
        ttk.Label(person_row, text="Person", style="Muted.TLabel").pack(side=tk.LEFT)
        self.home_person_combo = ttk.Combobox(person_row, textvariable=self.home_person_var, width=24)
        self.home_person_combo.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(8, 0))
        self._bind_person_autocomplete(self.home_person_combo, refresh_callback=self._on_home_person_filter_changed)
        self.home_person_combo.bind("<<ComboboxSelected>>", lambda _event: self._on_home_person_filter_changed(), add="+")
        toggle_row = tk.Frame(sheet_panel, bg=palette["panel"])
        toggle_row.pack(fill=tk.X, pady=(0, 8))
        self.home_tab_palette = palette
        self.home_incoming_tab = self._build_home_tab_button(toggle_row, "Incoming", lambda: self._set_home_sheet_kind("Incoming"))
        self.home_incoming_tab.grid(row=0, column=0, sticky="ew", padx=(0, 4))
        self.home_working_tab = self._build_home_tab_button(toggle_row, "Working", lambda: self._set_home_sheet_kind("Working"))
        self.home_working_tab.grid(row=0, column=1, sticky="ew", padx=(0, 4))
        self.home_received_tab = self._build_home_tab_button(toggle_row, "Received", lambda: self._set_home_sheet_kind("Received"))
        self.home_received_tab.grid(row=0, column=2, sticky="ew", padx=(0, 4))
        self.home_edit_markers_tab = self._build_home_tab_button(toggle_row, "Edit Markers", self.open_sheet_marker_editor)
        self.home_edit_markers_tab.grid(row=0, column=3, sticky="ew")
        for col in range(4):
            toggle_row.columnconfigure(col, weight=1, uniform="home_tabs")
        self.home_sheet_list = tk.Listbox(
            sheet_panel,
            width=1,
            height=28,
            activestyle="none",
            exportselection=False,
            bg=palette["panel"],
            fg=palette["subtle_text"],
            selectbackground=palette["selection"],
            selectforeground="#000000",
            highlightthickness=1,
            highlightbackground=palette["border"],
            highlightcolor=palette["selection"],
            relief=tk.FLAT,
            borderwidth=0,
            font=("Segoe UI", 10),
        )
        self.home_sheet_list.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        self.home_sheet_list.bind("<<ListboxSelect>>", lambda _event: self._load_home_selected_marker())
        self.home_sheet_list.bind("<Button-3>", self._show_home_sheet_context_menu)
        self.home_sheet_list.bind("<Button-2>", self._show_home_sheet_context_menu)
        ttk.Button(sheet_panel, text="Refresh Home", command=self.refresh_home, style="Primary.TButton").pack(fill=tk.X)

        right = ttk.Frame(body, style="App.TFrame")
        right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        metrics = ttk.Frame(right, style="App.TFrame")
        metrics.pack(fill=tk.BOTH, expand=True)
        volume_panel = ttk.Frame(metrics, style="Panel.TFrame", padding=(12, 12))
        volume_panel.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        ttk.Label(volume_panel, text="Incoming Volume by Sheet", style="Panel.TLabel").pack(anchor=tk.W)
        self.incoming_volume_tree = self._build_home_tree(
            volume_panel,
            columns=("sheet", "person", "cards", "received", "volume", "status"),
            headings={"sheet": "Sheet", "person": "Person", "cards": "Cards", "received": "Received", "volume": "Price Volume", "status": "Status"},
            widths={"sheet": 320, "person": 130, "cards": 80, "received": 95, "volume": 130, "status": 150},
            height=9,
        )
        self.incoming_volume_tree.tag_configure("total_divider", background="#1f1f1f", foreground="#ffffff", font=("Segoe UI Semibold", 10))
        self.incoming_volume_tree.tag_configure("total_row", background="#242424", foreground="#ffffff", font=("Segoe UI Semibold", 10))

        partial_panel = ttk.Frame(metrics, style="Panel.TFrame", padding=(12, 12))
        partial_panel.pack(fill=tk.BOTH, expand=True)
        ttk.Label(partial_panel, text="Partially Received Incoming Sheets", style="Panel.TLabel").pack(anchor=tk.W)
        self.partial_received_tree = self._build_home_tree(
            partial_panel,
            columns=("sheet", "progress", "volume", "person", "tracking", "all_received"),
            headings={"sheet": "Sheet", "progress": "Received", "volume": "Price Volume", "person": "Person", "tracking": "Tracking", "all_received": "All Received"},
            widths={"sheet": 280, "progress": 100, "volume": 130, "person": 130, "tracking": 180, "all_received": 110},
            height=8,
        )
        self.partial_received_tree.tag_configure("partial_sheet", background="#4a3d12", foreground="#fff3b0")

    def _build_home_tree(
        self,
        parent: ttk.Frame,
        columns: tuple[str, ...],
        headings: dict[str, str],
        widths: dict[str, int],
        height: int,
        scrollbars: bool = False,
    ) -> ttk.Treeview:
        container = ttk.Frame(parent, style="Panel.TFrame") if scrollbars else parent
        tree = ttk.Treeview(container, columns=columns, show="headings", selectmode="browse", height=height)
        for col in columns:
            tree.heading(col, text=headings[col], anchor=tk.W)
            tree.column(col, width=widths[col], minwidth=60, stretch=col == "sheet", anchor=tk.W)
        if scrollbars:
            container.pack(fill=tk.BOTH, expand=True, pady=(8, 0))
            tree.grid(row=0, column=0, sticky="nsew")
            y_scroll = ttk.Scrollbar(container, orient=tk.VERTICAL, command=tree.yview)
            y_scroll.grid(row=0, column=1, sticky="ns")
            x_scroll = ttk.Scrollbar(container, orient=tk.HORIZONTAL, command=tree.xview)
            x_scroll.grid(row=1, column=0, sticky="ew")
            tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)
            container.columnconfigure(0, weight=1)
            container.rowconfigure(0, weight=1)
        else:
            tree.pack(fill=tk.BOTH, expand=True, pady=(8, 0))
        return tree

    def _build_payouts_tab(self) -> None:
        controls = ttk.Frame(self.payouts_tab, style="Panel.TFrame", padding=(16, 12))
        controls.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(controls, text="Filter by Assigned Person", style="Panel.TLabel").grid(row=0, column=0, sticky="w")
        self.payout_person_combo = ttk.Combobox(controls, textvariable=self.payout_person_var, width=30)
        self.payout_person_combo.grid(row=0, column=1, sticky="w", padx=(8, 10))
        self._bind_person_autocomplete(self.payout_person_combo, refresh_callback=self.refresh_payouts_tab)
        self.payout_person_combo.bind("<<ComboboxSelected>>", lambda _event: self.refresh_payouts_tab(), add="+")
        controls.columnconfigure(2, weight=1)
        ttk.Label(controls, textvariable=self.payout_status_var, style="Muted.TLabel").grid(row=1, column=0, columnspan=3, sticky="w", pady=(10, 0))

        body = ttk.Frame(self.payouts_tab, style="App.TFrame")
        body.pack(fill=tk.BOTH, expand=True)
        summary_panel = ttk.Frame(body, style="Panel.TFrame", padding=(12, 12))
        summary_panel.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))
        ttk.Label(summary_panel, text="Active Balances", style="Panel.TLabel").pack(anchor=tk.W)
        self.payout_summary_tree = self._build_home_tree(
            summary_panel,
            columns=("person", "sheets", "cards", "balance"),
            headings={"person": "Person", "sheets": "Sheets", "cards": "Cards", "balance": "Balance Owed"},
            widths={"person": 220, "sheets": 80, "cards": 80, "balance": 130},
            height=18,
        )
        self.payout_summary_tree.tag_configure("total_divider", background="#1f1f1f", foreground="#ffffff", font=("Segoe UI Semibold", 10))
        self.payout_summary_tree.tag_configure("total_row", background="#242424", foreground="#ffffff", font=("Segoe UI Semibold", 10))
        self.payout_summary_tree.bind("<ButtonRelease-1>", self.mark_payout_person_paid)

        detail_panel = ttk.Frame(body, style="Panel.TFrame", padding=(12, 12))
        detail_panel.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        ttk.Label(detail_panel, text="Payment Sheets", style="Panel.TLabel").pack(anchor=tk.W)
        self.payout_detail_tree = self._build_home_tree(
            detail_panel,
            columns=("sheet", "stage", "person", "cards", "received", "volume", "status"),
            headings={"sheet": "Sheet", "stage": "Stage", "person": "Person", "cards": "Cards", "received": "Received", "volume": "Balance", "status": "Status"},
            widths={"sheet": 280, "stage": 90, "person": 150, "cards": 80, "received": 95, "volume": 130, "status": 140},
            height=18,
        )
        self.payout_detail_tree.configure(selectmode="extended")
        self.payout_detail_tree.bind("<ButtonRelease-1>", self.open_payout_marker_editor)

    def _build_inventory_tab(self) -> None:
        controls = ttk.Frame(self.inventory_tab, style="Panel.TFrame", padding=(16, 12))
        controls.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(controls, text="Inventory", style="Panel.TLabel", font=("Segoe UI Semibold", 13)).grid(row=0, column=0, sticky="w")
        ttk.Label(controls, text="Person", style="Muted.TLabel").grid(row=0, column=1, sticky="e", padx=(18, 6))
        self.inventory_person_combo = ttk.Combobox(controls, textvariable=self.inventory_person_var, width=22)
        self.inventory_person_combo.grid(row=0, column=2, sticky="w")
        self._bind_person_autocomplete(self.inventory_person_combo, refresh_callback=self.refresh_inventory_tab)
        self.inventory_person_combo.bind("<<ComboboxSelected>>", lambda _event: self.refresh_inventory_tab(), add="+")
        ttk.Label(controls, text="Sport", style="Muted.TLabel").grid(row=0, column=3, sticky="e", padx=(14, 6))
        sport_combo = ttk.Combobox(controls, textvariable=self.inventory_sport_var, values=ASSIGNMENT_CATEGORY_OPTIONS, width=14)
        sport_combo.grid(row=0, column=4, sticky="w")
        sport_combo.bind("<<ComboboxSelected>>", lambda _event: self.refresh_inventory_tab(), add="+")
        ttk.Label(controls, text="Min", style="Muted.TLabel").grid(row=0, column=5, sticky="e", padx=(14, 6))
        ttk.Entry(controls, textvariable=self.inventory_min_var, width=9).grid(row=0, column=6, sticky="w")
        ttk.Label(controls, text="Max", style="Muted.TLabel").grid(row=0, column=7, sticky="e", padx=(10, 6))
        ttk.Entry(controls, textvariable=self.inventory_max_var, width=9).grid(row=0, column=8, sticky="w")
        controls.columnconfigure(9, weight=1)
        ttk.Label(controls, textvariable=self.inventory_metric_var, style="Panel.TLabel").grid(row=0, column=9, sticky="e", padx=(18, 0))
        ttk.Label(controls, text="Search Cert/Card", style="Muted.TLabel").grid(row=1, column=0, sticky="w", pady=(10, 0))
        ttk.Entry(controls, textvariable=self.inventory_search_var, width=42).grid(row=1, column=1, columnspan=4, sticky="w", padx=(8, 0), pady=(10, 0))
        action_row = ttk.Frame(controls, style="Panel.TFrame")
        action_row.grid(row=2, column=0, columnspan=11, sticky="w", pady=(10, 0))
        ttk.Button(action_row, text="Refresh", command=lambda: self.refresh_inventory_tab(enrich=True, filtered_only=True), style="Soft.TButton").pack(side=tk.LEFT)
        ttk.Button(action_row, text="Export", command=self.export_inventory, style="Primary.TButton").pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(action_row, text="Reconcile Received", command=self.reconcile_received_inventory, style="Soft.TButton").pack(side=tk.LEFT, padx=(8, 0))
        ttk.Label(controls, textvariable=self.inventory_status_var, style="Muted.TLabel").grid(row=3, column=0, columnspan=11, sticky="w", pady=(8, 0))
        for var in (self.inventory_sport_var, self.inventory_search_var, self.inventory_min_var, self.inventory_max_var):
            var.trace_add("write", lambda *_args: self.refresh_inventory_tab())

        self.inventory_tree = self._build_home_tree(
            self.inventory_tab,
            columns=INVENTORY_TABLE_COLUMNS,
            headings=INVENTORY_HEADINGS,
            widths=INVENTORY_COLUMN_WIDTHS,
            height=22,
            scrollbars=True,
        )
        self.inventory_tree.configure(selectmode="extended")
        self.inventory_tree.bind("<Button-3>", self._show_inventory_context_menu)
        self.inventory_tree.bind("<Button-2>", self._show_inventory_context_menu)
        self.refresh_inventory_tab()

    def _build_profit_tab(self) -> None:
        controls = ttk.Frame(self.profit_tab, style="Panel.TFrame", padding=(16, 12))
        controls.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(controls, text="Profit", style="Panel.TLabel", font=("Segoe UI Semibold", 13)).grid(row=0, column=0, sticky="w")
        ttk.Label(controls, text="Person", style="Muted.TLabel").grid(row=0, column=1, sticky="e", padx=(18, 6))
        self.profit_person_combo = ttk.Combobox(controls, textvariable=self.profit_person_var, width=28)
        self.profit_person_combo.grid(row=0, column=2, sticky="w")
        self._bind_person_autocomplete(self.profit_person_combo, refresh_callback=self.refresh_profit_tab)
        self.profit_person_combo.bind("<<ComboboxSelected>>", lambda _event: self.refresh_profit_tab(), add="+")
        ttk.Label(controls, text="Period", style="Muted.TLabel").grid(row=0, column=3, sticky="e", padx=(18, 6))
        self.profit_period_combo = ttk.Combobox(
            controls,
            textvariable=self.profit_period_var,
            values=PROFIT_PERIOD_OPTIONS,
            width=10,
            state="readonly",
        )
        self.profit_period_combo.grid(row=0, column=4, sticky="w")
        self.profit_period_combo.bind("<<ComboboxSelected>>", lambda _event: self.refresh_profit_tab(), add="+")
        ttk.Label(controls, text="Graph", style="Muted.TLabel").grid(row=0, column=5, sticky="e", padx=(18, 6))
        self.profit_graph_combo = ttk.Combobox(
            controls,
            textvariable=self.profit_graph_var,
            values=PROFIT_GRAPH_OPTIONS,
            width=14,
            state="readonly",
        )
        self.profit_graph_combo.grid(row=0, column=6, sticky="w")
        self.profit_graph_combo.bind("<<ComboboxSelected>>", lambda _event: self._draw_profit_chart(), add="+")
        ttk.Button(controls, text="Refresh", command=self.refresh_profit_tab, style="Soft.TButton").grid(row=0, column=7, sticky="w", padx=(10, 0))
        ttk.Button(controls, text="Add Expense", command=self.open_add_expense_popup, style="Soft.TButton").grid(row=0, column=8, sticky="w", padx=(8, 0))
        controls.columnconfigure(9, weight=1)
        ttk.Label(controls, textvariable=self.profit_metric_var, style="Panel.TLabel").grid(row=0, column=9, sticky="e")
        ttk.Label(controls, textvariable=self.profit_status_var, style="Muted.TLabel").grid(row=1, column=0, columnspan=10, sticky="w", pady=(8, 0))

        chart_panel = ttk.Frame(self.profit_tab, style="Panel.TFrame", padding=(12, 12))
        chart_panel.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(chart_panel, text="Daily Profit Over Time", style="Panel.TLabel").pack(anchor=tk.W)
        self.profit_chart_canvas = tk.Canvas(
            chart_panel,
            height=230,
            bg="#1f1f1f",
            highlightthickness=1,
            highlightbackground="#333333",
        )
        self.profit_chart_canvas.pack(fill=tk.X, expand=False, pady=(8, 0))
        self.profit_chart_canvas.bind("<Configure>", lambda _event: self._draw_profit_chart())

        ledger_panel = ttk.Frame(self.profit_tab, style="Panel.TFrame", padding=(12, 12))
        ledger_panel.pack(fill=tk.BOTH, expand=True)
        view_row = ttk.Frame(ledger_panel, style="Panel.TFrame")
        view_row.pack(anchor=tk.W, pady=(0, 10))
        self.profit_cards_button = ttk.Button(view_row, text="Sold Cards", command=lambda: self._set_profit_view_mode("Sold Cards"), style="Soft.TButton")
        self.profit_cards_button.pack(side=tk.LEFT)
        self.profit_sheets_button = ttk.Button(view_row, text="Sold Sheets", command=lambda: self._set_profit_view_mode("Sold Sheets"), style="Soft.TButton")
        self.profit_sheets_button.pack(side=tk.LEFT, padx=(8, 0))
        self.profit_expenses_button = ttk.Button(view_row, text="Expenses", command=lambda: self._set_profit_view_mode("Expenses"), style="Soft.TButton")
        self.profit_expenses_button.pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(view_row, text="Refund Selected", command=self.refund_selected_profit_to_inventory, style="Soft.TButton").pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(view_row, text="Delete Expense", command=self.delete_selected_profit_expenses, style="Soft.TButton").pack(side=tk.LEFT, padx=(8, 0))
        self.profit_table_title_var = tk.StringVar(value="Sold Cards")
        ttk.Label(ledger_panel, textvariable=self.profit_table_title_var, style="Panel.TLabel").pack(anchor=tk.W)
        self.profit_tree = self._build_home_tree(
            ledger_panel,
            columns=("date", "company", "card", "cert", "purchase", "sale", "profit", "sheet"),
            headings={
                "date": "Date",
                "company": "Company",
                "card": "Card",
                "cert": "Cert",
                "purchase": "Purchase",
                "sale": "Sale Price",
                "profit": "Profit",
                "sheet": "Company Sheet",
            },
            widths={"date": 95, "company": 150, "card": 440, "cert": 110, "purchase": 105, "sale": 105, "profit": 105, "sheet": 220},
            height=18,
        )
        self.profit_tree.tag_configure("profit_positive", foreground="#d7fbe8")
        self.profit_tree.tag_configure("profit_negative", foreground="#ffd1d1")
        self.profit_tree.tag_configure("total_row", background="#242424", foreground="#ffffff", font=("Segoe UI Semibold", 10))

    def _load_profit_ledger(self) -> list[dict[str, object]]:
        if not PROFIT_LEDGER_PATH.exists():
            return []
        try:
            raw = json.loads(PROFIT_LEDGER_PATH.read_text(encoding="utf-8"))
        except Exception:
            return []
        if not isinstance(raw, list):
            return []
        return [item for item in raw if isinstance(item, dict)]

    def _save_profit_ledger(self, rows: list[dict[str, object]]) -> None:
        atomic_write_json(PROFIT_LEDGER_PATH, rows)

    def _load_inventory_ledger(self) -> list[dict[str, object]]:
        if not INVENTORY_LEDGER_PATH.exists():
            return []
        try:
            raw = json.loads(INVENTORY_LEDGER_PATH.read_text(encoding="utf-8"))
        except Exception:
            return []
        entries = raw.get("items", raw) if isinstance(raw, dict) else raw
        return [item for item in entries if isinstance(item, dict)] if isinstance(entries, list) else []

    def _save_inventory_ledger(self, rows: list[dict[str, object]]) -> None:
        INVENTORY_LEDGER_PATH.parent.mkdir(parents=True, exist_ok=True)
        atomic_write_json(INVENTORY_LEDGER_PATH, {"items": rows})

    def _inventory_record_key(self, record: dict[str, object]) -> str:
        return "|".join(
            str(record.get(field) or "").strip().lower()
            for field in ("cert_number", "source_sheet", "assigned_person")
        )

    def _normalize_inventory_record(self, record: dict[str, object]) -> dict[str, object]:
        normalized = dict(record)
        normalized["date_added"] = str(normalized.get("date_added") or datetime.now().strftime("%Y-%m-%d"))[:10]
        normalized["assigned_person"] = str(normalized.get("assigned_person") or normalized.get("person") or "").strip() or "Unassigned"
        normalized["sport"] = str(normalized.get("sport") or "").strip()
        normalized["cert_number"] = str(normalized.get("cert_number") or "").strip()
        normalized["grader"] = str(normalized.get("grader") or "").strip()
        normalized["card_title"] = str(normalized.get("card_title") or "").strip()
        normalized["purchase_price"] = self._money_value(normalized.get("purchase_price"))
        normalized["card_ladder_value"] = self._money_value(normalized.get("card_ladder_value"))
        normalized["card_ladder_comps_average"] = self._money_value(normalized.get("card_ladder_comps_average") or normalized.get("comps"))
        normalized["cy_value"] = self._money_value(normalized.get("cy_value") or normalized.get("cy_estimate"))
        normalized["inventory_value"] = self._money_value(normalized.get("inventory_value") or normalized.get("value") or normalized.get("sale_price") or normalized.get("estimated_payout"))
        normalized["best_company"] = str(normalized.get("best_company") or normalized.get("company") or "").strip()
        normalized["estimated_payout"] = self._money_value(normalized.get("estimated_payout") or normalized.get("payout"))
        normalized["source_sheet"] = str(normalized.get("source_sheet") or "").strip()
        normalized["source"] = str(normalized.get("source") or "").strip()
        normalized["status"] = str(normalized.get("status") or "Active").strip() or "Active"
        normalized["notes"] = str(normalized.get("notes") or "").strip()
        normalized["inventory_key"] = str(normalized.get("inventory_key") or self._inventory_record_key(normalized))
        return normalized

    def _inventory_record_from_row(self, row: WorkbookRow, person: str, source_sheet: str = "", source: str = "", status: str = "Active", notes: str = "") -> dict[str, object]:
        card_title = str(row.card_title or "")
        sport = assignment_engine.parse_card_for_matching(card_title).get("sport") if card_title else ""
        return self._normalize_inventory_record(
            {
                "date_added": datetime.now().strftime("%Y-%m-%d"),
                "assigned_person": person or "Unassigned",
                "sport": sport,
                "cert_number": row.cert_number,
                "grader": row.grader,
                "card_title": row.card_title,
                "purchase_price": row.existing_value,
                "card_ladder_value": row.card_ladder_value,
                "card_ladder_comps_average": row.card_ladder_comps_average,
                "cy_value": row.cy_value,
                "inventory_value": row.card_ladder_comps_average or row.card_ladder_value or row.cy_value,
                "best_company": row.best_company,
                "estimated_payout": row.estimated_payout,
                "source_sheet": source_sheet,
                "source": source,
                "status": status,
                "notes": notes,
            }
        )

    def add_inventory_records(self, records: list[dict[str, object]], refresh: bool = True) -> int:
        if not records:
            return 0
        ledger = [self._normalize_inventory_record(record) for record in self._load_inventory_ledger()]
        by_key = {str(record.get("inventory_key") or ""): record for record in ledger}
        added = 0
        for record in records:
            normalized = self._normalize_inventory_record(record)
            normalized = self._enrich_inventory_record_assignment(normalized)
            key = str(normalized.get("inventory_key") or "")
            if not key:
                continue
            if key not in by_key:
                ledger.append(normalized)
                by_key[key] = normalized
                added += 1
            else:
                existing = by_key[key]
                existing.update(normalized)
                existing["status"] = "Active"
        self._save_inventory_ledger(ledger)
        if refresh:
            self.refresh_inventory_tab()
        return added

    def _mobile_app_url(self) -> str:
        host = "127.0.0.1"
        try:
            with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as handle:
                handle.connect(("8.8.8.8", 80))
                host = handle.getsockname()[0]
        except OSError:
            try:
                host = socket.gethostbyname(socket.gethostname())
            except OSError:
                host = "127.0.0.1"
        return f"http://{host}:{self.bridge.port}/mobile"

    def _mobile_inventory_payload_record(self, payload: dict) -> dict[str, object]:
        cert = scan_to_cert(payload.get("cert_number") or payload.get("cert") or payload.get("barcode") or "")
        card_title = str(payload.get("card_title") or payload.get("card") or "").strip()
        source = str(payload.get("source") or payload.get("seller") or "Mobile").strip() or "Mobile"
        person = str(payload.get("assigned_person") or payload.get("person") or "").strip() or "Unassigned"
        sport = str(payload.get("sport") or "").strip()
        if not sport and card_title:
            sport = str(assignment_engine.parse_card_for_matching(card_title).get("sport") or "")
        return self._normalize_inventory_record(
            {
                "date_added": datetime.now().strftime("%Y-%m-%d"),
                "assigned_person": person,
                "sport": sport,
                "cert_number": cert,
                "grader": normalize_grader(payload.get("grader") or ""),
                "card_title": card_title,
                "purchase_price": payload.get("purchase_price") or payload.get("purchase") or payload.get("price_paid"),
                "inventory_value": payload.get("inventory_value") or payload.get("value"),
                "source_sheet": str(payload.get("source_sheet") or "Mobile Inventory").strip() or "Mobile Inventory",
                "source": source,
                "status": "Active",
                "notes": str(payload.get("notes") or "").strip(),
            }
        )

    def _mobile_inventory_json_record(self, record: dict[str, object]) -> dict[str, object]:
        normalized = self._normalize_inventory_record(record)
        return {
            "inventory_key": normalized.get("inventory_key"),
            "date_added": normalized.get("date_added"),
            "assigned_person": normalized.get("assigned_person"),
            "sport": normalized.get("sport"),
            "cert_number": normalized.get("cert_number"),
            "grader": normalized.get("grader"),
            "card_title": normalized.get("card_title"),
            "purchase_price": normalized.get("purchase_price"),
            "purchase_price_display": format_money(self._money_value(normalized.get("purchase_price"))),
            "inventory_value": normalized.get("inventory_value"),
            "inventory_value_display": format_money(self._money_value(normalized.get("inventory_value"))),
            "best_company": normalized.get("best_company"),
            "estimated_payout": normalized.get("estimated_payout"),
            "estimated_payout_display": format_money(self._money_value(normalized.get("estimated_payout"))),
            "source_sheet": normalized.get("source_sheet"),
            "source": normalized.get("source"),
            "status": normalized.get("status"),
            "notes": normalized.get("notes"),
        }

    def mobile_inventory_search(self, payload: dict) -> dict:
        query = str(payload.get("query") or payload.get("q") or "").strip().lower()
        person = str(payload.get("person") or "").strip().lower()
        include_sold = bool(payload.get("include_sold"))
        rows = [self._normalize_inventory_record(record) for record in self._load_inventory_ledger()]
        results: list[dict[str, object]] = []
        for record in rows:
            status = str(record.get("status") or "").lower()
            if status != "active" and not include_sold:
                continue
            if person and person not in str(record.get("assigned_person") or "Unassigned").lower():
                continue
            haystack = " ".join(
                str(record.get(field) or "")
                for field in ("cert_number", "card_title", "grader", "assigned_person", "sport", "source", "best_company", "notes")
            ).lower()
            if query and any(part not in haystack for part in query.split()):
                continue
            results.append(self._mobile_inventory_json_record(record))
            if len(results) >= 75:
                break
        return {"ok": True, "count": len(results), "items": results, "people": self._known_people()}

    def mobile_inventory_add(self, payload: dict) -> dict:
        record = self._mobile_inventory_payload_record(payload)
        if not record.get("cert_number") and not record.get("card_title"):
            return {"ok": False, "error": "Enter or scan a cert number, or enter a card title."}
        cert = scan_to_cert(record.get("cert_number"))
        update_existing = bool(payload.get("update_existing"))
        with shared_lock(CARD_PIPELINE_DIR, "mobile-inventory", self.lucas_identity):
            ledger = [self._normalize_inventory_record(item) for item in self._load_inventory_ledger()]
            existing_index = next(
                (
                    index
                    for index, item in enumerate(ledger)
                    if cert and scan_to_cert(item.get("cert_number")) == cert and str(item.get("status") or "").lower() == "active"
                ),
                None,
            )
            if existing_index is not None and not update_existing:
                return {
                    "ok": False,
                    "duplicate": True,
                    "error": "That cert is already active in inventory.",
                    "record": self._mobile_inventory_json_record(ledger[existing_index]),
                }
            if existing_index is not None:
                existing = ledger[existing_index]
                for key, value in record.items():
                    if key in {"inventory_key", "date_added"}:
                        continue
                    if value not in ("", None):
                        existing[key] = value
                existing["status"] = "Active"
                ledger[existing_index] = self._enrich_inventory_record_assignment(self._normalize_inventory_record(existing), force=True)
                self._save_inventory_ledger(ledger)
                saved = ledger[existing_index]
                action = "updated"
            else:
                saved = self._enrich_inventory_record_assignment(record)
                ledger.append(saved)
                self._save_inventory_ledger(ledger)
                action = "added"
        self.events.put(("inventory_refresh", f"Mobile inventory {action}: {saved.get('cert_number') or saved.get('card_title') or 'card'}"))
        return {"ok": True, "action": action, "record": self._mobile_inventory_json_record(saved)}

    def mobile_inventory_mark_sold(self, payload: dict) -> dict:
        inventory_key = str(payload.get("inventory_key") or payload.get("key") or "").strip()
        if not inventory_key:
            return {"ok": False, "error": "Choose an inventory card to mark sold."}
        sale_price = self._money_value(payload.get("sale_price") or payload.get("amount") or payload.get("price"))
        if sale_price is None or sale_price < 0:
            return {"ok": False, "error": "Enter a valid sale price."}
        sale_date = str(payload.get("sale_date") or payload.get("date") or "").strip() or datetime.now().strftime("%Y-%m-%d")
        if self._profit_record_date(sale_date) is None:
            return {"ok": False, "error": "Enter the sale date as YYYY-MM-DD."}
        sale_method = str(payload.get("sale_method") or payload.get("method") or "").strip()
        company = str(payload.get("company") or payload.get("buyer") or "").strip()
        with shared_lock(CARD_PIPELINE_DIR, "mobile-inventory-sold", self.lucas_identity):
            ledger = [self._normalize_inventory_record(record) for record in self._load_inventory_ledger()]
            record = next((item for item in ledger if str(item.get("inventory_key") or "") == inventory_key), None)
            if record is None:
                return {"ok": False, "error": "That inventory card was not found."}
            if str(record.get("status") or "").lower() != "active":
                return {"ok": False, "error": "Only active inventory cards can be marked sold."}
            profit_record = self._inventory_sale_profit_record(record, company, float(sale_price), sale_date=sale_date, sale_method=sale_method)
            added = self._append_profit_records([profit_record])
            changed = self._mark_inventory_record_sold(inventory_key, company or "General Sold", float(sale_price))
        if not (added or changed):
            return {"ok": False, "error": "That sale already exists."}
        title = record.get("cert_number") or record.get("card_title") or "card"
        self.events.put(("inventory_refresh", f"Mobile marked sold: {title} for {format_money(sale_price)}."))
        self.events.put(("profit_refresh", f"Mobile marked sold: {title} for {format_money(sale_price)}."))
        return {
            "ok": True,
            "record": self._mobile_inventory_json_record(record),
            "sale": {
                "date": sale_date[:10],
                "method": sale_method,
                "company": company or "General Sold",
                "sale_price": round(float(sale_price), 2),
                "sale_price_display": format_money(sale_price),
                "profit": profit_record.get("profit"),
                "profit_display": format_money(profit_record.get("profit")),
            },
            "people": self._known_people(),
        }

    def _mobile_profit_rows(self, person: str = "", period: str = "Total") -> list[dict[str, object]]:
        needle = person.strip().lower()
        period_start, period_end = self._profit_period_bounds(period)
        rows = self._enrich_profit_records_with_people(self._load_profit_ledger())
        filtered: list[dict[str, object]] = []
        for record in rows:
            if needle and needle not in str(record.get("assigned_person") or "Unassigned").lower():
                continue
            if period_start is not None:
                sold_date = self._profit_record_date(record.get("date_added"))
                if sold_date is None or sold_date < period_start or sold_date > period_end:
                    continue
            filtered.append(record)
        return sorted(filtered, key=lambda item: (str(item.get("date_added") or ""), str(item.get("company") or ""), str(item.get("card_title") or "")), reverse=True)

    def _mobile_profit_chart_series(self, rows: list[dict[str, object]], period: str, graph: str) -> tuple[list[str], list[float]]:
        daily: dict[str, float] = {}
        for record in rows:
            profit = self._money_value(record.get("profit"))
            sold_date = self._profit_record_date(record.get("date_added"))
            if profit is None or sold_date is None:
                continue
            day = sold_date.isoformat()
            daily[day] = daily.get(day, 0.0) + float(profit)
        period_start, period_end = self._profit_period_bounds(period)
        if period_start is not None:
            cursor = period_start
            while cursor <= period_end:
                daily.setdefault(cursor.isoformat(), 0.0)
                cursor += timedelta(days=1)
        days = sorted(daily)
        values = [daily[day] for day in days]
        if graph == "Overall Profit":
            running = 0.0
            cumulative: list[float] = []
            for value in values:
                running += value
                cumulative.append(round(running, 2))
            values = cumulative
        return days, [round(value, 2) for value in values]

    def mobile_profit_summary(self, payload: dict) -> dict:
        period = str(payload.get("period") or "Total").strip()
        if period not in PROFIT_PERIOD_OPTIONS:
            period = "Total"
        graph = str(payload.get("graph") or "Daily Trend").strip()
        if graph not in PROFIT_GRAPH_OPTIONS:
            graph = "Daily Trend"
        rows = self._mobile_profit_rows(str(payload.get("person") or ""), period)
        total_purchase = 0.0
        total_sale = 0.0
        gross_profit = 0.0
        expenses = 0.0
        net_profit = 0.0
        complete_count = 0
        recent: list[dict[str, object]] = []
        for record in rows:
            is_expense = str(record.get("record_type") or "").strip().lower() == "expense"
            purchase = self._money_value(record.get("purchase_price"))
            sale = self._money_value(record.get("sale_price"))
            profit = self._money_value(record.get("profit"))
            if purchase is not None:
                total_purchase += purchase
            if sale is not None:
                total_sale += sale
            if profit is not None:
                net_profit += profit
                if is_expense:
                    expenses += abs(profit)
                else:
                    gross_profit += profit
                complete_count += 1
            if len(recent) < 25:
                recent.append(
                    {
                        "date": record.get("date_added") or "",
                        "person": record.get("assigned_person") or "Unassigned",
                        "type": "Expense" if is_expense else "Sale",
                        "title": record.get("card_title") or record.get("company") or "",
                        "company": record.get("company") or "",
                        "profit": round(profit or 0.0, 2) if profit is not None else None,
                        "profit_display": format_money(profit),
                    }
                )
        labels, values = self._mobile_profit_chart_series(rows, period, graph)
        return {
            "ok": True,
            "people": self._known_people(),
            "periods": list(PROFIT_PERIOD_OPTIONS),
            "graphs": list(PROFIT_GRAPH_OPTIONS),
            "totals": {
                "purchase": round(total_purchase, 2),
                "sale": round(total_sale, 2),
                "gross_profit": round(gross_profit, 2),
                "expenses": round(expenses, 2),
                "net_profit": round(net_profit, 2),
                "complete_count": complete_count,
                "row_count": len(rows),
            },
            "chart": {"labels": labels, "values": values},
            "recent": recent,
        }

    def mobile_expense_add(self, payload: dict) -> dict:
        person = str(payload.get("person") or payload.get("assigned_person") or "").strip()
        if not person:
            return {"ok": False, "error": "Choose the person this expense belongs to."}
        expense_date = str(payload.get("date") or payload.get("date_added") or "").strip() or datetime.now().strftime("%Y-%m-%d")
        if self._profit_record_date(expense_date) is None:
            return {"ok": False, "error": "Enter the expense date as YYYY-MM-DD."}
        amount = self._money_value(payload.get("amount") or payload.get("expense_amount"))
        if amount is None or amount <= 0:
            return {"ok": False, "error": "Enter an expense amount greater than zero."}
        expense_type = str(payload.get("expense_type") or payload.get("type") or "").strip()
        if expense_type not in EXPENSE_CATEGORY_OPTIONS:
            expense_type = "Fees"
        related_type = str(payload.get("related_type") or payload.get("tie_to") or "").strip()
        if related_type not in EXPENSE_LINK_OPTIONS:
            related_type = "General"
        related_sheet = str(payload.get("source_sheet") or payload.get("sheet") or "").strip()
        related_cert = str(payload.get("cert_number") or payload.get("cert") or "").strip()
        if related_type == "Sheet" and not related_sheet:
            return {"ok": False, "error": "Enter the sold sheet this expense belongs to."}
        if related_type == "Card" and not (related_sheet or related_cert):
            return {"ok": False, "error": "Enter a cert number or sold sheet for the card expense."}
        record = {
            "record_type": "expense",
            "expense_id": datetime.now().strftime("%Y%m%d%H%M%S%f"),
            "date_added": expense_date[:10],
            "assigned_person": person,
            "expense_type": expense_type,
            "expense_amount": amount,
            "related_type": related_type,
            "source_sheet": related_sheet,
            "cert_number": related_cert,
            "notes": str(payload.get("notes") or "").strip(),
        }
        added = self._append_profit_records([record])
        if not added:
            return {"ok": False, "error": "That expense already exists in the profit ledger."}
        self.events.put(("profit_refresh", f"Added {expense_type} expense for {person}: {format_money(amount)}."))
        return {"ok": True, "record": self._normalize_profit_record(record), "people": self._known_people()}

    def mobile_payouts(self, payload: dict) -> dict:
        needle = str(payload.get("person") or "").strip().lower()
        balances: dict[str, dict[str, float | int]] = {}
        details: list[dict[str, object]] = []
        for item in self._payout_sheet_items():
            person = str(item.get("person") or "Unassigned")
            if needle and needle not in person.lower():
                continue
            if not item.get("paid"):
                balance = balances.setdefault(person, {"sheets": 0, "cards": 0, "balance": 0.0})
                balance["sheets"] = int(balance["sheets"]) + 1
                balance["cards"] = int(balance["cards"]) + int(item.get("row_count") or 0)
                balance["balance"] = float(balance["balance"]) + float(item.get("payout_balance") or 0.0)
            details.append(
                {
                    "name": item.get("name") or "",
                    "stage": item.get("stage") or "",
                    "person": person,
                    "row_count": int(item.get("row_count") or 0),
                    "received_count": int(item.get("received_count") or 0),
                    "payout_balance": round(float(item.get("payout_balance") or 0.0), 2),
                    "payout_balance_display": format_money(float(item.get("payout_balance") or 0.0)),
                    "status": item.get("status") or "",
                    "paid": bool(item.get("paid")),
                }
            )
        summary = [
            {
                "person": person,
                "sheets": int(values["sheets"]),
                "cards": int(values["cards"]),
                "balance": round(float(values["balance"]), 2),
                "balance_display": format_money(float(values["balance"])),
            }
            for person, values in sorted(balances.items(), key=lambda pair: (-float(pair[1]["balance"]), pair[0].lower()))
        ]
        total_balance = sum(item["balance"] for item in summary)
        return {
            "ok": True,
            "people": self._known_people(),
            "summary": summary,
            "details": details,
            "totals": {
                "balance": round(total_balance, 2),
                "balance_display": format_money(total_balance),
                "sheets": sum(int(item["sheets"]) for item in summary),
                "cards": sum(int(item["cards"]) for item in summary),
            },
        }

    def _mobile_image_parts(self, image: str) -> tuple[str, str, bytes]:
        match = re.match(r"^data:([^;]+);base64,(.*)$", image, re.S)
        if match:
            mime_type = match.group(1) or "image/jpeg"
            image_b64 = match.group(2)
        else:
            mime_type = "image/jpeg"
            image_b64 = image
        return mime_type, image_b64, base64.b64decode(image_b64)

    def _parse_mobile_quick_card_response(self, raw: str) -> dict:
        text = str(raw or "").strip()
        text = re.sub(r"^```(?:json)?\s*|\s*```$", "", text, flags=re.I | re.S).strip()
        try:
            parsed = json.loads(text)
        except Exception:
            match = re.search(r"\{.*\}", text, re.S)
            if not match:
                return {}
            try:
                parsed = json.loads(match.group(0))
            except Exception:
                return {}
        return parsed if isinstance(parsed, dict) else {}

    def _mobile_quick_card_to_row(self, parsed: dict) -> dict[str, object]:
        grader = normalize_grader(parsed.get("grading_company") or parsed.get("grader") or "")
        title = build_card_title(
            {
                "description": "",
                "year": parsed.get("year"),
                "set": parsed.get("set"),
                "player": parsed.get("player") or parsed.get("subject"),
                "card_number": parsed.get("card_number"),
                "parallel": parsed.get("parallel"),
                "subset": parsed.get("subset") or parsed.get("attributes"),
                "grader": grader,
                "grade": parsed.get("grade"),
            }
        )
        label_text = str(parsed.get("label_text") or "").strip()
        if not title:
            title = str(parsed.get("card_title") or parsed.get("title") or "").strip()
        notes = clean_part("; ".join(part for part in ("Mobile quick scan", label_text[:180]) if part))
        return {
            "cert_number": scan_to_cert(parsed.get("cert_number")),
            "grader": grader or infer_grader(title),
            "card_title": title,
            "purchase_price": None,
            "source": "Mobile Photo",
            "notes": notes,
        }

    def _mobile_single_card_quick_read(self, client, mime_type: str, image_bytes: bytes) -> dict[str, object] | None:
        if genai_types is None:
            return None
        prompt = (
            "Read this single trading card or graded slab photo for inventory entry. "
            "Assume the user is photographing one card/slab. Extract only visible facts; do not guess. "
            "Return JSON only with keys: grading_company, cert_number, player, year, set, card_number, "
            "parallel, subset, attributes, grade, card_title, label_text, confidence. "
            "Normalize cert_number to digits only when possible. If a field is unreadable, use an empty string. "
            "For BGS/Beckett, grade must be the overall slab grade, not a subgrade."
        )
        response = client.models.generate_content(
            model="gemini-2.5-flash",
            contents=[
                prompt,
                genai_types.Part.from_bytes(data=image_bytes, mime_type=mime_type),
            ],
            config=genai_types.GenerateContentConfig(
                thinking_config=genai_types.ThinkingConfig(thinking_budget=0),
                max_output_tokens=700,
                response_mime_type="application/json",
                temperature=0,
            ),
        )
        parsed = self._parse_mobile_quick_card_response(response.text or "")
        row = self._mobile_quick_card_to_row(parsed)
        if row.get("cert_number") or row.get("card_title") or row.get("grader"):
            return row
        return None

    def mobile_card_identify(self, payload: dict) -> dict:
        image = str(payload.get("image") or "").strip()
        if not image:
            return {"ok": False, "error": "Take or choose a card photo first."}
        if genai is None or identify_cards_sync is None:
            return {"ok": False, "error": "Photo OCR dependencies are not available."}
        self._load_photo_env()
        api_key = os.environ.get("GOOGLE_API_KEY", "").strip()
        if not api_key:
            return {"ok": False, "error": "Missing GOOGLE_API_KEY for photo card search."}
        try:
            mime_type, image_b64, image_bytes = self._mobile_image_parts(image)
        except Exception as error:
            return {"ok": False, "error": f"Could not read that photo: {error}"}
        try:
            client = genai.Client(api_key=api_key)
            row = self._mobile_single_card_quick_read(client, mime_type, image_bytes)
            if row is None:
                cards = identify_cards_sync(client, image_b64)
                rows = [
                    self._photo_card_to_row(Path("mobile-photo.jpg"), card)
                    for card in cards
                    if self._photo_card_has_inventory(card)
                ]
                if not rows:
                    return {"ok": False, "error": "No card was found in that photo."}
                row = rows[0]
                cards_found = len(rows)
                mode = "fallback"
            else:
                cards_found = 1
                mode = "quick"
        except (TemporaryModelUnavailable, ModelQuotaExceeded, ModelResponseParseError) as error:
            return {"ok": False, "error": str(error)}
        except Exception as error:
            return {"ok": False, "error": f"Photo search failed: {error}"}
        query = scan_to_cert(row.get("cert_number")) or str(row.get("card_title") or "").strip()
        return {
            "ok": True,
            "query": query,
            "card": {
                "cert_number": row.get("cert_number"),
                "grader": row.get("grader"),
                "card_title": row.get("card_title"),
                "notes": row.get("notes"),
            },
            "cards_found": cards_found,
            "mode": mode,
        }

    def _retarget_inventory_rows_for_source(self, source_sheet_name: str, assigned_person: str) -> int:
        source_name = Path(str(source_sheet_name or "")).name.strip().lower()
        if not source_name:
            return 0
        new_person = str(assigned_person or "").strip() or "Unassigned"
        changed = 0
        merged: dict[str, dict[str, object]] = {}
        for record in [self._normalize_inventory_record(item) for item in self._load_inventory_ledger()]:
            if Path(str(record.get("source_sheet") or "")).name.strip().lower() == source_name:
                if str(record.get("assigned_person") or "").strip() != new_person:
                    changed += 1
                record["assigned_person"] = new_person
                record["best_company"] = ""
                record["estimated_payout"] = None
                record.pop("inventory_key", None)
                record = self._normalize_inventory_record(record)
                record = self._enrich_inventory_record_assignment(record)
            key = str(record.get("inventory_key") or "")
            if not key:
                continue
            existing = merged.get(key)
            if existing is None:
                merged[key] = record
                continue
            status = "Active" if "active" in {str(existing.get("status") or "").lower(), str(record.get("status") or "").lower()} else str(record.get("status") or existing.get("status") or "")
            existing.update(record)
            existing["status"] = status or "Active"
        if changed:
            self._save_inventory_ledger(list(merged.values()))
        return changed

    def _received_certs_in_workbook(self, path: Path) -> set[str]:
        certs: set[str] = set()
        if not path.exists():
            return certs
        workbook = load_workbook(path, read_only=True, data_only=True)
        try:
            for sheet in workbook.worksheets:
                headers = {
                    re.sub(r"[^a-z0-9]", "", str(cell.value or "").strip().lower()): index
                    for index, cell in enumerate(sheet[1], start=1)
                    if str(cell.value or "").strip()
                }
                received_col = headers.get("received")
                cert_col = headers.get("certificationnumber") or headers.get("certnumber") or headers.get("cert")
                if not received_col or not cert_col:
                    continue
                for row_index in range(2, sheet.max_row + 1):
                    received_text = str(sheet.cell(row_index, received_col).value or "").strip().upper()
                    if received_text not in {"X", "Y", "YES", "TRUE", "1", "RECEIVED"}:
                        continue
                    cert = scan_to_cert(sheet.cell(row_index, cert_col).value)
                    if cert:
                        certs.add(cert)
        finally:
            workbook.close()
        return certs

    def _company_sheet_source_cert_keys(self) -> set[tuple[str, str]]:
        keys: set[tuple[str, str]] = set()
        for record in read_company_profit_records(COMPANY_SHEETS_DIR):
            source_sheet = Path(str(record.get("source_sheet") or "")).name.strip().lower()
            cert = scan_to_cert(record.get("cert_number"))
            if source_sheet and cert:
                keys.add((source_sheet, cert))
        return keys

    def _received_inventory_candidate_records_for_sheet(
        self,
        stage: str,
        path: Path,
        person: str,
        company_keys: set[tuple[str, str]] | None = None,
    ) -> list[dict[str, object]]:
        assigned_person = str(person or "").strip()
        if not assigned_person or not path.exists():
            return []
        received_certs = None if stage == "Received" else self._received_certs_in_workbook(path)
        if received_certs == set():
            return []
        try:
            rows = read_simple_spreadsheet(path)
        except Exception:
            return []
        company_keys = company_keys if company_keys is not None else self._company_sheet_source_cert_keys()
        candidates: list[dict[str, object]] = []
        for row in rows:
            cert = scan_to_cert(row.get("cert_number"))
            if not cert:
                continue
            if received_certs is not None and cert not in received_certs:
                continue
            if (path.name.lower(), cert) in company_keys:
                continue
            card_title = str(row.get("card_title") or "")
            candidates.append(
                self._normalize_inventory_record(
                    {
                        "date_added": datetime.now().strftime("%Y-%m-%d"),
                        "assigned_person": assigned_person,
                        "sport": assignment_engine.parse_card_for_matching(card_title).get("sport") if card_title else "",
                        "cert_number": cert,
                        "grader": row.get("grader") or "",
                        "card_title": card_title,
                        "purchase_price": row.get("purchase_price"),
                        "card_ladder_value": row.get("card_ladder_value"),
                        "card_ladder_comps_average": row.get("card_ladder_comps_average"),
                        "cy_value": row.get("cy_value"),
                        "inventory_value": row.get("card_ladder_comps_average") or row.get("card_ladder_value") or row.get("cy_value"),
                        "best_company": row.get("best_company") or "",
                        "estimated_payout": row.get("estimated_payout"),
                        "source_sheet": path.name,
                        "source": row.get("source") or "",
                        "status": "Active",
                        "notes": "Backfilled from received sheets",
                    }
                )
            )
        return candidates

    def _received_inventory_candidate_records(self) -> list[dict[str, object]]:
        company_keys = self._company_sheet_source_cert_keys()
        candidates: list[dict[str, object]] = []
        for stage, directory in (("Received", RECEIVED_SHEETS_DIR), ("Incoming", INCOMING_SHEETS_DIR), ("Working", WORKING_SHEETS_DIR)):
            if not directory.exists():
                continue
            for path in sorted(directory.glob("*.xlsx"), key=lambda item: item.name.lower()):
                marker = self.home_sheet_markers.get(self._home_sheet_key(stage, path.name), {})
                person = str(marker.get("assigned_person") or "").strip()
                if not person:
                    continue
                candidates.extend(self._received_inventory_candidate_records_for_sheet(stage, path, person, company_keys))
        return candidates

    def reconcile_received_inventory(self) -> None:
        added, candidates = self._sync_received_inventory_to_ledger()
        self.refresh_inventory_tab(enrich=True)
        self.status_var.set(f"Reconciled received inventory: added {added} active card(s) from {candidates} candidate row(s).")
        if added:
            messagebox.showinfo("Inventory reconciled", f"Added {added} received card(s) to active inventory.")
        else:
            messagebox.showinfo("Inventory reconciled", "No missing received inventory cards were found.")

    def _sync_received_inventory_to_ledger(self) -> tuple[int, int]:
        records = self._received_inventory_candidate_records()
        added = self.add_inventory_records(records, refresh=False)
        return added, len(records)

    def _sync_received_sheet_inventory_to_ledger(self, stage: str, path: Path, person: str) -> tuple[int, int]:
        records = self._received_inventory_candidate_records_for_sheet(stage, path, person)
        added = self.add_inventory_records(records, refresh=False)
        return added, len(records)

    def _inventory_workbook_row(self, record: dict[str, object], excel_row: int) -> WorkbookRow:
        inventory_value = self._money_value(record.get("inventory_value"))
        card_ladder_value = self._money_value(record.get("card_ladder_value"))
        comps_average = self._money_value(record.get("card_ladder_comps_average"))
        cy_value = self._money_value(record.get("cy_value"))
        if card_ladder_value is None and comps_average is None and cy_value is None:
            card_ladder_value = inventory_value
            comps_average = inventory_value
        return WorkbookRow(
            excel_row=excel_row,
            cert_number=str(record.get("cert_number") or ""),
            grader=str(record.get("grader") or ""),
            card_title=str(record.get("card_title") or ""),
            existing_value=self._money_value(record.get("purchase_price")),
            card_ladder_value=card_ladder_value,
            card_ladder_comps_average=comps_average,
            cy_value=cy_value,
            best_company=str(record.get("best_company") or ""),
            estimated_payout=self._money_value(record.get("estimated_payout")),
            company_pile=True,
            status="Moved from inventory",
            notes=str(record.get("notes") or "Moved from inventory"),
        )

    def _inventory_source_sheet_path(self, source_sheet: str) -> Path | None:
        name = Path(str(source_sheet or "")).name
        if not name:
            return None
        for directory in (RECEIVED_SHEETS_DIR, WORKING_SHEETS_DIR, INCOMING_SHEETS_DIR):
            path = directory / name
            if path.exists():
                return path
        return None

    def _hydrate_inventory_record_source_values(self, record: dict[str, object]) -> dict[str, object]:
        normalized = self._normalize_inventory_record(record)
        if (
            normalized.get("card_ladder_value") is not None
            and normalized.get("card_ladder_comps_average") is not None
            and normalized.get("cy_value") is not None
        ):
            return normalized
        cert = scan_to_cert(normalized.get("cert_number"))
        path = self._inventory_source_sheet_path(str(normalized.get("source_sheet") or ""))
        if not cert or path is None:
            return normalized
        try:
            rows = read_simple_spreadsheet(path)
        except Exception:
            return normalized
        for row in rows:
            if scan_to_cert(row.get("cert_number")) != cert:
                continue
            for source_field in ("card_ladder_value", "card_ladder_comps_average", "cy_value"):
                if normalized.get(source_field) is None:
                    normalized[source_field] = self._money_value(row.get(source_field))
            return self._normalize_inventory_record(normalized)
        return normalized

    def _enrich_inventory_record_assignment(self, record: dict[str, object], force: bool = False) -> dict[str, object]:
        hydrator = getattr(self, "_hydrate_inventory_record_source_values", None)
        normalized = hydrator(record) if callable(hydrator) else self._normalize_inventory_record(record)
        if not force and normalized.get("best_company") and normalized.get("estimated_payout") is not None:
            return normalized
        try:
            row = self._inventory_workbook_row(normalized, 1)
            recommendation = self.assignment_engine.recommend(row, person=str(normalized.get("assigned_person") or ""))
        except Exception:
            return normalized
        if recommendation.payout is None:
            normalized["best_company"] = normalized.get("best_company") or NO_COMPANY_TAKES_LABEL
            normalized["estimated_payout"] = None
            normalized["inventory_value"] = getattr(recommendation, "source_value", None) or assignment_engine.assignment_value(row)
            return normalized
        normalized["best_company"] = recommendation.company
        normalized["estimated_payout"] = recommendation.payout
        normalized["inventory_value"] = getattr(recommendation, "source_value", None) or assignment_engine.assignment_value(row)
        return normalized

    def _mark_inventory_records_moved_to_company(self, moved_keys: set[str]) -> None:
        if not moved_keys:
            return
        ledger = [self._normalize_inventory_record(record) for record in self._load_inventory_ledger()]
        kept = [record for record in ledger if str(record.get("inventory_key") or "") not in moved_keys]
        if len(kept) != len(ledger):
            self._save_inventory_ledger(kept)

    def _mark_inventory_record_sold(self, inventory_key: str, company: str, sale_price: float) -> int:
        if not inventory_key:
            return 0
        ledger = [self._normalize_inventory_record(record) for record in self._load_inventory_ledger()]
        kept = [record for record in ledger if str(record.get("inventory_key") or "") != inventory_key]
        changed = len(ledger) - len(kept)
        if changed:
            self._save_inventory_ledger(kept)
        return changed

    def _general_sold_sheet_name(self, person: str) -> str:
        return f"{str(person or '').strip() or 'Unassigned'} General Sold"

    def _inventory_sale_profit_record(
        self,
        record: dict[str, object],
        company: str,
        sale_price: float,
        sale_date: str | None = None,
        sale_method: str = "",
    ) -> dict[str, object]:
        normalized = self._normalize_inventory_record(record)
        assigned_person = str(normalized.get("assigned_person") or "Unassigned").strip() or "Unassigned"
        company_name = str(company or "").strip()
        source_sheet = normalized.get("source_sheet") or "Inventory"
        if not company_name:
            company_name = "General Sold"
            source_sheet = self._general_sold_sheet_name(assigned_person)
        sold_date = str(sale_date or "").strip()[:10]
        if CardPipelineApp._profit_record_date(self, sold_date) is None:
            sold_date = datetime.now().strftime("%Y-%m-%d")
        notes = str(normalized.get("notes") or "").strip()
        method = str(sale_method or "").strip()
        if method:
            notes = clean_part("; ".join(part for part in (notes, f"Sale method: {method}") if part))
        return self._normalize_profit_record(
            {
                "date_added": sold_date,
                "company": company_name,
                "weekly_sheet_name": "Inventory Sale",
                "source_sheet": source_sheet,
                "source": normalized.get("source") or "Inventory",
                "cert_number": normalized.get("cert_number") or "",
                "grader": normalized.get("grader") or "",
                "card_title": normalized.get("card_title") or "",
                "purchase_price": normalized.get("purchase_price"),
                "sale_price": sale_price,
                "sale_method": method,
                "assigned_person": assigned_person,
                "status": "Sold from inventory",
                "notes": notes,
            }
        )

    def mark_inventory_record_sold(self, record: dict[str, object], company: str, sale_price: float, sale_date: str | None = None, sale_method: str = "") -> bool:
        normalized = self._normalize_inventory_record(record)
        if str(normalized.get("status") or "").lower() != "active":
            return False
        company = str(company or "").strip()
        sold_company = company or "General Sold"
        profit_record = self._inventory_sale_profit_record(normalized, company, sale_price, sale_date=sale_date, sale_method=sale_method)
        added = self.record_profit_sales([profit_record])
        changed = self._mark_inventory_record_sold(str(normalized.get("inventory_key") or ""), sold_company, sale_price)
        return bool(added or changed)

    def _inventory_sale_dialog(self, record: dict[str, object]) -> tuple[str, float] | None:
        normalized = self._normalize_inventory_record(record)
        default_sale = self._money_value(normalized.get("estimated_payout")) or self._money_value(normalized.get("inventory_value")) or 0.0
        company_var = tk.StringVar(value="")
        sale_var = tk.StringVar(value=f"{default_sale:.2f}" if default_sale else "")
        result: dict[str, object] = {}

        popup = tk.Toplevel(self)
        popup.title("Mark Inventory Sold")
        popup.configure(bg="#1f1f1f")
        popup.transient(self)
        popup.grab_set()
        popup.resizable(False, False)

        frame = ttk.Frame(popup, style="Panel.TFrame", padding=(18, 16))
        frame.pack(fill=tk.BOTH, expand=True)
        ttk.Label(frame, text="Mark Sold", style="Panel.TLabel", font=("Segoe UI Semibold", 12)).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 2))
        ttk.Label(frame, text=str(normalized.get("card_title") or normalized.get("cert_number") or "Inventory card"), style="Muted.TLabel", wraplength=420).grid(row=1, column=0, columnspan=2, sticky="w", pady=(0, 14))
        ttk.Label(frame, text="Company / buyer", style="Panel.TLabel").grid(row=2, column=0, sticky="w", padx=(0, 10), pady=(0, 10))
        ttk.Entry(frame, textvariable=company_var, width=34).grid(row=2, column=1, sticky="ew", pady=(0, 10))
        ttk.Label(frame, text="Sale price", style="Panel.TLabel").grid(row=3, column=0, sticky="w", padx=(0, 10), pady=(0, 14))
        sale_entry = ttk.Entry(frame, textvariable=sale_var, width=34)
        sale_entry.grid(row=3, column=1, sticky="ew", pady=(0, 14))
        status_var = tk.StringVar(value="Leave company / buyer blank for that person's General Sold sheet.")
        ttk.Label(frame, textvariable=status_var, style="Muted.TLabel").grid(row=4, column=0, columnspan=2, sticky="w", pady=(0, 14))

        def submit() -> None:
            sale_price = self._money_value(sale_var.get())
            if sale_price is None or sale_price < 0:
                status_var.set("Enter a valid sale price.")
                sale_entry.focus_set()
                return
            result["company"] = company_var.get().strip()
            result["sale_price"] = float(sale_price)
            popup.destroy()

        buttons = ttk.Frame(frame, style="Panel.TFrame")
        buttons.grid(row=5, column=0, columnspan=2, sticky="e")
        ttk.Button(buttons, text="Cancel", command=popup.destroy, style="Soft.TButton").pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(buttons, text="Mark Sold", command=submit, style="Primary.TButton").pack(side=tk.LEFT)
        frame.columnconfigure(1, weight=1)
        popup.bind("<Return>", lambda _event: submit())
        popup.bind("<Escape>", lambda _event: popup.destroy())
        popup.update_idletasks()
        x = self.winfo_rootx() + max(80, (self.winfo_width() - popup.winfo_width()) // 2)
        y = self.winfo_rooty() + max(80, (self.winfo_height() - popup.winfo_height()) // 2)
        popup.geometry(f"+{x}+{y}")
        sale_entry.focus_set()
        self.wait_window(popup)
        if "sale_price" not in result:
            return None
        return str(result.get("company") or ""), float(result["sale_price"])

    def mark_selected_inventory_sold(self) -> None:
        if not hasattr(self, "inventory_tree"):
            return
        selected = list(self.inventory_tree.selection())
        records = [self.inventory_tree_records.get(iid) for iid in selected if self.inventory_tree_records.get(iid)]
        if len(records) != 1:
            messagebox.showinfo("Choose one card", "Select one active inventory card to mark sold.")
            return
        record = self._normalize_inventory_record(records[0])
        if str(record.get("status") or "").lower() != "active":
            messagebox.showinfo("Active card required", "Only active inventory cards can be marked sold.")
            return
        sale = self._inventory_sale_dialog(record)
        if sale is None:
            return
        company, sale_price = sale
        if self.mark_inventory_record_sold(record, company, float(sale_price)):
            self.refresh_inventory_tab()
            self.refresh_profit_tab()
            self.status_var.set(f"Marked inventory card sold: {record.get('cert_number') or record.get('card_title') or 'card'} for {format_money(sale_price)}.")

    def move_selected_inventory_to_company_sheets(self) -> None:
        if not hasattr(self, "inventory_tree"):
            return
        selected = list(self.inventory_tree.selection())
        records = [self.inventory_tree_records.get(iid) for iid in selected if self.inventory_tree_records.get(iid)]
        movable_records = [record for record in records if self._inventory_record_can_move_to_company_sheet(record)]
        if not movable_records:
            messagebox.showinfo("Choose inventory", "Select one or more active inventory rows with an assignable best company.")
            return
        confirmed = messagebox.askyesno(
            "Move inventory card(s)?",
            f"Move {len(movable_records)} active inventory card(s) to company sheets?",
        )
        if not confirmed:
            return
        rows: list[WorkbookRow] = []
        source_lookup: dict[int, str] = {}
        sheet_source_lookup: dict[int, str] = {}
        people_by_cert: dict[str, str] = {}
        keys_by_cert: dict[str, str] = {}
        unassigned = 0
        for index, record in enumerate(movable_records, start=1):
            row = self._inventory_workbook_row(record, index)
            recommendation = self.assignment_engine.recommend(row, person=str(record.get("assigned_person") or ""))
            if recommendation.payout is None:
                unassigned += 1
                continue
            row.best_company = recommendation.company
            row.estimated_payout = recommendation.payout
            rows.append(row)
            source_lookup[index] = str(record.get("source") or "Inventory")
            sheet_source_lookup[index] = str(record.get("source_sheet") or "Inventory")
            cert = scan_to_cert(row.cert_number)
            if cert:
                people_by_cert[cert] = str(record.get("assigned_person") or "")
                keys_by_cert[cert] = str(record.get("inventory_key") or "")
        if not rows:
            messagebox.showinfo("No company match", "No selected inventory cards matched an assignable company.")
            return
        with shared_lock(CARD_PIPELINE_DIR, "inventory-company-sheets", self.lucas_identity):
            company_result = append_company_sheet_rows(COMPANY_SHEETS_DIR, rows, source_lookup, sheet_source_lookup)
            added_records = list(company_result.get("added_records") or [])
            moved_keys: set[str] = set()
            for record in added_records:
                cert = scan_to_cert(record.get("cert_number"))
                record["assigned_person"] = people_by_cert.get(cert, "")
                key = keys_by_cert.get(cert, "")
                if key:
                    moved_keys.add(key)
            if added_records:
                self.record_profit_sales(added_records)
            self._mark_inventory_records_moved_to_company(moved_keys)
        self.refresh_inventory_tab()
        self.refresh_profit_tab()
        added = int(company_result.get("rows_added") or 0)
        errors = company_result.get("errors") or []
        suffix = f" {unassigned} card(s) had no assignable company." if unassigned else ""
        self.status_var.set(f"Moved {added} inventory card(s) to company sheets.{suffix}")
        if errors:
            messagebox.showwarning("Inventory move completed with warnings", "\n".join([f"Moved rows: {added}", *errors[:8]]))

    def _inventory_record_can_move_to_company_sheet(self, record: dict[str, object] | None) -> bool:
        if not record or str(record.get("status") or "").lower() != "active":
            return False
        best_company = str(record.get("best_company") or "").strip()
        return bool(best_company) and best_company.upper() != NO_COMPANY_TAKES_LABEL

    def _inventory_tree_cell_text(self, row_id: str, column_id: str) -> str:
        if not row_id or not column_id:
            return ""
        values = self.inventory_tree.item(row_id, "values") or ()
        try:
            index = int(str(column_id).lstrip("#")) - 1
        except ValueError:
            return ""
        if index < 0 or index >= len(values):
            return ""
        return str(values[index] or "")

    def _inventory_tree_row_text(self, row_id: str) -> str:
        if not row_id:
            return ""
        values = self.inventory_tree.item(row_id, "values") or ()
        return "\t".join(str(value or "") for value in values)

    def _copy_inventory_text(self, text: str, label: str = "inventory value") -> None:
        self.clipboard_clear()
        self.clipboard_append(str(text or ""))
        if hasattr(self, "status_var"):
            self.status_var.set(f"Copied {label}.")

    def copy_inventory_cell_value(self, row_id: str, column_id: str) -> None:
        self._copy_inventory_text(self._inventory_tree_cell_text(row_id, column_id), "inventory cell")

    def copy_inventory_row_values(self, row_id: str) -> None:
        self._copy_inventory_text(self._inventory_tree_row_text(row_id), "inventory row")

    def _show_inventory_context_menu(self, event) -> str:
        if not hasattr(self, "inventory_tree"):
            return "break"
        row_id = self.inventory_tree.identify_row(event.y)
        if not row_id:
            return "break"
        column_id = self.inventory_tree.identify_column(event.x)
        if row_id not in self.inventory_tree.selection():
            self.inventory_tree.selection_set(row_id)
            self.inventory_tree.focus(row_id)
        records = [self.inventory_tree_records.get(iid) for iid in self.inventory_tree.selection()]
        active_records = [record for record in records if record and str(record.get("status") or "").lower() == "active"]
        menu = tk.Menu(self, tearoff=False, bg="#1f1f1f", fg="#ffffff", activebackground="#1ed760", activeforeground="#000000")
        menu.add_command(label="Copy Cell", command=lambda row=row_id, column=column_id: self.copy_inventory_cell_value(row, column))
        menu.add_command(label="Copy Row", command=lambda row=row_id: self.copy_inventory_row_values(row))
        if len(active_records) == 1 and len(records) == 1:
            menu.add_separator()
            menu.add_command(label="Mark Sold", command=self.mark_selected_inventory_sold)
        if records and all(self._inventory_record_can_move_to_company_sheet(record) for record in records):
            if len(active_records) == 1 and len(records) == 1:
                menu.add_separator()
            menu.add_command(label="Move to Company Sheets", command=self.move_selected_inventory_to_company_sheets)
        if records:
            menu.add_separator()
            menu.add_command(label="Delete from Inventory", command=self.delete_selected_inventory_records)
        if menu.index("end") is None:
            return "break"
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()
        return "break"

    def refresh_inventory_tab(self, reconcile: bool = False, enrich: bool = False, filtered_only: bool = False) -> None:
        if reconcile and not getattr(self, "_inventory_reconcile_running", False):
            self._inventory_reconcile_running = True
            try:
                self._sync_received_inventory_to_ledger()
            finally:
                self._inventory_reconcile_running = False
        stored_rows = [self._normalize_inventory_record(record) for record in self._load_inventory_ledger()]
        active_rows = [record for record in stored_rows if str(record.get("status") or "").lower() == "active"]
        if len(active_rows) != len(stored_rows):
            self._save_inventory_ledger(active_rows)
            stored_rows = active_rows
        if enrich and filtered_only:
            filtered_keys = {str(record.get("inventory_key") or "") for record in self._filtered_inventory_records(stored_rows)}
            self.inventory_rows = [
                self._enrich_inventory_record_assignment(record, force=True) if str(record.get("inventory_key") or "") in filtered_keys else record
                for record in stored_rows
            ]
        else:
            self.inventory_rows = [self._enrich_inventory_record_assignment(record) for record in stored_rows] if enrich else stored_rows
        if enrich and self.inventory_rows != stored_rows:
            self._save_inventory_ledger(self.inventory_rows)
        self.filtered_inventory_rows = self._filtered_inventory_records(self.inventory_rows)
        if not hasattr(self, "inventory_tree"):
            return
        self._refresh_person_combo_values()
        self.inventory_tree.delete(*self.inventory_tree.get_children())
        self.inventory_tree_records = {}
        total_purchase = 0.0
        total_value = 0.0
        for record in self.filtered_inventory_rows:
            purchase = self._money_value(record.get("purchase_price"))
            value = self._money_value(record.get("inventory_value"))
            card_ladder = self._money_value(record.get("card_ladder_value"))
            comps = self._money_value(record.get("card_ladder_comps_average"))
            cy_value = self._money_value(record.get("cy_value"))
            if purchase is not None:
                total_purchase += purchase
            if value is not None:
                total_value += value
            iid = self.inventory_tree.insert(
                "",
                tk.END,
                values=(
                    record.get("date_added") or "",
                    record.get("assigned_person") or "Unassigned",
                    record.get("sport") or "",
                    record.get("cert_number") or "",
                    record.get("grader") or "",
                    record.get("card_title") or "",
                    format_money(purchase),
                    format_money(card_ladder),
                    format_money(comps),
                    format_money(cy_value),
                    record.get("cy_confidence") if record.get("cy_confidence") is not None else "",
                    record.get("best_company") or "",
                    format_money(record.get("estimated_payout")),
                    record.get("source_sheet") or "",
                    record.get("status") or "",
                ),
            )
            self.inventory_tree_records[iid] = record
        self.inventory_metric_var.set(f"Cards: {len(self.filtered_inventory_rows)}   Purchase Total: {format_money(total_purchase)}   Source Value: {format_money(total_value)}")
        self.inventory_status_var.set(f"Loaded {len(self.filtered_inventory_rows)}/{len(self.inventory_rows)} inventory card(s) from {INVENTORY_LEDGER_PATH.name}.")

    def delete_selected_inventory_records(self) -> None:
        if not hasattr(self, "inventory_tree"):
            return
        records = [self.inventory_tree_records.get(iid) for iid in self.inventory_tree.selection()]
        keys = {str(record.get("inventory_key") or "") for record in records if record}
        keys.discard("")
        if not keys:
            messagebox.showinfo("Choose inventory", "Select one or more inventory rows to delete.")
            return
        confirmed = messagebox.askyesno(
            "Delete from inventory?",
            f"Delete {len(keys)} selected inventory item(s)?\n\nThis only removes them from Inventory.",
        )
        if not confirmed:
            return
        deleted = self._delete_inventory_records_by_keys(keys)
        self.refresh_inventory_tab()
        self.status_var.set(f"Deleted {deleted} inventory item(s).")

    def _delete_inventory_records_by_keys(self, keys: set[str]) -> int:
        if not keys:
            return 0
        with shared_lock(CARD_PIPELINE_DIR, "inventory-delete", self.lucas_identity):
            rows = [self._normalize_inventory_record(record) for record in self._load_inventory_ledger()]
            kept = [record for record in rows if str(record.get("inventory_key") or "") not in keys]
            deleted = len(rows) - len(kept)
            if deleted:
                self._save_inventory_ledger(kept)
        return deleted

    def _filtered_inventory_records(self, rows: list[dict[str, object]]) -> list[dict[str, object]]:
        person = self.inventory_person_var.get().strip().lower() if hasattr(self, "inventory_person_var") else ""
        sport = self.inventory_sport_var.get().strip().lower() if hasattr(self, "inventory_sport_var") else ""
        search = self.inventory_search_var.get().strip().lower() if hasattr(self, "inventory_search_var") else ""
        min_value = self._money_value(self.inventory_min_var.get()) if hasattr(self, "inventory_min_var") else None
        max_value = self._money_value(self.inventory_max_var.get()) if hasattr(self, "inventory_max_var") else None
        filtered: list[dict[str, object]] = []
        for record in rows:
            if str(record.get("status") or "").lower() != "active":
                continue
            if person and person not in str(record.get("assigned_person") or "Unassigned").lower():
                continue
            if sport and sport not in str(record.get("sport") or "").lower():
                continue
            if search:
                searchable = f"{record.get('cert_number') or ''} {record.get('card_title') or ''}".lower()
                if any(part not in searchable for part in search.split()):
                    continue
            value = self._money_value(record.get("inventory_value") or record.get("purchase_price")) or 0.0
            if min_value is not None and value < min_value:
                continue
            if max_value is not None and value > max_value:
                continue
            filtered.append(record)
        return filtered

    def export_inventory(self) -> None:
        rows = self.filtered_inventory_rows if hasattr(self, "filtered_inventory_rows") else []
        if not rows:
            messagebox.showinfo("No inventory", "No inventory rows match the current filters.")
            return
        path = filedialog.asksaveasfilename(
            title="Export inventory",
            defaultextension=".xlsx",
            filetypes=[("Excel workbook", "*.xlsx")],
            initialfile=f"inventory-{datetime.now():%Y%m%d-%H%M%S}.xlsx",
        )
        if not path:
            return
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Inventory"
        headers = ["Date Added", "Person", "Sport", "Certification Number", "Grader", "Card Description", "Purchase Price", "Card Ladder", "Comps", "CY Estimate", "CY Confidence", "Best Company", "Estimated Payout", "Source Sheet", "Source", "Status", "Notes"]
        sheet.append(headers)
        for record in rows:
            sheet.append([
                record.get("date_added") or "",
                record.get("assigned_person") or "",
                record.get("sport") or "",
                record.get("cert_number") or "",
                record.get("grader") or "",
                record.get("card_title") or "",
                record.get("purchase_price"),
                record.get("card_ladder_value"),
                record.get("card_ladder_comps_average"),
                record.get("cy_value"),
                record.get("cy_confidence"),
                record.get("best_company") or "",
                record.get("estimated_payout"),
                record.get("source_sheet") or "",
                record.get("source") or "",
                record.get("status") or "",
                record.get("notes") or "",
            ])
        sheet.auto_filter.ref = sheet.dimensions
        sheet.freeze_panes = "A2"
        for index, width in enumerate([14, 18, 14, 22, 12, 60, 16, 16, 16, 16, 14, 20, 16, 28, 24, 14, 36], start=1):
            sheet.column_dimensions[sheet.cell(1, index).column_letter].width = width
        workbook.save(path)
        self.status_var.set(f"Exported inventory: {path}")

    def _profit_record_key(self, record: dict[str, object]) -> str:
        record_type = str(record.get("record_type") or "").strip().lower()
        if record_type == "expense":
            return "|".join(
                str(record.get(field) or "").strip().lower()
                for field in ("record_type", "expense_id", "assigned_person", "date_added", "expense_type", "expense_amount", "related_type", "source_sheet", "cert_number", "notes")
            )
        return "|".join(
            str(record.get(field) or "").strip().lower()
            for field in ("cert_number", "company", "date_added", "weekly_sheet_name", "source_sheet")
        )

    def _money_value(self, value: object) -> float | None:
        if value is None or value == "":
            return None
        match = re.search(r"-?[\d,]+(?:\.\d{1,2})?", str(value))
        if not match:
            return None
        try:
            return float(match.group(0).replace(",", ""))
        except ValueError:
            return None

    def _normalize_profit_record(self, record: dict[str, object]) -> dict[str, object]:
        normalized = dict(record)
        record_type = str(normalized.get("record_type") or "").strip().lower()
        if record_type == "expense":
            amount = self._money_value(normalized.get("expense_amount") or normalized.get("amount") or normalized.get("purchase_price")) or 0.0
            expense_type = str(normalized.get("expense_type") or normalized.get("category") or "Fees").strip() or "Fees"
            if expense_type not in EXPENSE_CATEGORY_OPTIONS:
                expense_type = "Fees"
            notes = str(normalized.get("notes") or "").strip()
            related_type = str(normalized.get("related_type") or normalized.get("tie_to") or "General").strip()
            if related_type not in EXPENSE_LINK_OPTIONS:
                related_type = "General"
            related_sheet = str(normalized.get("source_sheet") or normalized.get("related_sheet") or "").strip()
            related_cert = str(normalized.get("cert_number") or normalized.get("related_cert") or "").strip()
            normalized["record_type"] = "expense"
            normalized["expense_id"] = str(normalized.get("expense_id") or "").strip()
            normalized["expense_type"] = expense_type
            normalized["expense_amount"] = round(abs(amount), 2)
            normalized["related_type"] = related_type
            normalized["purchase_price"] = None
            normalized["sale_price"] = None
            normalized["profit"] = -round(abs(amount), 2)
            normalized["date_added"] = str(normalized.get("date_added") or datetime.now().strftime("%Y-%m-%d"))[:10]
            normalized["company"] = f"Expense: {expense_type}"
            normalized["card_title"] = notes or expense_type
            normalized["cert_number"] = related_cert if related_type == "Card" else ""
            normalized["weekly_sheet_name"] = ""
            normalized["source_sheet"] = related_sheet if related_type in {"Card", "Sheet"} and related_sheet else "Expenses"
            normalized["assigned_person"] = str(normalized.get("assigned_person") or normalized.get("person") or "").strip()
            normalized["notes"] = notes
            normalized["ledger_key"] = self._profit_record_key(normalized)
            return normalized
        purchase = self._money_value(normalized.get("purchase_price"))
        sale = self._money_value(normalized.get("sale_price"))
        normalized["purchase_price"] = purchase
        normalized["sale_price"] = sale
        normalized["profit"] = round(sale - purchase, 2) if sale is not None and purchase is not None else None
        normalized["date_added"] = str(normalized.get("date_added") or datetime.now().strftime("%Y-%m-%d"))
        normalized["company"] = str(normalized.get("company") or normalized.get("best_company") or "").strip()
        normalized["card_title"] = str(normalized.get("card_title") or "").strip()
        normalized["cert_number"] = str(normalized.get("cert_number") or "").strip()
        normalized["weekly_sheet_name"] = str(normalized.get("weekly_sheet_name") or "").strip()
        normalized["source_sheet"] = str(normalized.get("source_sheet") or "").strip()
        normalized["assigned_person"] = str(normalized.get("assigned_person") or normalized.get("person") or "").strip()
        normalized["ledger_key"] = self._profit_record_key(normalized)
        return normalized

    def _person_for_profit_record(self, record: dict[str, object]) -> str:
        existing = str(record.get("assigned_person") or "").strip()
        if existing:
            return existing
        source_sheet = Path(str(record.get("source_sheet") or "")).name
        if not source_sheet:
            return ""
        for stage in ("Incoming", "Received", "Working"):
            marker = self.home_sheet_markers.get(self._home_sheet_key(stage, source_sheet), {})
            person = str(marker.get("assigned_person") or "").strip()
            if person:
                return person
        for key, marker in self.home_sheet_markers.items():
            _stage, name = self._split_home_sheet_key(key)
            if Path(name).name == source_sheet:
                person = str(marker.get("assigned_person") or "").strip()
                if person:
                    return person
        return ""

    def _enrich_profit_records_with_people(self, rows: list[dict[str, object]]) -> list[dict[str, object]]:
        enriched: list[dict[str, object]] = []
        for record in rows:
            normalized = self._normalize_profit_record(record)
            normalized["assigned_person"] = self._person_for_profit_record(normalized)
            enriched.append(normalized)
        return enriched

    def _filtered_profit_records(self, rows: list[dict[str, object]]) -> list[dict[str, object]]:
        needle = self.profit_person_var.get().strip().lower() if hasattr(self, "profit_person_var") else ""
        period = self.profit_period_var.get().strip() if hasattr(self, "profit_period_var") else "Total"
        period_start, period_end = self._profit_period_bounds(period)
        filtered: list[dict[str, object]] = []
        for record in rows:
            if needle and needle not in (str(record.get("assigned_person") or "Unassigned").lower()):
                continue
            if period_start is not None:
                sold_date = self._profit_record_date(record.get("date_added"))
                if sold_date is None or sold_date < period_start or sold_date > period_end:
                    continue
            filtered.append(record)
        return filtered

    def _profit_record_date(self, value: object):
        text = str(value or "").strip()
        if not text:
            return None
        try:
            return datetime.strptime(text[:10], "%Y-%m-%d").date()
        except ValueError:
            return None

    def _profit_today(self):
        return datetime.now().date()

    def _profit_period_bounds(self, period: str, as_of=None):
        today = as_of or self._profit_today()
        label = (period or "Total").strip().lower()
        if label in {"5 days", "5 day", "five days"}:
            return today - timedelta(days=4), today
        if label == "week":
            return today - timedelta(days=6), today
        if label == "month":
            return today.replace(day=1), today
        if label in {"year", "ytd", "year to date"}:
            return today.replace(month=1, day=1), today
        return None, today

    def _profit_period_label(self) -> str:
        period = self.profit_period_var.get().strip() if hasattr(self, "profit_period_var") else "Total"
        return period if period in PROFIT_PERIOD_OPTIONS else "Total"

    def _profit_graph_label(self) -> str:
        graph = self.profit_graph_var.get().strip() if hasattr(self, "profit_graph_var") else "Daily Trend"
        return graph if graph in PROFIT_GRAPH_OPTIONS else "Daily Trend"

    def _profit_chart_series(self, rows: list[dict[str, object]]) -> tuple[list[str], list[float]]:
        daily: dict[str, float] = {}
        for record in rows:
            profit = self._money_value(record.get("profit"))
            sold_date = self._profit_record_date(record.get("date_added"))
            if profit is None or sold_date is None:
                continue
            day = sold_date.isoformat()
            daily[day] = daily.get(day, 0.0) + float(profit)
        period_start, period_end = self._profit_period_bounds(self._profit_period_label())
        if period_start is not None:
            cursor = period_start
            while cursor <= period_end:
                daily.setdefault(cursor.isoformat(), 0.0)
                cursor += timedelta(days=1)
        days = sorted(daily)
        daily_values = [daily[day] for day in days]
        if self._profit_graph_label() != "Overall Profit":
            return days, daily_values
        running = 0.0
        cumulative_values: list[float] = []
        for value in daily_values:
            running += value
            cumulative_values.append(running)
        return days, cumulative_values

    def _set_profit_view_mode(self, mode: str) -> None:
        self.profit_view_mode.set(mode)
        self.refresh_profit_tab()

    def _configure_profit_tree(self, mode: str) -> None:
        if not hasattr(self, "profit_tree"):
            return
        if mode == "Expenses":
            columns = ("date", "person", "type", "amount", "related", "cert", "notes")
            headings = {
                "date": "Date",
                "person": "Person",
                "type": "Type",
                "amount": "Amount",
                "related": "Related",
                "cert": "Cert",
                "notes": "Notes",
            }
            widths = {"date": 95, "person": 150, "type": 120, "amount": 105, "related": 260, "cert": 110, "notes": 320}
        elif mode == "Sold Sheets":
            columns = ("person", "sheet", "companies", "cards", "purchase", "sale", "profit", "last_sale")
            headings = {
                "person": "Person",
                "sheet": "Sold Sheet",
                "companies": "Companies",
                "cards": "Cards",
                "purchase": "Purchase",
                "sale": "Sale Price",
                "profit": "Profit",
                "last_sale": "Last Sale",
            }
            widths = {"person": 150, "sheet": 300, "companies": 220, "cards": 80, "purchase": 105, "sale": 105, "profit": 105, "last_sale": 95}
        else:
            columns = ("date", "person", "company", "card", "cert", "purchase", "sale", "profit", "sheet")
            headings = {
                "date": "Date",
                "person": "Person",
                "company": "Company",
                "card": "Card",
                "cert": "Cert",
                "purchase": "Purchase",
                "sale": "Sale Price",
                "profit": "Profit",
                "sheet": "Company Sheet",
            }
            widths = {"date": 95, "person": 135, "company": 140, "card": 390, "cert": 100, "purchase": 105, "sale": 105, "profit": 105, "sheet": 200}
        self.profit_tree.configure(columns=columns)
        for column in columns:
            self.profit_tree.heading(column, text=headings[column], anchor=tk.W)
            self.profit_tree.column(column, width=widths[column], minwidth=45, stretch=False)
        self.profit_table_title_var.set(mode)

    def _expense_related_label(self, record: dict[str, object]) -> str:
        related_type = str(record.get("related_type") or "General").strip() or "General"
        source_sheet = str(record.get("source_sheet") or "").strip()
        cert = str(record.get("cert_number") or "").strip()
        if related_type == "Card":
            parts = [part for part in (source_sheet if source_sheet != "Expenses" else "", cert) if part]
            return " | ".join(parts) or "Card"
        if related_type == "Sheet":
            return source_sheet if source_sheet and source_sheet != "Expenses" else "Sheet"
        return "General"

    def _profit_sheet_rows(self, rows: list[dict[str, object]]) -> list[dict[str, object]]:
        grouped: dict[tuple[str, str], dict[str, object]] = {}
        for record in rows:
            source_sheet = str(record.get("source_sheet") or "").strip() or "Unknown sheet"
            person = str(record.get("assigned_person") or "").strip() or "Unassigned"
            key = (person, source_sheet)
            group = grouped.setdefault(
                key,
                {
                    "person": person,
                    "sheet": source_sheet,
                    "companies": set(),
                    "cards": 0,
                    "purchase": 0.0,
                    "sale": 0.0,
                    "profit": 0.0,
                    "complete": 0,
                    "last_sale": "",
                },
            )
            company = str(record.get("company") or "").strip()
            if company:
                group["companies"].add(company)
            if str(record.get("record_type") or "").strip().lower() != "expense":
                group["cards"] = int(group["cards"]) + 1
            purchase = self._money_value(record.get("purchase_price"))
            sale = self._money_value(record.get("sale_price"))
            profit = self._money_value(record.get("profit"))
            if purchase is not None:
                group["purchase"] = float(group["purchase"]) + purchase
            if sale is not None:
                group["sale"] = float(group["sale"]) + sale
            if profit is not None:
                group["profit"] = float(group["profit"]) + profit
                group["complete"] = int(group["complete"]) + 1
            date = str(record.get("date_added") or "")[:10]
            if date and date > str(group["last_sale"]):
                group["last_sale"] = date
        result: list[dict[str, object]] = []
        for group in grouped.values():
            companies = sorted(group.pop("companies"), key=str.lower)
            group["companies"] = ", ".join(companies)
            result.append(group)
        return sorted(result, key=lambda item: (str(item.get("last_sale") or ""), str(item.get("person") or ""), str(item.get("sheet") or "")), reverse=True)

    def _append_profit_records(self, records: list[dict[str, object]]) -> int:
        if not records:
            return 0
        with shared_lock(CARD_PIPELINE_DIR, "profit-ledger", self.lucas_identity):
            ledger = [self._normalize_profit_record(record) for record in self._load_profit_ledger()]
            existing_keys = {str(record.get("ledger_key") or self._profit_record_key(record)) for record in ledger}
            added = 0
            for record in records:
                normalized = self._normalize_profit_record(record)
                key = str(normalized.get("ledger_key") or "")
                if not key or key in existing_keys:
                    continue
                normalized["recorded_by"] = self.lucas_identity.get("display_name", "")
                normalized["recorded_machine"] = self.lucas_identity.get("machine", "")
                ledger.append(normalized)
                existing_keys.add(key)
                added += 1
            if added:
                self._save_profit_ledger(ledger)
        return added

    def record_profit_sales(self, records: list[dict[str, object]]) -> int:
        added = CardPipelineApp._append_profit_records(self, records)
        self.refresh_profit_tab()
        return added

    def open_add_expense_popup(self) -> None:
        person_var = tk.StringVar(value=self.profit_person_var.get().strip() if hasattr(self, "profit_person_var") else "")
        date_var = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))
        type_var = tk.StringVar(value=EXPENSE_CATEGORY_OPTIONS[0])
        amount_var = tk.StringVar()
        link_var = tk.StringVar(value=EXPENSE_LINK_OPTIONS[0])
        sheet_var = tk.StringVar()
        cert_var = tk.StringVar()
        notes_var = tk.StringVar()

        popup = tk.Toplevel(self)
        popup.title("Add Expense")
        popup.configure(bg="#1f1f1f")
        popup.transient(self)
        popup.grab_set()
        popup.resizable(False, False)

        frame = ttk.Frame(popup, style="Panel.TFrame", padding=(18, 16))
        frame.pack(fill=tk.BOTH, expand=True)
        ttk.Label(frame, text="Add Expense", style="Panel.TLabel", font=("Segoe UI Semibold", 12)).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 12))
        ttk.Label(frame, text="Person", style="Panel.TLabel").grid(row=1, column=0, sticky="w", padx=(0, 10), pady=(0, 10))
        person_combo = ttk.Combobox(frame, textvariable=person_var, width=34)
        person_combo.grid(row=1, column=1, sticky="ew", pady=(0, 10))
        self._bind_person_autocomplete(person_combo)
        ttk.Label(frame, text="Date", style="Panel.TLabel").grid(row=2, column=0, sticky="w", padx=(0, 10), pady=(0, 10))
        ttk.Entry(frame, textvariable=date_var, width=18).grid(row=2, column=1, sticky="w", pady=(0, 10))
        ttk.Label(frame, text="Type", style="Panel.TLabel").grid(row=3, column=0, sticky="w", padx=(0, 10), pady=(0, 10))
        ttk.Combobox(frame, textvariable=type_var, values=EXPENSE_CATEGORY_OPTIONS, width=18, state="readonly").grid(row=3, column=1, sticky="w", pady=(0, 10))
        ttk.Label(frame, text="Amount", style="Panel.TLabel").grid(row=4, column=0, sticky="w", padx=(0, 10), pady=(0, 10))
        ttk.Entry(frame, textvariable=amount_var, width=18).grid(row=4, column=1, sticky="w", pady=(0, 10))
        ttk.Label(frame, text="Tie To", style="Panel.TLabel").grid(row=5, column=0, sticky="w", padx=(0, 10), pady=(0, 10))
        ttk.Combobox(frame, textvariable=link_var, values=EXPENSE_LINK_OPTIONS, width=18, state="readonly").grid(row=5, column=1, sticky="w", pady=(0, 10))
        ttk.Label(frame, text="Sheet", style="Panel.TLabel").grid(row=6, column=0, sticky="w", padx=(0, 10), pady=(0, 10))
        ttk.Entry(frame, textvariable=sheet_var, width=36).grid(row=6, column=1, sticky="ew", pady=(0, 10))
        ttk.Label(frame, text="Cert", style="Panel.TLabel").grid(row=7, column=0, sticky="w", padx=(0, 10), pady=(0, 10))
        ttk.Entry(frame, textvariable=cert_var, width=20).grid(row=7, column=1, sticky="w", pady=(0, 10))
        ttk.Label(frame, text="Notes", style="Panel.TLabel").grid(row=8, column=0, sticky="w", padx=(0, 10), pady=(0, 14))
        ttk.Entry(frame, textvariable=notes_var, width=36).grid(row=8, column=1, sticky="ew", pady=(0, 14))
        buttons = ttk.Frame(frame, style="Panel.TFrame")
        buttons.grid(row=9, column=0, columnspan=2, sticky="e")
        ttk.Button(buttons, text="Cancel", command=popup.destroy, style="Soft.TButton").pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(
            buttons,
            text="Save",
            command=lambda: self._save_expense_from_popup(person_var, date_var, type_var, amount_var, link_var, sheet_var, cert_var, notes_var, popup),
            style="Primary.TButton",
        ).pack(side=tk.LEFT)
        frame.columnconfigure(1, weight=1)
        popup.update_idletasks()
        x = self.winfo_rootx() + max(80, (self.winfo_width() - popup.winfo_width()) // 2)
        y = self.winfo_rooty() + max(80, (self.winfo_height() - popup.winfo_height()) // 2)
        popup.geometry(f"+{x}+{y}")

    def _save_expense_from_popup(
        self,
        person_var: tk.StringVar,
        date_var: tk.StringVar,
        type_var: tk.StringVar,
        amount_var: tk.StringVar,
        link_var: tk.StringVar,
        sheet_var: tk.StringVar,
        cert_var: tk.StringVar,
        notes_var: tk.StringVar,
        popup: tk.Toplevel,
    ) -> None:
        person = person_var.get().strip()
        if not person:
            messagebox.showinfo("Person required", "Choose the person this expense belongs to.")
            return
        expense_date = date_var.get().strip()
        if self._profit_record_date(expense_date) is None:
            messagebox.showinfo("Date required", "Enter the expense date as YYYY-MM-DD.")
            return
        amount = self._money_value(amount_var.get())
        if amount is None or amount <= 0:
            messagebox.showinfo("Amount required", "Enter an expense amount greater than zero.")
            return
        expense_type = type_var.get().strip()
        if expense_type not in EXPENSE_CATEGORY_OPTIONS:
            expense_type = "Fees"
        related_type = link_var.get().strip()
        if related_type not in EXPENSE_LINK_OPTIONS:
            related_type = "General"
        related_sheet = sheet_var.get().strip()
        related_cert = cert_var.get().strip()
        if related_type == "Sheet" and not related_sheet:
            messagebox.showinfo("Sheet required", "Enter the sold sheet this expense belongs to.")
            return
        if related_type == "Card" and not (related_sheet or related_cert):
            messagebox.showinfo("Card required", "Enter a cert number or sold sheet for the card expense.")
            return
        record = {
            "record_type": "expense",
            "expense_id": datetime.now().strftime("%Y%m%d%H%M%S%f"),
            "date_added": expense_date[:10],
            "assigned_person": person,
            "expense_type": expense_type,
            "expense_amount": amount,
            "related_type": related_type,
            "source_sheet": related_sheet,
            "cert_number": related_cert,
            "notes": notes_var.get().strip(),
        }
        added = self.record_profit_sales([record])
        if added:
            popup.destroy()
            self.status_var.set(f"Added {expense_type} expense for {person}: {format_money(amount)}.")
        else:
            messagebox.showinfo("Expense not added", "That expense already exists in the profit ledger.")

    def _delete_profit_expense_records(self, records: list[dict[str, object]]) -> int:
        expense_keys: set[str] = set()
        for record in records:
            normalized = self._normalize_profit_record(record)
            if str(normalized.get("record_type") or "").strip().lower() != "expense":
                continue
            key = str(normalized.get("ledger_key") or self._profit_record_key(normalized) or "")
            if key:
                expense_keys.add(key)
        if not expense_keys:
            return 0
        with shared_lock(CARD_PIPELINE_DIR, "profit-expense-delete", self.lucas_identity):
            ledger = [self._normalize_profit_record(record) for record in self._load_profit_ledger()]
            kept: list[dict[str, object]] = []
            deleted = 0
            for record in ledger:
                key = str(record.get("ledger_key") or self._profit_record_key(record) or "")
                is_expense = str(record.get("record_type") or "").strip().lower() == "expense"
                if is_expense and key in expense_keys:
                    deleted += 1
                    continue
                kept.append(record)
            if deleted:
                self._save_profit_ledger(kept)
            return deleted

    def delete_selected_profit_expenses(self) -> None:
        if not hasattr(self, "profit_tree"):
            return
        selected = list(self.profit_tree.selection())
        records = [self.profit_tree_records.get(iid) for iid in selected if self.profit_tree_records.get(iid)]
        if not records:
            messagebox.showinfo("Choose expense", "Select one or more expense rows to delete.")
            return
        if any(str(record.get("record_type") or "").strip().lower() != "expense" for record in records):
            messagebox.showinfo("Delete expenses only", "Only expense rows can be deleted here. Sold cards should be refunded instead.")
            return
        confirmed = messagebox.askyesno(
            "Delete expense(s)?",
            f"Delete {len(records)} expense row(s) from the profit ledger?",
        )
        if not confirmed:
            return
        deleted = self._delete_profit_expense_records(records)
        self.refresh_profit_tab()
        self.status_var.set(f"Deleted {deleted} expense row(s) from the profit ledger.")

    def refund_selected_profit_to_inventory(self) -> None:
        if not hasattr(self, "profit_tree"):
            return
        selected = list(self.profit_tree.selection())
        records = [self.profit_tree_records.get(iid) for iid in selected if self.profit_tree_records.get(iid)]
        if not records:
            messagebox.showinfo("Choose sold cards", "Select one or more sold card rows to refund.")
            return
        if any(str(record.get("record_type") or "").strip().lower() == "expense" for record in records):
            messagebox.showinfo("Cannot refund expenses", "Expense rows adjust profit only and cannot be returned to inventory.")
            return
        confirmed = messagebox.askyesno(
            "Refund selected card(s)?",
            f"Refund {len(records)} sold card(s) and return them to active inventory?",
        )
        if not confirmed:
            return
        refunded = 0
        inventory_records: list[dict[str, object]] = []
        with shared_lock(CARD_PIPELINE_DIR, "refund-inventory", self.lucas_identity):
            ledger = [self._normalize_profit_record(record) for record in self._load_profit_ledger()]
            refund_keys = {str(self._normalize_profit_record(record).get("ledger_key") or "") for record in records}
            kept = [record for record in ledger if str(record.get("ledger_key") or "") not in refund_keys]
            refunded = len(ledger) - len(kept)
            if refunded:
                self._save_profit_ledger(kept)
            for record in records:
                normalized = self._normalize_profit_record(record)
                source_sheet = str(normalized.get("source_sheet") or "")
                cert = str(normalized.get("cert_number") or "")
                if source_sheet and cert:
                    remove_company_sheet_rows_for_source(COMPANY_SHEETS_DIR, source_sheet, {cert})
                inventory_records.append(
                    self._normalize_inventory_record(
                        {
                            "date_added": datetime.now().strftime("%Y-%m-%d"),
                            "assigned_person": normalized.get("assigned_person") or self._person_for_profit_record(normalized) or "Unassigned",
                            "sport": assignment_engine.parse_card_for_matching(str(normalized.get("card_title") or "")).get("sport"),
                            "cert_number": normalized.get("cert_number") or "",
                            "grader": normalized.get("grader") or "",
                            "card_title": normalized.get("card_title") or "",
                            "purchase_price": normalized.get("purchase_price"),
                            "card_ladder_value": normalized.get("card_ladder_value"),
                            "card_ladder_comps_average": normalized.get("card_ladder_comps_average") or normalized.get("comps"),
                            "cy_value": normalized.get("cy_value") or normalized.get("cy_estimate"),
                            "inventory_value": normalized.get("sale_price") or normalized.get("card_ladder_value") or normalized.get("comps") or normalized.get("cy_estimate"),
                            "source_sheet": normalized.get("source_sheet") or "",
                            "source": normalized.get("source") or "",
                            "status": "Active",
                            "notes": "Refunded from sold cards",
                        }
                    )
                )
            self.add_inventory_records(inventory_records)
        self.refresh_profit_tab()
        self.refresh_inventory_tab()
        self.status_var.set(f"Refunded {refunded or len(records)} card(s) back to active inventory.")

    def refresh_profit_tab(self) -> None:
        ledger = [self._normalize_profit_record(record) for record in self._load_profit_ledger()]
        existing_keys = {str(record.get("ledger_key") or self._profit_record_key(record)) for record in ledger}
        backfilled = 0
        for record in read_company_profit_records(COMPANY_SHEETS_DIR):
            normalized = self._normalize_profit_record(record)
            key = str(normalized.get("ledger_key") or "")
            if not key or key in existing_keys:
                continue
            ledger.append(normalized)
            existing_keys.add(key)
            backfilled += 1
        if backfilled:
            with shared_lock(CARD_PIPELINE_DIR, "profit-ledger", self.lucas_identity):
                current = [self._normalize_profit_record(record) for record in self._load_profit_ledger()]
                current_keys = {str(record.get("ledger_key") or self._profit_record_key(record)) for record in current}
                for record in ledger:
                    key = str(record.get("ledger_key") or "")
                    if key and key not in current_keys:
                        current.append(record)
                        current_keys.add(key)
                self._save_profit_ledger(current)
                ledger = current
        self.profit_rows = self._enrich_profit_records_with_people(ledger)
        self.profit_rows.sort(key=lambda record: (str(record.get("date_added") or ""), str(record.get("company") or ""), str(record.get("card_title") or "")), reverse=True)
        self.filtered_profit_rows = self._filtered_profit_records(self.profit_rows)
        if not hasattr(self, "profit_tree"):
            return
        self._refresh_person_combo_values()
        mode = self.profit_view_mode.get()
        self._configure_profit_tree(mode)
        self.profit_tree.delete(*self.profit_tree.get_children())
        self.profit_tree_records = {}
        total_purchase = 0.0
        total_sale = 0.0
        total_profit = 0.0
        total_expenses = 0.0
        total_gross_profit = 0.0
        complete_count = 0
        for record in self.filtered_profit_rows:
            is_expense = str(record.get("record_type") or "").strip().lower() == "expense"
            purchase = self._money_value(record.get("purchase_price"))
            sale = self._money_value(record.get("sale_price"))
            profit = self._money_value(record.get("profit"))
            if purchase is not None:
                total_purchase += purchase
            if sale is not None:
                total_sale += sale
            if profit is not None:
                total_profit += profit
                if is_expense:
                    total_expenses += abs(profit)
                else:
                    total_gross_profit += profit
                complete_count += 1
            if mode == "Expenses":
                if not is_expense:
                    continue
                iid = self.profit_tree.insert(
                    "",
                    tk.END,
                    values=(
                        record.get("date_added") or "",
                        record.get("assigned_person") or "Unassigned",
                        record.get("expense_type") or "",
                        format_money(record.get("expense_amount")),
                        self._expense_related_label(record),
                        record.get("cert_number") or "",
                        record.get("notes") or "",
                    ),
                    tags=("profit_negative",),
                )
                self.profit_tree_records[iid] = record
            elif mode != "Sold Sheets":
                tag = "profit_negative" if profit is not None and profit < 0 else "profit_positive"
                iid = self.profit_tree.insert(
                    "",
                    tk.END,
                    values=(
                        record.get("date_added") or "",
                        record.get("assigned_person") or "Unassigned",
                        record.get("company") or "",
                        record.get("card_title") or "",
                        record.get("cert_number") or "",
                        format_money(purchase),
                        format_money(sale),
                        format_money(profit),
                        record.get("weekly_sheet_name") or record.get("source_sheet") or "",
                    ),
                    tags=(tag,),
                )
                self.profit_tree_records[iid] = record
        if mode == "Sold Sheets":
            for sheet_row in self._profit_sheet_rows(self.filtered_profit_rows):
                profit = self._money_value(sheet_row.get("profit"))
                tag = "profit_negative" if profit is not None and profit < 0 else "profit_positive"
                self.profit_tree.insert(
                    "",
                    tk.END,
                    values=(
                        sheet_row.get("person") or "",
                        sheet_row.get("sheet") or "",
                        sheet_row.get("companies") or "",
                        sheet_row.get("cards") or 0,
                        format_money(sheet_row.get("purchase")),
                        format_money(sheet_row.get("sale")),
                        format_money(profit),
                        sheet_row.get("last_sale") or "",
                    ),
                    tags=(tag,),
                )
        display_count = len([record for record in self.filtered_profit_rows if str(record.get("record_type") or "").strip().lower() == "expense"]) if mode == "Expenses" else len(self.filtered_profit_rows)
        if self.filtered_profit_rows:
            total_values = (
                ("TOTAL", "", "", format_money(total_expenses), "", "", "")
                if mode == "Expenses"
                else
                ("TOTAL", "", "", f"{len(self.filtered_profit_rows)} card(s)", format_money(total_purchase), format_money(total_sale), format_money(total_profit), "")
                if mode == "Sold Sheets"
                else ("TOTAL", "", "", f"{len(self.filtered_profit_rows)} card(s)", "", format_money(total_purchase), format_money(total_sale), format_money(total_profit), "")
            )
            self.profit_tree.insert(
                "",
                tk.END,
                values=total_values,
                tags=("total_row",),
            )
        self.profit_metric_var.set(
            f"{self._profit_period_label()}   Sales: {format_money(total_sale)}   Gross: {format_money(total_gross_profit)}   Expenses: {format_money(total_expenses)}   Net: {format_money(total_profit)}"
        )
        missing = len(self.filtered_profit_rows) - complete_count
        suffix = f" | {missing} card(s) missing purchase or sale price" if missing else ""
        filter_label = self.profit_person_var.get().strip()
        filter_suffix = f" | Filter: {filter_label}" if filter_label else ""
        period_suffix = f" | Period: {self._profit_period_label()}"
        backfill_suffix = f" | backfilled {backfilled} from company sheets" if backfilled else ""
        self.profit_status_var.set(f"Loaded {display_count}/{len(self.profit_rows)} profit row(s) from {PROFIT_LEDGER_PATH.name}{filter_suffix}{period_suffix}{suffix}{backfill_suffix}.")
        self._draw_profit_chart()

    def _profit_month_key(self, value: object) -> str:
        text = str(value or "").strip()
        try:
            return datetime.strptime(text[:10], "%Y-%m-%d").strftime("%Y-%m")
        except ValueError:
            return text[:7] if len(text) >= 7 else "Unknown"

    def _draw_profit_chart(self) -> None:
        if not hasattr(self, "profit_chart_canvas"):
            return
        canvas = self.profit_chart_canvas
        canvas.delete("all")
        width = max(canvas.winfo_width(), 400)
        height = max(canvas.winfo_height(), 220)
        pad_left, pad_right, pad_top, pad_bottom = 62, 22, 26, 44
        plot_w = max(width - pad_left - pad_right, 10)
        plot_h = max(height - pad_top - pad_bottom, 10)
        chart_rows = self.filtered_profit_rows if hasattr(self, "filtered_profit_rows") else self.profit_rows
        days, chart_values = self._profit_chart_series(chart_rows)
        if not days:
            canvas.create_text(width / 2, height / 2, text="No profit data yet", fill="#b3b3b3", font=("Segoe UI", 12, "bold"))
            return
        graph_label = self._profit_graph_label()
        values = chart_values + [0.0]
        min_y = min(values)
        max_y = max(values)
        if min_y == max_y:
            min_y -= 1
            max_y += 1
        def x_at(index: int) -> float:
            if len(days) == 1:
                return pad_left + plot_w / 2
            return pad_left + (plot_w * index / (len(days) - 1))
        def y_at(value: float) -> float:
            return pad_top + (max_y - value) / (max_y - min_y) * plot_h
        zero_y = y_at(0.0)
        grid_lines = 4
        for line_index in range(grid_lines + 1):
            y = pad_top + (plot_h * line_index / grid_lines)
            value = max_y - ((max_y - min_y) * line_index / grid_lines)
            canvas.create_line(pad_left, y, pad_left + plot_w, y, fill="#2f2f2f")
            canvas.create_text(8, y - 6, anchor="nw", text=format_money(value), fill="#8f8f8f", font=("Segoe UI", 8))
        vertical_lines = min(max(len(days), 2), 8)
        for line_index in range(vertical_lines):
            x = pad_left + (plot_w * line_index / max(vertical_lines - 1, 1))
            canvas.create_line(x, pad_top, x, pad_top + plot_h, fill="#2a2a2a")
        canvas.create_line(pad_left, pad_top, pad_left, pad_top + plot_h, fill="#555555")
        canvas.create_line(pad_left, zero_y, pad_left + plot_w, zero_y, fill="#555555")
        points = [(x_at(index), y_at(value)) for index, value in enumerate(chart_values)]
        for first, second in zip(points, points[1:]):
            canvas.create_line(*first, *second, fill="#22c55e", width=3)
        for index, (x, y) in enumerate(points):
            value = chart_values[index]
            color = "#22c55e" if value >= 0 else "#ef4444"
            canvas.create_oval(x - 4, y - 4, x + 4, y + 4, fill=color, outline="")
            if len(days) <= 14 or index % max(1, len(days) // 8) == 0:
                canvas.create_text(x, height - 24, text=days[index][5:], fill="#b3b3b3", font=("Segoe UI", 8))
        canvas.create_text(pad_left, 8, anchor="nw", text=f"Line: {graph_label.lower()} ({self._profit_period_label()})", fill="#22c55e", font=("Segoe UI", 9, "bold"))

    def choose_working_folder(self) -> None:
        selected = filedialog.askdirectory(
            title="Choose WORKING SHEETS folder",
            initialdir=str(WORKING_SHEETS_DIR if WORKING_SHEETS_DIR.exists() else CARD_PIPELINE_DIR if CARD_PIPELINE_DIR.exists() else ROOT),
        )
        if not selected:
            return
        set_pipeline_from_working_dir(Path(selected))
        settings = load_app_settings()
        settings["pipeline_root"] = str(CARD_PIPELINE_DIR)
        settings["working_sheets_dir"] = str(WORKING_SHEETS_DIR)
        save_app_settings(settings)
        self.pipeline_root_var.set(str(CARD_PIPELINE_DIR))
        for directory in (WORKING_SHEETS_DIR, INCOMING_SHEETS_DIR, RECEIVED_SHEETS_DIR):
            directory.mkdir(parents=True, exist_ok=True)
        COMPANY_SHEETS_DIR.mkdir(parents=True, exist_ok=True)
        self._load_player_overrides()
        self.assignment_engine = AssignmentEngine.load()
        self.assignment_config_status.set(self._assignment_config_status())
        self.home_sheet_markers = self._load_sheet_markers()
        self.refresh_profit_tab()
        self.status_var.set(f"Working folder set to {WORKING_SHEETS_DIR}")
        self.refresh_home()
        self.refresh_working_sheets()
        self.refresh_incoming_index()
        self.refresh_received_sheets()

    def choose_pipeline_root(self) -> None:
        self.choose_working_folder()

    def _build_home_tab_button(self, parent: tk.Frame, text: str, command) -> tk.Button:
        palette = self.home_tab_palette
        return tk.Button(
            parent,
            text=text,
            command=command,
            bg=palette["soft_button"],
            fg=palette["muted"],
            activebackground=palette["soft_button_hover"],
            activeforeground=palette["text"],
            relief=tk.FLAT,
            borderwidth=0,
            highlightthickness=0,
            padx=8,
            pady=6,
            font=("Segoe UI Semibold", 9),
            cursor="hand2",
        )

    def _set_home_sheet_kind(self, kind: str) -> None:
        self.home_sheet_kind.set(kind)
        self._update_home_sheet_tabs()
        self._refresh_home_sheet_list()

    def _on_home_person_filter_changed(self) -> None:
        person = self.home_person_var.get().strip()
        for var in (self.payout_person_var, self.inventory_person_var, self.profit_person_var):
            if var.get().strip() != person:
                var.set(person)
        self._refresh_home_sheet_list()
        self._refresh_home_metrics()
        self.refresh_payouts_tab()
        self.refresh_inventory_tab()
        self.refresh_profit_tab()

    def _home_person_filter(self) -> str:
        return self.home_person_var.get().strip().lower() if hasattr(self, "home_person_var") else ""

    def _home_sheet_matches_person_filter(self, key: str) -> bool:
        needle = self._home_person_filter()
        if not needle:
            return True
        marker = self.home_sheet_markers.get(key, {})
        person = str(marker.get("assigned_person") or "Unassigned").strip().lower()
        return needle in person

    def _filtered_home_sheet_names(self, kind: str) -> list[str]:
        return [
            name
            for name in self.home_sheet_paths.get(kind, {})
            if self._home_sheet_matches_person_filter(self._home_sheet_key(kind, name))
        ]

    def _update_home_sheet_tabs(self) -> None:
        if not hasattr(self, "home_incoming_tab") or not hasattr(self, "home_working_tab"):
            return
        palette = self.home_tab_palette
        active_kind = self.home_sheet_kind.get()
        active = {"bg": palette["panel_high"], "fg": palette["text"], "activebackground": palette["panel_high"], "activeforeground": palette["text"]}
        inactive = {"bg": palette["soft_button"], "fg": palette["muted"], "activebackground": palette["soft_button_hover"], "activeforeground": palette["text"]}
        self.home_incoming_tab.configure(**(active if active_kind == "Incoming" else inactive))
        self.home_working_tab.configure(**(active if active_kind == "Working" else inactive))
        if hasattr(self, "home_received_tab"):
            self.home_received_tab.configure(**(active if active_kind == "Received" else inactive))
        if hasattr(self, "home_edit_markers_tab"):
            self.home_edit_markers_tab.configure(**inactive)

    def refresh_home(self) -> None:
        self.home_sheet_paths = {"Incoming": {}, "Working": {}, "Received": {}}
        self.home_sheet_summaries = {}
        errors: list[str] = []
        conflict_files = self._shared_conflict_files()
        if conflict_files:
            errors.append(f"Shared conflicts: {', '.join(path.name for path in conflict_files[:3])}")
        for kind, directory in (("Incoming", INCOMING_SHEETS_DIR), ("Working", WORKING_SHEETS_DIR), ("Received", RECEIVED_SHEETS_DIR)):
            try:
                directory.mkdir(parents=True, exist_ok=True)
                paths = sorted(directory.glob("*.xlsx"), key=lambda path: path.stat().st_mtime, reverse=True)
            except Exception as error:
                errors.append(f"{kind}: {error}")
                continue
            self.home_sheet_paths[kind] = {path.name: path for path in paths}
            for path in paths:
                key = self._home_sheet_key(kind, path.name)
                try:
                    summary = summarize_workbook(path)
                except Exception as error:
                    errors.append(f"{path.name}: {error}")
                    summary = {"name": path.name, "row_count": 0, "received_count": 0, "purchase_total": 0.0, "all_received": False, "partially_received": False}
                self.home_sheet_summaries[key] = summary
        self._refresh_home_sheet_list()
        self._refresh_home_metrics()
        self.refresh_payouts_tab()
        self._update_home_sheet_tabs()
        if errors:
            self.status_var.set(f"Home refreshed with {len(errors)} sheet issue(s).")
        else:
            self.status_var.set("Home metrics refreshed.")

    def _shared_conflict_files(self) -> list[Path]:
        if not CARD_PIPELINE_DIR.exists():
            return []
        patterns = ("*conflict*.json", "*conflicted*.json", "*copy*.json")
        found: list[Path] = []
        for pattern in patterns:
            found.extend(CARD_PIPELINE_DIR.glob(pattern))
        return sorted({path for path in found if path.name.lower() != PROFIT_LEDGER_PATH.name.lower() and path.name.lower() != SHEET_MARKERS_PATH.name.lower()})

    def _start_startup_refresh(self) -> None:
        self.status_var.set("Loading sheet lists...")
        thread = threading.Thread(target=self._startup_refresh_worker, daemon=True)
        thread.start()

    def _refresh_startup_google_sheet_caches(self) -> dict[str, object]:
        result = {"refreshed": 0, "errors": []}
        sources = self._saved_google_sheet_sources()
        if not sources:
            return result
        try:
            with shared_lock(CARD_PIPELINE_DIR, "google-sheet-cache", self.lucas_identity):
                for source in sources:
                    url = str(source.get("url") or "").strip()
                    path = source.get("path")
                    name = str(source.get("name") or "google-sheet")
                    if not url or not path:
                        continue
                    try:
                        output_path = path_from_source_value(path, ASSIGNMENT_CONFIG_PATH.parent)
                        export_google_sheet_to_xlsx(url, output_path, interactive=False)
                        result["refreshed"] = int(result["refreshed"]) + 1
                    except Exception as error:
                        result["errors"].append(f"{name}: {error}")
        except Exception as error:
            result["errors"].append(str(error))
        return result

    def _refresh_keep_source_registry(self) -> None:
        try:
            self.state.register_keep_note_sources(self._saved_google_keep_sources())
        except Exception:
            return

    def _saved_google_keep_sources(self) -> list[dict[str, object]]:
        try:
            raw = json.loads(ASSIGNMENT_CONFIG_PATH.read_text(encoding="utf-8"))
        except Exception:
            return []
        entries = raw.get("companies", raw) if isinstance(raw, dict) else raw
        if not isinstance(entries, list):
            return []
        output_dir = CARD_PIPELINE_DIR / "ASSIGNMENT RULES" / "KEEP EXPORTS"
        sources: list[dict[str, object]] = []
        seen: set[tuple[str, str]] = set()
        for entry in entries:
            if not isinstance(entry, dict):
                continue
            for source in (
                entry.get("rules") or entry.get("rules_source") or entry.get("rulesSource"),
                entry.get("payout") or entry.get("payout_source") or entry.get("payoutSource"),
            ):
                prepared = self._google_keep_cache_source(source, output_dir)
                if not prepared:
                    continue
                key = (str(prepared.get("url") or ""), str(prepared.get("path") or ""))
                if key in seen:
                    continue
                seen.add(key)
                sources.append(prepared)
        return sources

    def _google_keep_cache_source(self, source: object, output_dir: Path) -> dict[str, object] | None:
        if isinstance(source, dict):
            url = str(source.get("url") or "").strip()
            path = source.get("path") or source.get("file")
            if str(source.get("kind") or "").strip() == "google_keep" and url:
                name = str(source.get("name") or (Path(str(path)).stem if path else "") or "Google Keep note")
                cache_path = Path(str(path)) if path else output_dir / f"{safe_filename(name)}.txt"
                try:
                    cache_path.relative_to(ASSIGNMENT_CONFIG_PATH.parent)
                    cache_path = output_dir / f"{safe_filename(name)}.txt"
                except ValueError:
                    pass
                return {"url": url, "path": str(cache_path), "name": name}
            return None
        raw = normalize_source_value(source)
        if not is_google_keep_url(raw):
            return None
        name = "Google Keep note"
        return {"url": raw, "path": str(output_dir / f"{safe_filename(name)}.txt"), "name": name}

    def _saved_google_sheet_sources(self) -> list[dict[str, object]]:
        try:
            raw = json.loads(ASSIGNMENT_CONFIG_PATH.read_text(encoding="utf-8"))
        except Exception:
            return []
        entries = raw.get("companies", raw) if isinstance(raw, dict) else raw
        if not isinstance(entries, list):
            return []
        output_dir = CARD_PIPELINE_DIR / "ASSIGNMENT RULES" / "SHEET EXPORTS"
        sources: list[dict[str, object]] = []
        seen: set[tuple[str, str]] = set()
        for entry in entries:
            if not isinstance(entry, dict):
                continue
            for source in (
                entry.get("rules") or entry.get("rules_source") or entry.get("rulesSource"),
                entry.get("payout") or entry.get("payout_source") or entry.get("payoutSource"),
            ):
                prepared = self._google_sheet_cache_source(source, output_dir)
                if not prepared:
                    continue
                key = (str(prepared.get("url") or ""), str(prepared.get("path") or ""))
                if key in seen:
                    continue
                seen.add(key)
                sources.append(prepared)
        return sources

    def _google_sheet_cache_source(self, source: object, output_dir: Path) -> dict[str, object] | None:
        if isinstance(source, dict):
            url = str(source.get("url") or "").strip()
            path = source.get("path") or source.get("file")
            if str(source.get("kind") or "").strip() == "google_sheet" and url:
                name = str(source.get("name") or (Path(str(path)).stem if path else "") or "Google Sheet")
                cache_path = Path(str(path)) if path else output_dir / f"{safe_filename(name)}.xlsx"
                return {"url": url, "path": str(cache_path), "name": name}
            return None
        raw = normalize_source_value(source)
        if not raw:
            return None
        if is_google_sheet_url(raw):
            return {"url": raw, "path": str(output_dir / "Google Sheet.xlsx"), "name": "Google Sheet"}
        path = Path(raw).expanduser()
        if not path.is_absolute():
            path = ASSIGNMENT_CONFIG_PATH.parent / path
        if path.suffix.lower() != ".gsheet":
            return None
        try:
            shortcut = load_gsheet_shortcut(path)
            url = gsheet_shortcut_url(shortcut)
        except Exception:
            return None
        if not url:
            return None
        name = str(shortcut.get("name") or path.stem or "google-sheet")
        return {"url": url, "path": str(output_dir / f"{safe_filename(name)}.xlsx"), "name": name}

    def _startup_refresh_worker(self) -> None:
        payload = {
            "working_paths": {},
            "received_paths": {},
            "incoming_index": {},
            "incoming_path_count": 0,
            "home_paths": {"Incoming": {}, "Working": {}, "Received": {}},
            "home_summaries": {},
            "google_sheet_cache": {"refreshed": 0, "errors": []},
            "errors": [],
        }
        errors: list[str] = payload["errors"]
        google_cache_result = self._refresh_startup_google_sheet_caches()
        payload["google_sheet_cache"] = google_cache_result
        errors.extend(google_cache_result.get("errors") or [])

        try:
            CARD_PIPELINE_DIR.mkdir(parents=True, exist_ok=True)
            WORKING_SHEETS_DIR.mkdir(parents=True, exist_ok=True)
            working_paths = sorted(WORKING_SHEETS_DIR.glob("*.xlsx"), key=lambda path: path.stat().st_mtime, reverse=True)
            payload["working_paths"] = {path.name: path for path in working_paths}
            payload["home_paths"]["Working"] = {path.name: path for path in working_paths}
        except Exception as error:
            errors.append(f"Working: {error}")
            working_paths = []

        try:
            RECEIVED_SHEETS_DIR.mkdir(parents=True, exist_ok=True)
            received_paths = sorted(RECEIVED_SHEETS_DIR.glob("*.xlsx"), key=lambda path: path.stat().st_mtime, reverse=True)
            payload["received_paths"] = {path.name: path for path in received_paths}
            payload["home_paths"]["Received"] = {path.name: path for path in received_paths}
        except Exception as error:
            errors.append(f"Received: {error}")
            received_paths = []

        try:
            INCOMING_SHEETS_DIR.mkdir(parents=True, exist_ok=True)
            incoming_paths = sorted(INCOMING_SHEETS_DIR.glob("*.xlsx"), key=lambda path: path.stat().st_mtime, reverse=True)
            payload["home_paths"]["Incoming"] = {path.name: path for path in incoming_paths}
            payload["incoming_path_count"] = len(incoming_paths)
        except Exception as error:
            errors.append(f"Incoming: {error}")
            incoming_paths = []

        index: dict[str, dict[str, object]] = {}
        for path in sorted(incoming_paths, key=lambda path: path.name.lower()):
            try:
                rows = read_simple_spreadsheet(path)
            except Exception as error:
                errors.append(f"{path.name}: {error}")
                continue
            for row in rows:
                cert = scan_to_cert(row.get("cert_number"))
                if not cert or cert in index:
                    continue
                index[cert] = {
                    "sheet": path.name,
                    "path": path,
                    "card_title": row.get("card_title") or "",
                    "grader": row.get("grader") or "",
                    "purchase_price": row.get("purchase_price"),
                    "card_ladder_value": row.get("card_ladder_value"),
                    "card_ladder_comps_average": row.get("card_ladder_comps_average"),
                    "cy_value": row.get("cy_value"),
                    "cy_confidence": row.get("cy_confidence"),
                    "card_ladder_comps": row.get("card_ladder_comps") or "",
                    "best_company": row.get("best_company") or "",
                    "estimated_payout": row.get("estimated_payout"),
                }
        payload["incoming_index"] = index

        for kind, paths in (("Incoming", incoming_paths), ("Working", working_paths), ("Received", received_paths)):
            for path in paths:
                key = self._home_sheet_key(kind, path.name)
                try:
                    summary = summarize_workbook(path)
                except Exception as error:
                    errors.append(f"{path.name}: {error}")
                    summary = {"name": path.name, "row_count": 0, "received_count": 0, "purchase_total": 0.0, "all_received": False, "partially_received": False}
                payload["home_summaries"][key] = summary

        self.events.put(("startup_refresh", payload))

    def _apply_startup_refresh(self, payload: dict[str, object]) -> None:
        self.working_sheet_paths = dict(payload.get("working_paths") or {})
        if hasattr(self, "working_sheet_list"):
            self.working_sheet_list.delete(0, tk.END)
            for name in self.working_sheet_paths:
                self.working_sheet_list.insert(tk.END, name)
        if self.working_sheet_paths and self.selected_working_sheet.get() not in self.working_sheet_paths:
            self.selected_working_sheet.set(next(iter(self.working_sheet_paths)))
        self._select_working_sheet_in_list()

        self.received_sheet_paths = dict(payload.get("received_paths") or {})
        if hasattr(self, "received_sheet_combo"):
            received_names = list(self.received_sheet_paths)
            self.received_sheet_combo["values"] = received_names
            if received_names and self.selected_received_sheet.get() not in self.received_sheet_paths:
                self.selected_received_sheet.set(received_names[0])
            elif not received_names:
                self.selected_received_sheet.set("")

        self.incoming_cert_index = dict(payload.get("incoming_index") or {})
        self._match_all_review_rows()
        self._refresh_table()
        self.review_status.set(f"Indexed {len(self.incoming_cert_index)} cert(s) from {int(payload.get('incoming_path_count') or 0)} incoming sheet(s).")

        self.home_sheet_paths = dict(payload.get("home_paths") or {"Incoming": {}, "Working": {}, "Received": {}})
        self.home_sheet_summaries = dict(payload.get("home_summaries") or {})
        self._refresh_home_sheet_list()
        self._refresh_home_metrics()
        self.refresh_payouts_tab()
        self._update_home_sheet_tabs()
        google_cache = payload.get("google_sheet_cache") if isinstance(payload.get("google_sheet_cache"), dict) else {}
        refreshed_google_sheets = int((google_cache or {}).get("refreshed") or 0)
        self._refresh_keep_source_registry()
        if refreshed_google_sheets:
            self.assignment_engine = AssignmentEngine.load()
            self.assignment_config_status.set(self._assignment_config_status())

        errors = list(payload.get("errors") or [])
        if errors:
            self.status_var.set(f"Startup sheet refresh finished with {len(errors)} issue(s).")
        elif refreshed_google_sheets:
            self.status_var.set(f"Sheet lists loaded. Refreshed {refreshed_google_sheets} Google Sheet cache(s).")
        else:
            self.status_var.set("Sheet lists loaded.")

    def _refresh_home_sheet_list(self) -> None:
        if not hasattr(self, "home_sheet_list"):
            return
        kind = self.home_sheet_kind.get()
        self.home_sheet_list.delete(0, tk.END)
        for name in self._filtered_home_sheet_names(kind):
            self.home_sheet_list.insert(tk.END, name)
        if self.home_sheet_list.size():
            self.home_sheet_list.selection_set(0)
            self._load_home_selected_marker()
        else:
            self.home_selected_sheet_key = ""

    def _refresh_home_metrics(self) -> None:
        if not hasattr(self, "incoming_volume_tree"):
            return
        for tree in (self.incoming_volume_tree, self.partial_received_tree):
            tree.delete(*tree.get_children())
        incoming_names = self._filtered_home_sheet_names("Incoming")
        total_cards = 0
        total_received = 0
        total_volume = 0.0
        for name in incoming_names:
            key = self._home_sheet_key("Incoming", name)
            summary = self.home_sheet_summaries.get(key, {})
            marker = self.home_sheet_markers.get(key, {})
            total = int(summary.get("row_count") or 0)
            received = int(summary.get("received_count") or 0)
            volume = float(summary.get("purchase_total") or 0.0)
            total_cards += total
            total_received += received
            total_volume += volume
            self.incoming_volume_tree.insert(
                "",
                tk.END,
                values=(
                    name,
                    str(marker.get("assigned_person") or ""),
                    total,
                    received,
                    format_money(volume),
                    self._incoming_sheet_status(marker, summary),
                ),
            )
            if summary.get("partially_received"):
                self.partial_received_tree.insert(
                    "",
                    tk.END,
                    tags=("partial_sheet",),
                    values=(
                        name,
                        f"{received}/{total}",
                        format_money(volume),
                        str(marker.get("assigned_person") or ""),
                        str(marker.get("tracking_number") or ""),
                        "Yes" if marker.get("all_received") else "",
                    ),
                )
        if incoming_names:
            self.incoming_volume_tree.insert(
                "",
                tk.END,
                tags=("total_divider",),
                values=("━━━━━━", "━━━━━━", "━━━━━━", "━━━━━━", "━━━━━━", "━━━━━━"),
            )
            self.incoming_volume_tree.insert(
                "",
                tk.END,
                tags=("total_row",),
                values=("TOTAL", "", total_cards, total_received, format_money(total_volume), ""),
            )

    def _incoming_sheet_status(self, marker: dict[str, object], summary: dict[str, object]) -> str:
        has_tracking = bool(str(marker.get("tracking_number") or "").strip())
        received = int(summary.get("received_count") or 0)
        if has_tracking or received:
            return "Awaiting Receive"
        return "Awaiting tracking"

    def refresh_payouts_tab(self) -> None:
        if not hasattr(self, "payout_summary_tree"):
            return
        self._refresh_person_combo_values()
        self.payout_summary_tree.delete(*self.payout_summary_tree.get_children())
        self.payout_detail_tree.delete(*self.payout_detail_tree.get_children())
        self.payout_summary_people = {}
        self.payout_detail_keys = {}

        balances: dict[str, dict[str, float | int]] = {}
        detail_count = 0
        filter_person = self.payout_person_var.get().strip().lower()
        for item in self._payout_sheet_items():
            person = item["person"] or "Unassigned"
            if filter_person and filter_person not in person.lower():
                continue
            if not item["paid"]:
                balance = balances.setdefault(person, {"sheets": 0, "cards": 0, "balance": 0.0})
                balance["sheets"] = int(balance["sheets"]) + 1
                balance["cards"] = int(balance["cards"]) + int(item["row_count"])
                balance["balance"] = float(balance["balance"]) + float(item["payout_balance"])
            iid = f"payout:{detail_count}"
            self.payout_detail_keys[iid] = str(item["key"])
            self.payout_detail_tree.insert(
                "",
                tk.END,
                iid=iid,
                values=(
                    item["name"],
                    item["stage"],
                    item["person"],
                    item["row_count"],
                    f"{item['received_count']}/{item['row_count']}",
                    format_money(float(item["payout_balance"])),
                    item["status"],
                ),
            )
            detail_count += 1

        for index, (person, values) in enumerate(sorted(balances.items(), key=lambda pair: (-float(pair[1]["balance"]), pair[0].lower()))):
            iid = f"payout-summary:{index}"
            self.payout_summary_people[iid] = person
            self.payout_summary_tree.insert(
                "",
                tk.END,
                iid=iid,
                values=(
                    person,
                    int(values["sheets"]),
                    int(values["cards"]),
                    format_money(float(values["balance"])),
                ),
            )

        total_balance = sum(float(values["balance"]) for values in balances.values())
        total_sheets = sum(int(values["sheets"]) for values in balances.values())
        total_cards = sum(int(values["cards"]) for values in balances.values())
        if balances:
            self.payout_summary_tree.insert(
                "",
                tk.END,
                tags=("total_divider",),
                values=("━━━━━━", "━━━━━━", "━━━━━━", "━━━━━━"),
            )
            self.payout_summary_tree.insert(
                "",
                tk.END,
                tags=("total_row",),
                values=("TOTAL", total_sheets, total_cards, format_money(total_balance)),
            )
        filter_label = self.payout_person_var.get().strip()
        suffix = f" | Filter: {filter_label}" if filter_label else ""
        self.payout_status_var.set(f"{detail_count} payment sheet(s) | Active balance: {format_money(total_balance)}{suffix}")

    def _payout_sheet_items(self) -> list[dict[str, object]]:
        items: list[dict[str, object]] = []
        seller_names = self._seller_terms_seller_names()
        realized_profit_groups = self._realized_profit_groups_by_person_sheet()
        for stage in ("Incoming", "Received"):
            for name in self.home_sheet_paths.get(stage, {}):
                key = self._home_sheet_key(stage, name)
                marker = self.home_sheet_markers.get(key, {})
                summary = self.home_sheet_summaries.get(key, {})
                paid = bool(marker.get("paid"))
                row_count = int(summary.get("row_count") or 0)
                received_count = int(summary.get("received_count") or 0)
                if stage == "Received":
                    received_count = int(summary.get("received_count") or row_count)
                status = "Paid" if paid else self._payout_sheet_status(stage, marker, summary)
                person = str(marker.get("assigned_person") or "").strip()
                person_key = person.lower()
                is_seller_payout = bool(person_key and person_key in seller_names)
                if not is_seller_payout or stage != "Received":
                    continue
                purchase_total = float(summary.get("purchase_total") or 0.0)
                estimated_payout_total = float(summary.get("estimated_payout_total") or 0.0)
                realized_profit_total = float(realized_profit_groups.get((person.lower(), Path(name).name.lower()), {}).get("profit") or 0.0)
                payout_balance, payout_basis = self._active_payout_balance(
                    person,
                    purchase_total,
                    estimated_payout_total,
                    seller_names,
                    realized_profit_total=realized_profit_total,
                )
                items.append(
                    {
                        "key": key,
                        "stage": stage,
                        "name": name,
                        "person": person,
                        "paid": paid,
                        "row_count": row_count,
                        "received_count": received_count,
                        "purchase_total": purchase_total,
                        "estimated_payout_total": estimated_payout_total,
                        "estimated_profit": round(estimated_payout_total - purchase_total, 2),
                        "realized_profit_total": round(realized_profit_total, 2),
                        "payout_balance": payout_balance,
                        "payout_basis": payout_basis,
                        "status": status,
                    }
                )
        for (person_key, source_key), group in sorted(realized_profit_groups.items(), key=lambda pair: (pair[0][0], pair[0][1])):
            if person_key in seller_names:
                continue
            realized_profit_total = float(group.get("profit") or 0.0)
            if realized_profit_total <= 0:
                continue
            person = str(group.get("person") or "").strip() or "Unassigned"
            source_sheet = str(group.get("source_sheet") or "").strip() or "Sold Cards"
            key = self._sold_payout_key(person, source_sheet)
            marker = self.home_sheet_markers.get(key, {})
            paid = bool(marker.get("paid"))
            payout_balance, payout_basis = self._active_payout_balance(
                person,
                float(group.get("purchase_total") or 0.0),
                float(group.get("sale_total") or 0.0),
                seller_names,
                realized_profit_total=realized_profit_total,
            )
            items.append(
                {
                    "key": key,
                    "stage": "Sold",
                    "name": source_sheet,
                    "person": person,
                    "paid": paid,
                    "row_count": int(group.get("row_count") or 0),
                    "received_count": int(group.get("row_count") or 0),
                    "purchase_total": float(group.get("purchase_total") or 0.0),
                    "estimated_payout_total": float(group.get("sale_total") or 0.0),
                    "estimated_profit": round(float(group.get("sale_total") or 0.0) - float(group.get("purchase_total") or 0.0), 2),
                    "realized_profit_total": round(realized_profit_total, 2),
                    "payout_balance": payout_balance,
                    "payout_basis": payout_basis,
                    "status": "Paid" if paid else "Sold",
                }
            )
        return items

    def _realized_profit_totals_by_person_sheet(self) -> dict[tuple[str, str], float]:
        return {
            key: float(group.get("profit") or 0.0)
            for key, group in self._realized_profit_groups_by_person_sheet().items()
        }

    def _realized_profit_groups_by_person_sheet(self) -> dict[tuple[str, str], dict[str, object]]:
        totals: dict[tuple[str, str], float] = defaultdict(float)
        groups: dict[tuple[str, str], dict[str, object]] = {}
        for record in self._enrich_profit_records_with_people(self._load_profit_ledger()):
            person = str(record.get("assigned_person") or "").strip()
            source_sheet = Path(str(record.get("source_sheet") or "")).name.strip()
            profit = self._money_value(record.get("profit"))
            if not person or not source_sheet or profit is None:
                continue
            key = (person.lower(), source_sheet.lower())
            group = groups.setdefault(
                key,
                {
                    "person": person,
                    "source_sheet": source_sheet,
                    "row_count": 0,
                    "purchase_total": 0.0,
                    "sale_total": 0.0,
                    "profit": 0.0,
                },
            )
            purchase = self._money_value(record.get("purchase_price")) or 0.0
            sale = self._money_value(record.get("sale_price")) or 0.0
            group["row_count"] = int(group["row_count"]) + 1
            group["purchase_total"] = float(group["purchase_total"]) + purchase
            group["sale_total"] = float(group["sale_total"]) + sale
            totals[key] += profit
            group["profit"] = totals[key]
        return groups

    def _sold_payout_key(self, person: str, source_sheet: str) -> str:
        return self._home_sheet_key("Sold", f"{str(person or '').strip()}|{Path(str(source_sheet or '')).name}")

    def _split_sold_payout_name(self, name: str) -> tuple[str, str]:
        if "|" not in name:
            return "", name
        person, source_sheet = name.split("|", 1)
        return person, source_sheet

    def _seller_terms_seller_names(self) -> set[str]:
        return {
            str(term.get("seller") or "").strip().lower()
            for term in self._load_seller_terms()
            if str(term.get("seller") or "").strip()
        }

    def _active_payout_balance(
        self,
        person: str,
        purchase_total: float,
        estimated_payout_total: float,
        seller_names: set[str] | None = None,
        realized_profit_total: float | None = None,
    ) -> tuple[float, str]:
        normalized_person = str(person or "").strip().lower()
        seller_names = seller_names if seller_names is not None else self._seller_terms_seller_names()
        if normalized_person and normalized_person in seller_names:
            return round(float(purchase_total or 0.0), 2), "Seller purchase total"
        realized_profit = float(realized_profit_total or 0.0)
        return max(0.0, round(realized_profit / 2.0, 2)), "Team half sold profit"

    def _payout_sheet_status(self, stage: str, marker: dict[str, object], summary: dict[str, object]) -> str:
        received_count = int(summary.get("received_count") or 0)
        if stage == "Received" or marker.get("all_received") or summary.get("all_received"):
            return "Unpaid"
        if received_count:
            return "Partially Received"
        return "Unreceived"

    def _network_mode_enabled(self) -> bool:
        var = getattr(self, "create_network_mode_var", None)
        return bool(var.get()) if var is not None else False

    def _set_create_network_controls_visible(self, visible: bool) -> None:
        widgets = (
            getattr(self, "network_seller_label", None),
            getattr(self, "seller_terms_seller_combo", None),
            getattr(self, "network_sheet_type_label", None),
            getattr(self, "seller_terms_sheet_type_combo", None),
        )
        for widget in widgets:
            if widget is None:
                continue
            if visible:
                widget.grid()
            else:
                widget.grid_remove()

    def _toggle_create_network_mode(self) -> None:
        enabled = self._network_mode_enabled()
        settings = load_app_settings()
        settings["network_mode"] = enabled
        save_app_settings(settings)
        self.app_settings = settings
        self._set_create_network_controls_visible(enabled)
        if not enabled:
            restored = self._restore_create_seller_term_prices()
            if restored:
                self._refresh_table()
            self.status_var.set("Network Mode off. Seller terms hidden.")
        else:
            self.status_var.set("Network Mode on. Seller and Sheet Type are available in Create.")
            self.apply_create_seller_terms(show_status=False)

    def _load_seller_terms(self) -> list[dict[str, object]]:
        if not SELLER_TERMS_PATH.exists():
            return []
        try:
            with SELLER_TERMS_PATH.open("r", encoding="utf-8-sig", newline="") as handle:
                rows = list(csv.DictReader(handle))
        except Exception:
            return []
        terms: list[dict[str, object]] = []
        for row in rows:
            normalized = {re.sub(r"[^a-z0-9]+", "", str(key or "").lower()): value for key, value in row.items()}
            seller = str(normalized.get("seller") or normalized.get("person") or normalized.get("name") or "").strip()
            sheet_type = str(normalized.get("sheettype") or normalized.get("type") or normalized.get("company") or "").strip()
            value_source = str(normalized.get("valuesource") or normalized.get("source") or "").strip()
            rate = self._seller_terms_rate(normalized.get("sellerrate") or normalized.get("rate") or normalized.get("payout") or normalized.get("percentage"))
            deduction = self._seller_terms_rate(normalized.get("deduction") or normalized.get("sellerdeduction") or normalized.get("deductionpercent") or normalized.get("deductionpercentage"))
            if seller and sheet_type and (rate is not None or deduction is not None):
                terms.append({"seller": seller, "sheet_type": sheet_type, "value_source": value_source, "rate": rate, "deduction": deduction})
        return terms

    def _refresh_seller_terms_dropdowns(self) -> None:
        if hasattr(self, "seller_terms_sheet_type_combo"):
            sheet_types = sorted({str(term.get("sheet_type") or "") for term in self._load_seller_terms() if term.get("sheet_type")}, key=str.lower)
            self.seller_terms_sheet_type_combo["values"] = sheet_types

    def _seller_terms_rate(self, value: object) -> float | None:
        raw = str(value or "").strip()
        text = raw.replace("%", "").strip()
        if not text:
            return None
        try:
            numeric = float(text)
        except ValueError:
            return None
        rate = numeric / 100 if "%" in raw or numeric > 1 else numeric
        return rate if rate >= 0 else None

    def _seller_terms_match(self, seller: str, sheet_type: str) -> dict[str, object] | None:
        seller_key = seller.strip().lower()
        type_key = sheet_type.strip().lower()
        if not seller_key or not type_key:
            return None
        for term in self._load_seller_terms():
            if str(term.get("seller") or "").strip().lower() == seller_key and str(term.get("sheet_type") or "").strip().lower() == type_key:
                return term
        return None

    def _seller_terms_company_price(self, row: WorkbookRow, company_name: str, rate: float | None = None, deduction: float | None = None) -> float | None:
        company_key = company_name.strip().lower()
        if not company_key:
            return None
        for decision in self.assignment_engine.evaluate(row):
            if decision.company.strip().lower() != company_key:
                continue
            if decision.source_value is None:
                return None
            if deduction is not None:
                if decision.payout is None:
                    return None
                return max(0.0, round(decision.payout - (decision.source_value * deduction), 2))
            if rate is not None:
                return round(decision.source_value * rate, 2)
            return None
        return None

    def _restore_create_seller_term_prices(self) -> int:
        restored = 0
        for row in self.intake_rows:
            if hasattr(row, "_seller_terms_base_purchase"):
                base_value = getattr(row, "_seller_terms_base_purchase")
                if row.existing_value != base_value:
                    row.existing_value = base_value
                    restored += 1
        return restored

    def apply_create_seller_terms(self, show_status: bool = True) -> int:
        if not self._network_mode_enabled():
            restored = self._restore_create_seller_term_prices()
            if restored:
                self._refresh_table()
            return 0
        seller = self.seller_terms_seller_var.get().strip() if hasattr(self, "seller_terms_seller_var") else ""
        sheet_type = self.seller_terms_sheet_type_var.get().strip() if hasattr(self, "seller_terms_sheet_type_var") else ""
        if not seller or not sheet_type:
            restored = self._restore_create_seller_term_prices()
            if restored:
                self._refresh_table()
            if show_status and (seller or sheet_type):
                self.status_var.set("Seller terms need both Seller and Sheet Type.")
            return 0
        term = self._seller_terms_match(seller, sheet_type)
        if not term:
            restored = self._restore_create_seller_term_prices()
            if restored:
                self._refresh_table()
            if show_status:
                self.status_var.set(f"No seller terms found for {seller} / {sheet_type}.")
            return 0
        rate = self._money_value(term.get("rate"))
        deduction = self._money_value(term.get("deduction"))
        if deduction is None and rate is None:
            if show_status:
                self.status_var.set(f"Seller terms for {seller} / {sheet_type} need either a Deduction or Seller Rate.")
            return 0
        changed = 0
        skipped = 0
        for row in self.intake_rows:
            if not hasattr(row, "_seller_terms_base_purchase"):
                setattr(row, "_seller_terms_base_purchase", row.existing_value)
            if deduction is not None:
                seller_price = self._seller_terms_company_price(row, sheet_type, deduction=deduction)
            else:
                seller_price = self._seller_terms_company_price(row, sheet_type, rate=rate)
            if seller_price is None:
                skipped += 1
                continue
            if row.existing_value != seller_price:
                row.existing_value = seller_price
                changed += 1
        if changed:
            self._refresh_table()
        if show_status:
            if deduction is not None:
                suffix = f" ({skipped} skipped: no matching {sheet_type} payout)" if skipped else ""
                self.status_var.set(f"Applied seller terms: {seller} / {sheet_type} payout minus {deduction:.0%}.{suffix}")
            else:
                suffix = f" ({skipped} skipped: no matching {sheet_type} source value)" if skipped else ""
                self.status_var.set(f"Applied seller terms: {seller} / {sheet_type} at {rate:.0%} of {sheet_type} rule value.{suffix}")
        return changed

    def _known_assigned_people(self) -> list[str]:
        people = {
            str(marker.get("assigned_person") or "").strip()
            for marker in self.home_sheet_markers.values()
            if str(marker.get("assigned_person") or "").strip()
        }
        return sorted(people, key=str.lower)

    def _known_people(self) -> list[str]:
        people_set = set(self._known_assigned_people())
        if hasattr(self, "profit_rows"):
            people_set.update(
                str(record.get("assigned_person") or "").strip()
                for record in self.profit_rows
                if str(record.get("assigned_person") or "").strip()
            )
        if hasattr(self, "inventory_rows"):
            people_set.update(
                str(record.get("assigned_person") or "").strip()
                for record in self.inventory_rows
                if str(record.get("assigned_person") or "").strip()
            )
        for record in self._load_profit_ledger():
            person = str(record.get("assigned_person") or record.get("person") or "").strip()
            if person:
                people_set.add(person)
        for record in self._load_inventory_ledger():
            person = str(record.get("assigned_person") or record.get("person") or "").strip()
            if person:
                people_set.add(person)
        return sorted(people_set, key=str.lower)

    def _refresh_person_combo_values(self, filter_text: str = "") -> None:
        people = self._known_people()
        if filter_text:
            needle = filter_text.strip().lower()
            people = [person for person in people if needle in person.lower()]
        if hasattr(self, "payout_person_combo"):
            self.payout_person_combo["values"] = people
        if hasattr(self, "home_person_combo"):
            self.home_person_combo["values"] = people
        if hasattr(self, "profit_person_combo"):
            self.profit_person_combo["values"] = people
        if hasattr(self, "inventory_person_combo"):
            self.inventory_person_combo["values"] = people
        if hasattr(self, "seller_terms_seller_combo"):
            self.seller_terms_seller_combo["values"] = people

    def _bind_person_autocomplete(self, combo: ttk.Combobox, refresh_callback=None) -> None:
        combo["values"] = self._known_people()
        combo.configure(postcommand=lambda widget=combo: self._refresh_person_combo_widget(widget))
        combo.bind("<FocusIn>", lambda _event, widget=combo: self._refresh_person_combo_widget(widget), add="+")
        combo.bind("<KeyRelease>", lambda event, widget=combo: self._filter_person_combo(widget, event, refresh_callback=refresh_callback), add="+")

    def _refresh_person_combo_widget(self, combo: ttk.Combobox) -> None:
        typed = combo.get().strip().lower()
        people = self._known_people()
        if typed:
            people = [person for person in people if typed in person.lower()]
        combo["values"] = people

    def _filter_person_combo(self, combo: ttk.Combobox, event, refresh_callback=None) -> None:
        if event.keysym in {"Up", "Down", "Left", "Right", "Return", "KP_Enter", "Escape", "Tab"}:
            return
        self._refresh_person_combo_widget(combo)
        if refresh_callback:
            refresh_callback()

    def delete_person_records(self, person: str) -> dict[str, int]:
        target = person.strip().lower()
        counts = {"markers": 0, "inventory": 0, "profit": 0}
        if not target:
            return counts
        for marker in self.home_sheet_markers.values():
            if str(marker.get("assigned_person") or "").strip().lower() == target:
                marker["assigned_person"] = ""
                counts["markers"] += 1
        if counts["markers"]:
            self._save_sheet_markers()

        inventory_rows = [self._normalize_inventory_record(record) for record in self._load_inventory_ledger()]
        for record in inventory_rows:
            if str(record.get("assigned_person") or "").strip().lower() == target:
                record["assigned_person"] = "Unassigned"
                counts["inventory"] += 1
        if counts["inventory"]:
            self._save_inventory_ledger(inventory_rows)

        profit_rows = [self._normalize_profit_record(record) for record in self._load_profit_ledger()]
        for record in profit_rows:
            if str(record.get("assigned_person") or "").strip().lower() == target:
                record["assigned_person"] = ""
                counts["profit"] += 1
        if counts["profit"]:
            self._save_profit_ledger(profit_rows)
        return counts

    def open_delete_person_dialog(self) -> None:
        people = [person for person in self._known_people() if person.lower() != "unassigned"]
        selected = self.payout_person_var.get().strip() if hasattr(self, "payout_person_var") else ""
        if selected and selected not in people and selected.lower() != "unassigned":
            people.insert(0, selected)
        if not people:
            messagebox.showinfo("No people", "No assigned people are available to delete.")
            return
        person_var = tk.StringVar(value=selected if selected in people else people[0])
        popup = tk.Toplevel(self)
        popup.title("Delete Person")
        popup.configure(bg="#1f1f1f")
        popup.transient(self)
        popup.grab_set()
        popup.resizable(False, False)

        frame = ttk.Frame(popup, style="Panel.TFrame", padding=(18, 16))
        frame.pack(fill=tk.BOTH, expand=True)
        ttk.Label(frame, text="Delete Person", style="Panel.TLabel", font=("Segoe UI Semibold", 12)).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 12))
        ttk.Label(frame, text="Person", style="Panel.TLabel").grid(row=1, column=0, sticky="w", padx=(0, 10), pady=(0, 12))
        combo = ttk.Combobox(frame, textvariable=person_var, values=people, width=34)
        combo.grid(row=1, column=1, sticky="ew", pady=(0, 12))

        def apply_delete() -> None:
            person = person_var.get().strip()
            if not person:
                return
            confirmed = messagebox.askyesno(
                "Delete person?",
                f"Remove {person} from sheet assignments, inventory, and profit records?",
                parent=popup,
            )
            if not confirmed:
                return
            counts = self.delete_person_records(person)
            for var in (self.payout_person_var, self.inventory_person_var, self.profit_person_var):
                if var.get().strip().lower() == person.lower():
                    var.set("")
            self.refresh_home()
            self.refresh_inventory_tab()
            self.refresh_profit_tab()
            self.refresh_payouts_tab()
            popup.destroy()
            self.status_var.set(
                f"Deleted person {person}: {counts['markers']} sheet marker(s), {counts['inventory']} inventory row(s), {counts['profit']} profit row(s)."
            )

        buttons = ttk.Frame(frame, style="Panel.TFrame")
        buttons.grid(row=2, column=0, columnspan=2, sticky="e")
        ttk.Button(buttons, text="Cancel", command=popup.destroy, style="Soft.TButton").pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(buttons, text="Delete", command=apply_delete, style="Primary.TButton").pack(side=tk.LEFT)
        frame.columnconfigure(1, weight=1)
        popup.update_idletasks()
        x = self.winfo_rootx() + max(80, (self.winfo_width() - popup.winfo_width()) // 2)
        y = self.winfo_rooty() + max(80, (self.winfo_height() - popup.winfo_height()) // 2)
        popup.geometry(f"+{x}+{y}")

    def _selected_payout_keys(self) -> list[str]:
        if not hasattr(self, "payout_detail_tree"):
            return []
        return [self.payout_detail_keys.get(iid, "") for iid in self.payout_detail_tree.selection() if self.payout_detail_keys.get(iid)]

    def mark_payout_person_paid(self, event=None) -> None:
        if event is not None:
            row_id = self.payout_summary_tree.identify_row(event.y)
            if not row_id:
                return
            self.payout_summary_tree.selection_set(row_id)
        selected = self.payout_summary_tree.selection()
        if not selected:
            return
        person = self.payout_summary_people.get(selected[0])
        if not person:
            return
        if person in {"TOTAL", "━━━━━━"}:
            return
        matching_items = [
            item
            for item in self._payout_sheet_items()
            if not item["paid"] and (item["person"] or "Unassigned") == person
        ]
        if not matching_items:
            self.payout_status_var.set(f"No unpaid sheets found for {person}.")
            return
        total_balance = sum(float(item["payout_balance"]) for item in matching_items)
        total_cards = sum(int(item["row_count"]) for item in matching_items)
        self.open_mark_payout_person_paid_popup(person, matching_items, total_cards, total_balance)

    def open_mark_payout_person_paid_popup(
        self,
        person: str,
        matching_items: list[dict[str, object]],
        total_cards: int,
        total_balance: float,
    ) -> None:
        confirmed_var = tk.BooleanVar(value=False)
        popup = tk.Toplevel(self)
        popup.title("Mark Payout Paid")
        popup.configure(bg="#1f1f1f")
        popup.transient(self)
        popup.grab_set()
        popup.resizable(False, False)

        frame = ttk.Frame(popup, style="Panel.TFrame", padding=(18, 16))
        frame.pack(fill=tk.BOTH, expand=True)
        ttk.Label(frame, text="Mark Payout Paid", style="Panel.TLabel", font=("Segoe UI Semibold", 12)).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 2))
        ttk.Label(frame, text=person, style="Muted.TLabel").grid(row=1, column=0, columnspan=2, sticky="w", pady=(0, 14))
        ttk.Label(frame, text="Sheets", style="Panel.TLabel").grid(row=2, column=0, sticky="w", padx=(0, 18), pady=(0, 8))
        ttk.Label(frame, text=str(len(matching_items)), style="Panel.TLabel").grid(row=2, column=1, sticky="w", pady=(0, 8))
        ttk.Label(frame, text="Cards", style="Panel.TLabel").grid(row=3, column=0, sticky="w", padx=(0, 18), pady=(0, 8))
        ttk.Label(frame, text=str(total_cards), style="Panel.TLabel").grid(row=3, column=1, sticky="w", pady=(0, 8))
        ttk.Label(frame, text="Total Balance", style="Panel.TLabel").grid(row=4, column=0, sticky="w", padx=(0, 18), pady=(0, 14))
        ttk.Label(frame, text=format_money(total_balance), style="Panel.TLabel", font=("Segoe UI Semibold", 11)).grid(row=4, column=1, sticky="w", pady=(0, 14))

        buttons = ttk.Frame(frame, style="Panel.TFrame")
        buttons.grid(row=6, column=0, columnspan=2, sticky="e")
        ttk.Button(buttons, text="Cancel", command=popup.destroy, style="Soft.TButton").pack(side=tk.LEFT, padx=(0, 8))
        confirm_button = ttk.Button(
            buttons,
            text="Confirm Mark Paid",
            style="Primary.TButton",
            state=tk.DISABLED,
            command=lambda: self._apply_payout_person_paid(person, matching_items, total_balance, popup),
        )
        confirm_button.pack(side=tk.LEFT)

        def toggle_confirm() -> None:
            confirm_button.configure(state=tk.NORMAL if confirmed_var.get() else tk.DISABLED)

        ttk.Checkbutton(
            frame,
            text="I confirm this person's full active balance has been paid.",
            variable=confirmed_var,
            command=toggle_confirm,
            style="Panel.TCheckbutton",
        ).grid(row=5, column=0, columnspan=2, sticky="w", pady=(0, 14))
        frame.columnconfigure(1, weight=1)
        popup.update_idletasks()
        x = self.winfo_rootx() + max(80, (self.winfo_width() - popup.winfo_width()) // 2)
        y = self.winfo_rooty() + max(80, (self.winfo_height() - popup.winfo_height()) // 2)
        popup.geometry(f"+{x}+{y}")

    def _apply_payout_person_paid(
        self,
        person: str,
        matching_items: list[dict[str, object]],
        total_balance: float,
        popup: tk.Toplevel | None = None,
    ) -> None:
        for item in matching_items:
            key = str(item["key"])
            kind, _name = self._split_home_sheet_key(key)
            marker = dict(self.home_sheet_markers.get(key, {}))
            summary = self.home_sheet_summaries.get(key, {})
            marker["assigned_person"] = str(item["person"] or "").strip()
            marker["paid"] = True
            marker["all_received"] = bool(marker.get("all_received") or summary.get("all_received") or kind == "Received")
            marker["tracking_number"] = str(marker.get("tracking_number") or "")
            self.home_sheet_markers[key] = marker
        self._save_sheet_markers()
        self.refresh_home()
        if popup is not None:
            popup.destroy()
        self.status_var.set(f"Marked {len(matching_items)} sheet(s) paid for {person}: {format_money(total_balance)}.")

    def open_payout_marker_editor(self, event=None) -> None:
        if event is not None:
            row_id = self.payout_detail_tree.identify_row(event.y)
            if not row_id:
                return
            self.payout_detail_tree.selection_set(row_id)
        keys = self._selected_payout_keys()
        if not keys:
            return
        key = keys[0]
        kind, name = self._split_home_sheet_key(key)
        marker = self.home_sheet_markers.get(key, {})
        summary = self.home_sheet_summaries.get(key, {})
        if kind == "Sold":
            key_person, source_sheet = self._split_sold_payout_name(name)
            person = str(marker.get("assigned_person") or key_person).strip()
            display_name = source_sheet
            realized_profit_total = self._realized_profit_totals_by_person_sheet().get((person.lower(), Path(source_sheet).name.lower()), 0.0)
        else:
            person = str(marker.get("assigned_person") or "").strip()
            display_name = name
            realized_profit_total = self._realized_profit_totals_by_person_sheet().get((person.lower(), Path(name).name.lower()), 0.0)
        balance, basis = self._active_payout_balance(
            person,
            float(summary.get("purchase_total") or 0.0),
            float(summary.get("estimated_payout_total") or 0.0),
            realized_profit_total=realized_profit_total,
        )
        paid_var = tk.BooleanVar(value=bool(marker.get("paid")))
        person_var = tk.StringVar(value=str(marker.get("assigned_person") or "").strip())

        popup = tk.Toplevel(self)
        popup.title("Payout Sheet")
        popup.configure(bg="#1f1f1f")
        popup.transient(self)
        popup.grab_set()
        popup.resizable(False, False)

        frame = ttk.Frame(popup, style="Panel.TFrame", padding=(18, 16))
        frame.pack(fill=tk.BOTH, expand=True)
        ttk.Label(frame, text=display_name, style="Panel.TLabel", font=("Segoe UI Semibold", 12)).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 2))
        ttk.Label(frame, text=f"{kind} | Balance: {format_money(balance)} | {basis}", style="Muted.TLabel").grid(row=1, column=0, columnspan=2, sticky="w", pady=(0, 14))
        ttk.Label(frame, text="Assigned Person", style="Panel.TLabel").grid(row=2, column=0, sticky="w", padx=(0, 10), pady=(0, 10))
        person_combo = ttk.Combobox(frame, textvariable=person_var, width=34)
        person_combo.grid(row=2, column=1, sticky="ew", pady=(0, 10))
        self._bind_person_autocomplete(person_combo)
        ttk.Checkbutton(frame, text="Paid", variable=paid_var, style="Panel.TCheckbutton").grid(row=3, column=0, columnspan=2, sticky="w", pady=(0, 14))
        buttons = ttk.Frame(frame, style="Panel.TFrame")
        buttons.grid(row=4, column=0, columnspan=2, sticky="e")
        ttk.Button(buttons, text="Cancel", command=popup.destroy, style="Soft.TButton").pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(
            buttons,
            text="Save",
            command=lambda: self.save_payout_sheet_marker(key, person_var.get().strip(), bool(paid_var.get()), popup),
            style="Primary.TButton",
        ).pack(side=tk.LEFT)
        frame.columnconfigure(1, weight=1)
        popup.update_idletasks()
        x = self.winfo_rootx() + max(80, (self.winfo_width() - popup.winfo_width()) // 2)
        y = self.winfo_rooty() + max(80, (self.winfo_height() - popup.winfo_height()) // 2)
        popup.geometry(f"+{x}+{y}")

    def save_payout_sheet_marker(self, key: str, person: str, paid: bool, popup: tk.Toplevel | None = None) -> None:
        marker = dict(self.home_sheet_markers.get(key, {}))
        kind, _name = self._split_home_sheet_key(key)
        summary = self.home_sheet_summaries.get(key, {})
        marker["assigned_person"] = person.strip()
        marker["paid"] = bool(paid)
        marker["all_received"] = bool(marker.get("all_received") or summary.get("all_received") or kind == "Received")
        marker["tracking_number"] = str(marker.get("tracking_number") or "")
        self.home_sheet_markers[key] = marker
        self._save_sheet_markers()
        self.refresh_home()
        if popup is not None:
            popup.destroy()
        self.status_var.set(f"Updated payout marker for {self._split_home_sheet_key(key)[1]}.")

    def _load_home_selected_marker(self) -> None:
        if not hasattr(self, "home_sheet_list"):
            return
        selected = self.home_sheet_list.curselection()
        if not selected:
            return
        kind = self.home_sheet_kind.get()
        name = str(self.home_sheet_list.get(selected[0]))
        key = self._home_sheet_key(kind, name)
        self.home_selected_sheet_key = key

    def _show_home_sheet_context_menu(self, event) -> str:
        if not hasattr(self, "home_sheet_list"):
            return "break"
        index = self.home_sheet_list.nearest(event.y)
        if index < 0 or index >= self.home_sheet_list.size():
            return "break"
        self.home_sheet_list.selection_clear(0, tk.END)
        self.home_sheet_list.selection_set(index)
        self.home_sheet_list.activate(index)
        self._load_home_selected_marker()
        menu = tk.Menu(self, tearoff=False, bg="#1f1f1f", fg="#ffffff", activebackground="#1ed760", activeforeground="#000000")
        kind, _name = self._split_home_sheet_key(self.home_selected_sheet_key)
        move_menu = tk.Menu(menu, tearoff=False, bg="#1f1f1f", fg="#ffffff", activebackground="#1ed760", activeforeground="#000000")
        for target_stage in ("Incoming", "Working", "Received"):
            move_menu.add_command(
                label=f"Move to {target_stage}",
                command=lambda stage=target_stage: self.move_selected_home_sheet_to_stage(stage),
                state=tk.DISABLED if target_stage == kind else tk.NORMAL,
            )
        menu.add_cascade(label="Move Sheet", menu=move_menu)
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()
        return "break"

    def move_selected_home_sheet_to_stage(self, target_stage: str) -> None:
        if not self.home_selected_sheet_key:
            messagebox.showinfo("Choose sheet", "Choose a sheet on Home before moving.")
            return
        source_stage, name = self._split_home_sheet_key(self.home_selected_sheet_key)
        if source_stage not in {"Incoming", "Working", "Received"} or not name:
            messagebox.showinfo("Cannot move", "Only Incoming, Working, and Received sheets can be moved from Home.")
            return
        if target_stage not in {"Incoming", "Working", "Received"}:
            messagebox.showinfo("Cannot move", "Choose Incoming, Working, or Received.")
            return
        if source_stage == target_stage:
            return
        path = self._sheet_path_for_stage(source_stage, name)
        if not self._sheet_path_is_visible_home_sheet(source_stage, path):
            messagebox.showerror("Move blocked", f"Move is only allowed inside {source_stage} sheets.")
            return
        confirmed = messagebox.askyesno(
            "Move sheet?",
            (
                f"Move this sheet from {source_stage} to {target_stage}?\n\n{name}\n\n"
                "If moving out of Received, L.U.C.A.S will clear received/paid markers and remove company-sheet/profit rows created from this sheet."
            ),
        )
        if not confirmed:
            return
        try:
            with shared_lock(CARD_PIPELINE_DIR, "sheet-stage-move", self.lucas_identity):
                moved_key, cleanup = self._move_home_sheet_to_stage(self.home_selected_sheet_key, target_stage)
                self.home_selected_sheet_key = moved_key
                self.home_sheet_kind.set(target_stage)
                self._save_sheet_markers()
        except Exception as error:
            messagebox.showerror("Move failed", str(error))
            return
        self.refresh_pipeline()
        self.refresh_home()
        cleanup_note = ""
        if cleanup:
            cleanup_note = (
                f" Cleared {cleanup.get('received_rows_cleared', 0)} received mark(s), "
                f"removed {cleanup.get('company_rows_removed', 0)} company row(s), "
                f"and removed {cleanup.get('profit_rows_removed', 0)} profit ledger row(s)."
            )
        self.status_var.set(f"Moved {name} from {source_stage} to {target_stage}.{cleanup_note}")

    def open_sheet_marker_editor(self) -> None:
        if not self.home_selected_sheet_key:
            messagebox.showinfo("Choose sheet", "Choose a sheet on Home before editing markers.")
            return
        kind, name = self._split_home_sheet_key(self.home_selected_sheet_key)
        marker = self.home_sheet_markers.get(self.home_selected_sheet_key, {})
        summary = self.home_sheet_summaries.get(self.home_selected_sheet_key, {})
        incoming_proper_var = tk.BooleanVar(value=(kind == "Incoming"))
        all_received_var = tk.BooleanVar(value=bool(marker.get("all_received") or summary.get("all_received")))
        tracking_var = tk.StringVar(value=str(marker.get("tracking_number") or ""))
        person_var = tk.StringVar(value=str(marker.get("assigned_person") or ""))

        popup = tk.Toplevel(self)
        popup.title("Edit Sheet Markers")
        popup.configure(bg="#1f1f1f")
        popup.transient(self)
        popup.grab_set()
        popup.resizable(False, False)

        frame = ttk.Frame(popup, style="Panel.TFrame", padding=(18, 16))
        frame.pack(fill=tk.BOTH, expand=True)
        ttk.Label(frame, text=name, style="Panel.TLabel", font=("Segoe UI Semibold", 12)).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 2))
        ttk.Label(frame, text=kind, style="Muted.TLabel").grid(row=1, column=0, columnspan=2, sticky="w", pady=(0, 14))
        ttk.Checkbutton(frame, text="Incoming", variable=incoming_proper_var, style="Panel.TCheckbutton").grid(row=2, column=0, columnspan=2, sticky="w", pady=(0, 10))
        ttk.Label(frame, text="Tracking Number", style="Panel.TLabel").grid(row=3, column=0, sticky="w", padx=(0, 10), pady=(0, 10))
        ttk.Entry(frame, textvariable=tracking_var, width=34).grid(row=3, column=1, sticky="ew", pady=(0, 10))
        ttk.Checkbutton(frame, text="All Received", variable=all_received_var, style="Panel.TCheckbutton").grid(row=4, column=0, columnspan=2, sticky="w", pady=(0, 10))
        ttk.Label(frame, text="Assigned Person", style="Panel.TLabel").grid(row=5, column=0, sticky="w", padx=(0, 10), pady=(0, 14))
        person_combo = ttk.Combobox(frame, textvariable=person_var, width=34)
        person_combo.grid(row=5, column=1, sticky="ew", pady=(0, 14))
        self._bind_person_autocomplete(person_combo)
        buttons = ttk.Frame(frame, style="Panel.TFrame")
        buttons.grid(row=6, column=0, columnspan=2, sticky="e")
        ttk.Button(buttons, text="Cancel", command=popup.destroy, style="Soft.TButton").pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(
            buttons,
            text="Save Markers",
            command=lambda: self.save_home_sheet_markers(
                {
                    "incoming_proper": bool(incoming_proper_var.get()),
                    "tracking_number": tracking_var.get().strip(),
                    "all_received": bool(all_received_var.get()),
                    "assigned_person": person_var.get().strip(),
                },
                popup,
            ),
            style="Primary.TButton",
        ).pack(side=tk.LEFT)
        frame.columnconfigure(1, weight=1)
        popup.update_idletasks()
        x = self.winfo_rootx() + max(80, (self.winfo_width() - popup.winfo_width()) // 2)
        y = self.winfo_rooty() + max(80, (self.winfo_height() - popup.winfo_height()) // 2)
        popup.geometry(f"+{x}+{y}")

    def save_home_sheet_markers(self, marker: dict[str, object], popup: tk.Toplevel | None = None) -> None:
        if not self.home_selected_sheet_key:
            messagebox.showinfo("Choose sheet", "Choose a sheet on Home before saving markers.")
            return
        existing_marker = dict(self.home_sheet_markers.get(self.home_selected_sheet_key, {}))
        incoming_proper = bool(marker.get("incoming_proper"))
        old_assigned_person = str(existing_marker.get("assigned_person") or "").strip()
        marker = {
            "paid": bool(existing_marker.get("paid")),
            "tracking_number": str(marker.get("tracking_number") or "").strip(),
            "all_received": bool(marker.get("all_received")),
            "assigned_person": str(marker.get("assigned_person") or "").strip(),
        }
        key = self.home_selected_sheet_key
        moved = False
        inventory_rows_added = 0
        inventory_candidate_rows = 0
        try:
            with shared_lock(CARD_PIPELINE_DIR, "receive-company-sheets", self.lucas_identity):
                selected_kind, _selected_name = self._split_home_sheet_key(key)
                if selected_kind == "Received" and not marker["all_received"]:
                    moved_key = self._move_received_sheet_to_incoming(key)
                    if moved_key:
                        self._delete_sheet_marker(key)
                        key = moved_key
                        self.home_selected_sheet_key = key
                        self.home_sheet_kind.set("Incoming")
                        moved = True
                elif marker["all_received"]:
                    moved_key = self._move_sheet_to_received(key)
                    if moved_key:
                        self._delete_sheet_marker(key)
                        key = moved_key
                        self.home_selected_sheet_key = key
                        moved = True
                elif incoming_proper:
                    moved_key = self._move_working_sheet_to_incoming(key)
                    if moved_key:
                        self._delete_sheet_marker(key)
                        key = moved_key
                        self.home_selected_sheet_key = key
                        self.home_sheet_kind.set("Incoming")
                        moved = True
                if moved:
                    current_kind, _current_name = self._split_home_sheet_key(key)
                    marker = self._marker_for_stage(marker, current_kind)
                self.home_sheet_markers[key] = marker
                _current_kind, current_name = self._split_home_sheet_key(key)
                inventory_rows_reassigned = 0
                if old_assigned_person != str(marker.get("assigned_person") or "").strip():
                    inventory_rows_reassigned = self._retarget_inventory_rows_for_source(current_name, str(marker.get("assigned_person") or ""))
                if _current_kind == "Received" and marker["all_received"] and str(marker.get("assigned_person") or "").strip():
                    inventory_rows_added, inventory_candidate_rows = self._sync_received_sheet_inventory_to_ledger(
                        _current_kind,
                        self._sheet_path_for_stage(_current_kind, current_name),
                        str(marker.get("assigned_person") or ""),
                    )
                self._save_sheet_markers()
        except Exception as error:
            messagebox.showerror("Save failed", str(error))
            return
        self.refresh_working_sheets()
        self.refresh_received_sheets()
        self.refresh_incoming_index()
        self.refresh_home()
        if hasattr(self, "inventory_tree"):
            self.refresh_inventory_tab(enrich=True)
        if popup is not None:
            popup.destroy()
        inventory_note = f" Reassigned {inventory_rows_reassigned} inventory row(s)." if inventory_rows_reassigned else ""
        if inventory_rows_added:
            inventory_note += f" Added {inventory_rows_added} inventory row(s)."
        elif inventory_candidate_rows:
            inventory_note += " Inventory was already up to date."
        self.status_var.set(("Sheet markers saved and moved." if moved else "Sheet markers saved.") + inventory_note)

    def _home_sheet_key(self, kind: str, name: str) -> str:
        return f"{kind}|{name}"

    def _split_home_sheet_key(self, key: str) -> tuple[str, str]:
        if "|" not in key:
            return "", key
        kind, name = key.split("|", 1)
        return kind, name

    def _move_working_sheet_to_incoming(self, key: str) -> str:
        moved_key, _cleanup = self._move_home_sheet_to_stage(key, "Incoming")
        return moved_key

    def _move_sheet_to_received(self, key: str) -> str:
        moved_key, _cleanup = self._move_home_sheet_to_stage(key, "Received")
        return moved_key

    def _move_received_sheet_to_incoming(self, key: str) -> str:
        moved_key, _cleanup = self._move_home_sheet_to_stage(key, "Incoming")
        return moved_key

    def _assign_sheet_to_seller(self, stage: str, sheet_name: str, seller: str) -> bool:
        person = str(seller or "").strip()
        if stage not in {"Working", "Incoming", "Received"} or not sheet_name or not person:
            return False
        key = self._home_sheet_key(stage, sheet_name)
        marker = dict(self._load_sheet_markers().get(key, {}))
        marker.update(dict(self.home_sheet_markers.get(key, {})))
        marker["assigned_person"] = person
        marker = self._marker_for_stage(marker, stage)
        self.home_sheet_markers[key] = marker
        self._save_sheet_markers()
        return True

    def _move_home_sheet_to_stage(self, key: str, target_stage: str) -> tuple[str, dict[str, int]]:
        source_stage, name = self._split_home_sheet_key(key)
        if source_stage not in {"Incoming", "Working", "Received"} or target_stage not in {"Incoming", "Working", "Received"} or not name:
            return "", {}
        if source_stage == target_stage:
            return key, {}
        source = self._sheet_path_for_stage(source_stage, name)
        if not source.exists():
            raise FileNotFoundError(f"{source_stage} sheet not found: {source}")
        target_dir = {
            "Incoming": INCOMING_SHEETS_DIR,
            "Working": WORKING_SHEETS_DIR,
            "Received": RECEIVED_SHEETS_DIR,
        }[target_stage]
        target_dir.mkdir(parents=True, exist_ok=True)
        destination = target_dir / source.name
        if destination.exists():
            raise FileExistsError(f"{target_stage} sheet already exists: {destination.name}")

        cleanup: dict[str, int] = {}
        moving_out_of_received = source_stage == "Received" and target_stage != "Received"
        if moving_out_of_received:
            cleanup = self._cleanup_sheet_received_side_effects(source.name, source)
        shutil.move(str(source), str(destination))

        old_marker = dict(self.home_sheet_markers.get(key, {}))
        self._delete_sheet_marker(key)
        new_key = self._home_sheet_key(target_stage, destination.name)
        self.home_sheet_markers[new_key] = self._marker_for_stage(old_marker, target_stage)
        return new_key, cleanup

    def _marker_for_stage(self, marker: dict[str, object], stage: str) -> dict[str, object]:
        normalized = dict(marker)
        if stage == "Received":
            normalized["all_received"] = True
            normalized.setdefault("received_at", datetime.now().isoformat(timespec="seconds"))
        else:
            normalized["all_received"] = False
            normalized["paid"] = False
            for field in ("received_at", "archived_at", "archived_from"):
                normalized.pop(field, None)
        return normalized

    def _cleanup_sheet_received_side_effects(self, source_sheet_name: str, source_path: Path) -> dict[str, int]:
        clear_result = clear_received_in_workbooks([source_path])
        company_result = remove_company_sheet_rows_for_source(COMPANY_SHEETS_DIR, source_sheet_name)
        profit_rows_removed = self._remove_profit_ledger_rows_for_source(source_sheet_name)
        errors = list(clear_result.get("errors") or []) + list(company_result.get("errors") or [])
        if errors:
            raise RuntimeError("Move cleanup failed: " + "; ".join(str(error) for error in errors[:5]))
        return {
            "received_rows_cleared": int(clear_result.get("rows_cleared") or 0),
            "company_rows_removed": int(company_result.get("rows_removed") or 0),
            "profit_rows_removed": profit_rows_removed,
        }

    def _remove_profit_ledger_rows_for_source(self, source_sheet_name: str) -> int:
        source_name = Path(str(source_sheet_name or "")).name
        if not source_name:
            return 0
        ledger = [self._normalize_profit_record(record) for record in self._load_profit_ledger()]
        kept = [
            record
            for record in ledger
            if Path(str(record.get("source_sheet") or "")).name != source_name
        ]
        removed = len(ledger) - len(kept)
        if removed:
            self._save_profit_ledger(kept)
        return removed

    def _remove_inventory_rows_for_source(self, source_sheet_name: str) -> int:
        source_name = Path(str(source_sheet_name or "")).name.strip().lower()
        if not source_name:
            return 0
        ledger = [self._normalize_inventory_record(record) for record in self._load_inventory_ledger()]
        kept = [
            record
            for record in ledger
            if Path(str(record.get("source_sheet") or "")).name.strip().lower() != source_name
        ]
        removed = len(ledger) - len(kept)
        if removed:
            self._save_inventory_ledger(kept)
        return removed

    def _sheet_path_for_stage(self, kind: str, name: str) -> Path:
        if kind == "Working":
            return self.home_sheet_paths.get("Working", {}).get(name) or WORKING_SHEETS_DIR / name
        if kind == "Incoming":
            return self.home_sheet_paths.get("Incoming", {}).get(name) or INCOMING_SHEETS_DIR / name
        if kind == "Received":
            return self.received_sheet_paths.get(name) or RECEIVED_SHEETS_DIR / name
        return Path(name)

    def _sheet_path_is_visible_home_sheet(self, kind: str, path: Path) -> bool:
        expected = {
            "Incoming": INCOMING_SHEETS_DIR,
            "Working": WORKING_SHEETS_DIR,
            "Received": RECEIVED_SHEETS_DIR,
        }.get(kind)
        if expected is None:
            return False
        try:
            path.resolve().relative_to(expected.resolve())
            return path.suffix.lower() == ".xlsx"
        except Exception:
            return False

    def _move_fully_received_sheets_to_received(self, paths: list[Path]) -> list[str]:
        moved: list[str] = []
        for path in paths:
            if not path.exists():
                continue
            try:
                summary = summarize_workbook(path)
            except Exception:
                continue
            if not summary.get("all_received"):
                continue
            parent = path.parent.resolve()
            kind = "Incoming" if parent == INCOMING_SHEETS_DIR.resolve() else "Working" if parent == WORKING_SHEETS_DIR.resolve() else ""
            if not kind:
                continue
            old_key = self._home_sheet_key(kind, path.name)
            marker = dict(self.home_sheet_markers.get(old_key, {}))
            marker["all_received"] = True
            new_key = self._move_sheet_to_received(old_key)
            if new_key:
                self._delete_sheet_marker(old_key)
                self.home_sheet_markers[new_key] = marker
                moved.append(path.name)
        return moved

    def _load_sheet_markers(self) -> dict[str, dict[str, object]]:
        try:
            if not SHEET_MARKERS_PATH.exists():
                return {}
            raw = json.loads(SHEET_MARKERS_PATH.read_text(encoding="utf-8"))
            if isinstance(raw, dict):
                return {str(key): dict(value) for key, value in raw.items() if isinstance(value, dict)}
        except Exception:
            return {}
        return {}

    def _delete_sheet_marker(self, key: str) -> None:
        if not key:
            return
        self.home_sheet_markers.pop(key, None)
        self.deleted_sheet_marker_keys.add(key)

    def _save_sheet_markers(self) -> None:
        with shared_lock(CARD_PIPELINE_DIR, "sheet-markers", self.lucas_identity):
            latest = self._load_sheet_markers()
            latest.update(self.home_sheet_markers)
            for key in self.deleted_sheet_marker_keys - set(self.home_sheet_markers):
                latest.pop(key, None)
            self.deleted_sheet_marker_keys.clear()
            self.home_sheet_markers = latest
            atomic_write_json(SHEET_MARKERS_PATH, self.home_sheet_markers)

    def _show_mode(self) -> None:
        for child in self.mode_host.winfo_children():
            child.destroy()
        mode = self.input_mode.get()
        if mode == "Barcode Scanner":
            self._build_barcode_mode()
            self.after(100, self._arm_scanner)
        elif mode == "Manual Entry":
            self._build_manual_intake_mode()
        elif mode == "Photo OCR":
            self._build_file_mode(photo=True)
        else:
            self._build_file_mode(photo=False)
        if hasattr(self, "intake_tree"):
            self._refresh_table()

    def _show_review_mode(self) -> None:
        if not hasattr(self, "review_mode_host"):
            return
        for child in self.review_mode_host.winfo_children():
            child.destroy()
        if self.review_mode.get() == "Manual Receive":
            self._build_manual_review_mode()
        else:
            self._build_automatic_review_mode()
        self._refresh_table()

    def _build_manual_review_mode(self) -> None:
        self.review_mode_host.columnconfigure(8, weight=1)
        ttk.Label(self.review_mode_host, text="Double-click cells in the Receive table to enter certs or adjust matched details.", style="Muted.TLabel").grid(row=0, column=0, columnspan=9, sticky="w")

    def _build_manual_intake_mode(self) -> None:
        self.scanning_station_active = False
        self.scan_entry = None
        self.mode_host.columnconfigure(8, weight=1)
        ttk.Label(self.mode_host, text="Use the + Add row line in the Create table, then double-click cells to enter certs, card details, purchase, or CY fields.", style="Muted.TLabel").grid(row=0, column=0, columnspan=9, sticky="w")

    def _build_automatic_review_mode(self) -> None:
        self.review_mode_host.columnconfigure(8, weight=1)
        ttk.Label(self.review_mode_host, text="Input", style="Panel.TLabel").grid(row=0, column=0, sticky="w")
        selector = ttk.Combobox(
            self.review_mode_host,
            textvariable=self.review_input_mode,
            state="readonly",
            values=["Barcode Scanner", "Photo OCR"],
            width=18,
        )
        selector.grid(row=0, column=1, sticky="w", padx=(8, 16))
        selector.bind("<<ComboboxSelected>>", lambda _event: self._show_review_mode())
        if self.review_input_mode.get() == "Photo OCR":
            self._build_review_photo_controls(start_col=2)
        else:
            self._build_review_barcode_controls(start_col=2)

    def _build_review_barcode_controls(self, start_col: int) -> None:
        self.review_station_button = ttk.Button(self.review_mode_host, text="Enter Receive Scanning Mode", command=self.toggle_review_scanning, style="Primary.TButton")
        self.review_station_button.grid(row=0, column=start_col, sticky="w", padx=(0, 14))
        ttk.Label(self.review_mode_host, text="Scan", style="Panel.TLabel").grid(row=0, column=start_col + 1, sticky="w")
        self.review_scan_entry = ttk.Entry(self.review_mode_host, textvariable=self.review_scan_cert, width=28)
        self.review_scan_entry.grid(row=0, column=start_col + 2, sticky="w", padx=(8, 14))
        self.review_scan_entry.bind("<Return>", lambda _event: self.add_review_scanned_row())
        self.review_scan_entry.bind("<KP_Enter>", lambda _event: self.add_review_scanned_row())
        self._set_review_station_controls()
        if self.review_scanning_active:
            self.after(100, self._arm_review_scanner)

    def _build_review_photo_controls(self, start_col: int) -> None:
        self.review_scanning_active = False
        self.review_scan_entry = None
        ttk.Button(self.review_mode_host, text="Add Receive Photos", command=self.add_review_photos, style="Soft.TButton").grid(row=0, column=start_col, sticky="w", padx=(0, 8))
        ttk.Button(self.review_mode_host, text="Scan Receive Photos", command=self.scan_review_photos, style="Primary.TButton").grid(row=0, column=start_col + 1, sticky="w", padx=(0, 8))
        ttk.Button(self.review_mode_host, text="Clear Receive Photos", command=self.clear_review_photos, style="Soft.TButton").grid(row=0, column=start_col + 2, sticky="w")
        ttk.Label(self.review_mode_host, textvariable=self.review_photo_status, style="Muted.TLabel").grid(row=1, column=0, columnspan=9, sticky="w", pady=(10, 0))

    def _build_barcode_mode(self) -> None:
        self.mode_host.columnconfigure(7, weight=1)
        self.station_button = ttk.Button(self.mode_host, text="Enter Scanning Station Mode", command=self.toggle_scanning_station, style="Primary.TButton")
        self.station_button.grid(row=0, column=0, sticky="w", padx=(0, 14))
        ttk.Label(self.mode_host, text="Scan", style="Panel.TLabel").grid(row=0, column=1, sticky="w")
        self.scan_entry = ttk.Entry(self.mode_host, textvariable=self.scan_cert, width=28)
        self.scan_entry.grid(row=0, column=2, sticky="w", padx=(8, 14))
        self.scan_entry.bind("<Return>", lambda _event: self.add_scanned_row())
        self.scan_entry.bind("<KP_Enter>", lambda _event: self.add_scanned_row())
        ttk.Label(self.mode_host, text="Grader", style="Panel.TLabel").grid(row=0, column=3, sticky="w")
        ttk.Combobox(self.mode_host, textvariable=self.scan_grader, values=["PSA", "BGS", "SGC", "CGC"], state="readonly", width=8).grid(row=0, column=4, sticky="w", padx=(8, 14))
        ttk.Label(self.mode_host, textvariable=self.scan_status, style="Muted.TLabel").grid(row=1, column=0, columnspan=8, sticky="w", pady=(10, 0))
        self._set_station_controls()

    def _build_file_mode(self, photo: bool) -> None:
        if photo:
            self._build_photo_mode()
            return
        label = "Photo OCR Export" if photo else "Spreadsheet"
        self.mode_host.columnconfigure(1, weight=1)
        ttk.Label(self.mode_host, text=label, style="Panel.TLabel").grid(row=0, column=0, sticky="w", padx=(0, 8))
        ttk.Entry(self.mode_host, textvariable=self.file_path).grid(row=0, column=1, sticky="ew", padx=(0, 8))
        ttk.Button(self.mode_host, text="Browse", command=self.browse_file, style="Soft.TButton").grid(row=0, column=2, sticky="e")
        ttk.Label(self.mode_host, text="Sheet", style="Panel.TLabel").grid(row=0, column=3, sticky="w", padx=(14, 8))
        self.sheet_combo = ttk.Combobox(self.mode_host, textvariable=self.sheet_name, state="readonly", width=18)
        self.sheet_combo.grid(row=0, column=4, sticky="w")
        ttk.Button(self.mode_host, text="Load Rows", command=self.load_file_rows, style="Primary.TButton").grid(row=0, column=5, sticky="e", padx=(14, 0))

    def _build_photo_mode(self) -> None:
        self.mode_host.columnconfigure(4, weight=1)
        ttk.Button(self.mode_host, text="Add Photos", command=self.add_photos, style="Soft.TButton").grid(row=0, column=0, sticky="w", padx=(0, 8))
        ttk.Button(self.mode_host, text="Add Folder", command=self.add_photo_folder, style="Soft.TButton").grid(row=0, column=1, sticky="w", padx=(0, 8))
        ttk.Button(self.mode_host, text="Scan Photos", command=self.scan_photos, style="Primary.TButton").grid(row=0, column=2, sticky="w", padx=(0, 8))
        ttk.Button(self.mode_host, text="Clear Photos", command=self.clear_photos, style="Soft.TButton").grid(row=0, column=3, sticky="w", padx=(0, 8))
        ttk.Label(self.mode_host, textvariable=self.photo_status, style="Muted.TLabel").grid(row=1, column=0, columnspan=5, sticky="w", pady=(10, 0))

    def add_scanned_row(self) -> None:
        if not self.scanning_station_active:
            self.scan_status.set("Click Enter Scanning Station Mode before scanning.")
            return
        cert = scan_to_cert(self.scan_cert.get())
        if not cert:
            self.scan_status.set("No cert detected. Scan again.")
            self._arm_scanner()
            return
        grader = self.scan_grader.get().strip().upper()
        card = self.scan_card.get().strip()
        added_rows = self._append_rows([
            {
                "cert_number": cert,
                "grader": grader,
                "card_title": card,
                "purchase_price": None,
                "source": "Barcode",
                "notes": "" if cert and grader else "Missing cert or grader",
            }
        ])
        self.scan_cert.set("")
        self.scan_card.set("")
        self.scan_status.set(f"Added row {len(self.intake_rows) + 1}: {cert}. Scanner ready for next cert.")
        self.status_var.set(f"Added scanned card {cert}.")
        if added_rows:
            self._select_excel_row(added_rows[-1])
        self._arm_scanner()

    def add_manual_intake_row(self) -> int | None:
        added_rows = self._append_rows([
            {
                "cert_number": "",
                "grader": "",
                "card_title": "",
                "purchase_price": None,
                "source": "Manual",
                "notes": "Manual create row",
            }
        ])
        if added_rows:
            row_id = str(added_rows[-1])
            self.intake_tree.selection_set(row_id)
            self.intake_tree.focus(row_id)
            self.intake_tree.see(row_id)
            self.status_var.set("Manual create row added. Double-click cells to edit it.")
            return added_rows[-1]
        return None

    def browse_file(self) -> None:
        path = filedialog.askopenfilename(title="Choose workbook", filetypes=[("Excel workbook", "*.xlsx")])
        if not path:
            return
        self.file_path.set(path)
        try:
            names = workbook_sheet_names(Path(path))
            self.sheet_combo["values"] = names
            if names:
                self.sheet_name.set(names[0])
        except Exception as error:
            messagebox.showerror("Workbook error", str(error))

    def load_file_rows(self) -> None:
        path = Path(self.file_path.get())
        if not path.exists():
            messagebox.showinfo("Choose file", "Choose a workbook first.")
            return
        try:
            if self.input_mode.get() == "Photo OCR":
                rows = read_photo_export(path, self.sheet_name.get() or None)
            else:
                rows = read_simple_spreadsheet(path, self.sheet_name.get() or None)
        except Exception as error:
            messagebox.showerror("Load failed", self._create_sheet_load_error(error))
            return
        if not rows:
            messagebox.showinfo("No usable rows", self._create_sheet_no_rows_message(path))
            self.status_var.set(f"No usable rows found in {path.name}.")
            return
        self._append_rows(rows)
        self.status_var.set(f"Loaded {len(rows)} row(s) from {path.name}.")

    def _create_sheet_load_error(self, error: Exception) -> str:
        raw = str(error).strip()
        if isinstance(error, (TypeError, ValueError, KeyError, IndexError)):
            return (
                "This sheet does not match the Create import format.\n\n"
                "Expected either a simple sheet with Cert #, Card Description, and Purchase Price columns, "
                "or a Photo OCR export with certification/card fields.\n\n"
                f"Details: {raw or type(error).__name__}"
            )
        return raw or type(error).__name__

    def _create_sheet_no_rows_message(self, path: Path) -> str:
        return (
            f"No usable card rows were found in {path.name}.\n\n"
            "Create can import a simple sheet with Cert #, Card Description, and Purchase Price columns, "
            "or a Photo OCR export with certification/card fields."
        )

    def toggle_scanning_station(self) -> None:
        self.scanning_station_active = not self.scanning_station_active
        self._set_station_controls()
        if self.scanning_station_active:
            self.scan_status.set("Scanning station armed. Scan certs now; each scan adds the next row.")
            self._arm_scanner()
        else:
            self.scan_status.set("Scanning station is off.")

    def _set_station_controls(self) -> None:
        if not hasattr(self, "station_button"):
            return
        self.station_button.configure(text="Exit Scanning Station Mode" if self.scanning_station_active else "Enter Scanning Station Mode")
        if self.scan_entry is not None:
            self.scan_entry.configure(state=tk.NORMAL if self.scanning_station_active else tk.DISABLED)

    def add_photos(self) -> None:
        paths = filedialog.askopenfilenames(
            title="Choose card photos",
            filetypes=[("Images", "*.png *.jpg *.jpeg *.webp *.bmp"), ("All files", "*.*")],
        )
        self._add_photo_paths([Path(path) for path in paths])

    def add_photo_folder(self) -> None:
        folder = filedialog.askdirectory(title="Choose photo folder")
        if not folder:
            return
        extensions = {".png", ".jpg", ".jpeg", ".webp", ".bmp"}
        self._add_photo_paths([path for path in Path(folder).iterdir() if path.suffix.lower() in extensions])

    def clear_photos(self) -> None:
        if self.photo_worker and self.photo_worker.is_alive():
            messagebox.showinfo("Scan running", "Wait for the photo scan to finish before clearing photos.")
            return
        self.photo_paths = []
        self.photo_status.set("No photos selected.")

    def _add_photo_paths(self, paths: list[Path]) -> None:
        existing = {path.resolve() for path in self.photo_paths if path.exists()}
        added = 0
        for path in paths:
            if not path.exists() or path.resolve() in existing:
                continue
            self.photo_paths.append(path)
            existing.add(path.resolve())
            added += 1
        self.photo_status.set(f"{len(self.photo_paths)} photo(s) selected. Added {added}.")

    def scan_photos(self) -> None:
        if self.photo_worker and self.photo_worker.is_alive():
            messagebox.showinfo("Scan running", "Photo scan is already running.")
            return
        if not self.photo_paths:
            messagebox.showinfo("No photos", "Add photos before scanning.")
            return
        if genai is None or identify_cards_sync is None:
            messagebox.showerror("Missing dependency", "Photo OCR dependencies are not available.")
            return
        self._load_photo_env()
        api_key = os.environ.get("GOOGLE_API_KEY", "").strip()
        if not api_key:
            messagebox.showerror("Missing GOOGLE_API_KEY", "Create .env in the L.U.C.A.S project folder or set GOOGLE_API_KEY.")
            return
        self.photo_client = genai.Client(api_key=api_key)
        self.photo_status.set(f"Scanning 0/{len(self.photo_paths)} photo(s)...")
        self.photo_worker = threading.Thread(target=self._photo_scan_worker, daemon=True)
        self.photo_worker.start()

    def _photo_scan_worker(self) -> None:
        total = len(self.photo_paths)
        detected_total = 0
        for index, path in enumerate(list(self.photo_paths), start=1):
            try:
                self.events.put(("photo_status", f"Scanning {index}/{total}: {path.name}..."))
                image_b64 = base64.b64encode(path.read_bytes()).decode("utf-8")
                cards = identify_cards_sync(
                    self.photo_client,
                    image_b64,
                    progress_callback=lambda message, i=index, n=total, p=path: self.events.put(
                        ("photo_status", f"Scanning {i}/{n}: {p.name} - {message}")
                    ),
                )
                rows = [self._photo_card_to_row(path, card) for card in cards if self._photo_card_has_inventory(card)]
                detected_total += len(rows)
                self.events.put(("photo_rows", rows))
                self.events.put(("photo_status", f"Scanning {index}/{total}: {path.name} -> {len(rows)} card row(s)."))
            except (TemporaryModelUnavailable, ModelQuotaExceeded, ModelResponseParseError) as error:
                self.events.put(("photo_status", f"{path.name}: {error}"))
            except Exception as error:
                self.events.put(("photo_status", f"{path.name}: {error}"))
        self.events.put(("photo_status", f"Photo scan complete. Added {detected_total} card row(s)."))

    def _photo_card_to_row(self, path: Path, card: dict) -> dict[str, object]:
        grader = normalize_grader(card.get("grading_company"))
        title = build_card_title(
            {
                "description": "",
                "year": card.get("year"),
                "set": card.get("set"),
                "player": card.get("player"),
                "card_number": card.get("card_number"),
                "parallel": card.get("parallel"),
                "subset": card.get("subset") or card.get("attributes"),
                "grader": grader,
                "grade": card.get("grade"),
            }
        )
        cert = scan_to_cert(card.get("cert_number"))
        notes = clean_part(card.get("position") or "")
        if not any(card.get(key) for key in ("cert_number", "player", "year", "set", "card_number", "parallel", "subset", "grade", "label_text")):
            review_note = f"OCR review needed: {card.get('error')}" if card.get("error") else "detected slab - review needed"
            notes = clean_part("; ".join(part for part in (notes, review_note) if part))
        return {
            "cert_number": cert,
            "grader": grader or infer_grader(title),
            "card_title": title,
            "purchase_price": None,
            "source": f"Photo: {path.name}",
            "notes": notes,
        }

    def _photo_card_has_inventory(self, card: dict) -> bool:
        if any(card.get(key) for key in ("cert_number", "player", "year", "set", "card_number", "parallel", "subset", "grade", "label_text")):
            return True
        return bool(card.get("is_graded_slab") or card.get("detection_confidence") or card.get("error"))

    def _load_photo_env(self) -> None:
        if not load_dotenv:
            return
        load_dotenv(ROOT / ".env", override=False)
        load_dotenv(PHOTO_APP_DIR / ".env", override=False)
        load_dotenv(PHOTO_APP_ROOT / ".env", override=False)

    def refresh_incoming_index(self) -> None:
        try:
            INCOMING_SHEETS_DIR.mkdir(parents=True, exist_ok=True)
            paths = sorted(INCOMING_SHEETS_DIR.glob("*.xlsx"), key=lambda path: path.name.lower())
        except Exception as error:
            self.incoming_cert_index = {}
            self.review_status.set(f"Incoming sheets unavailable: {error}")
            return
        index: dict[str, dict[str, object]] = {}
        for path in paths:
            try:
                rows = read_simple_spreadsheet(path)
            except Exception:
                continue
            for row in rows:
                cert = scan_to_cert(row.get("cert_number"))
                if not cert or cert in index:
                    continue
                index[cert] = {
                    "sheet": path.name,
                    "path": path,
                    "card_title": row.get("card_title") or "",
                    "grader": row.get("grader") or "",
                    "purchase_price": row.get("purchase_price"),
                    "card_ladder_value": row.get("card_ladder_value"),
                    "card_ladder_comps_average": row.get("card_ladder_comps_average"),
                    "cy_value": row.get("cy_value"),
                    "cy_confidence": row.get("cy_confidence"),
                    "card_ladder_comps": row.get("card_ladder_comps") or "",
                    "best_company": row.get("best_company") or "",
                    "estimated_payout": row.get("estimated_payout"),
                }
        self.incoming_cert_index = index
        self._match_all_review_rows()
        self._refresh_table()
        self.review_status.set(f"Indexed {len(index)} cert(s) from {len(paths)} incoming sheet(s).")

    def refresh_received_sheets(self) -> None:
        try:
            RECEIVED_SHEETS_DIR.mkdir(parents=True, exist_ok=True)
            paths = sorted(RECEIVED_SHEETS_DIR.glob("*.xlsx"), key=lambda path: path.stat().st_mtime, reverse=True)
        except Exception as error:
            self.received_sheet_paths = {}
            if hasattr(self, "received_sheet_combo"):
                self.received_sheet_combo["values"] = []
            self.review_status.set(f"Received sheets unavailable: {error}")
            return
        self.received_sheet_paths = {path.name: path for path in paths}
        if hasattr(self, "received_sheet_combo"):
            names = list(self.received_sheet_paths)
            self.received_sheet_combo["values"] = names
            if names and self.selected_received_sheet.get() not in self.received_sheet_paths:
                self.selected_received_sheet.set(names[0])
            elif not names:
                self.selected_received_sheet.set("")

    def load_selected_received_sheet_for_review(self) -> None:
        name = self.selected_received_sheet.get()
        path = self.received_sheet_paths.get(name)
        if not path:
            messagebox.showinfo("Choose sheet", "Choose a received sheet to load into Assignment.")
            return
        self.review_status.set(f"Loading received sheet: {name}...")
        self.status_var.set(f"Loading received sheet: {name}...")
        threading.Thread(target=self._load_received_sheet_worker, args=(name, path), daemon=True).start()

    def _load_received_sheet_worker(self, name: str, path: Path) -> None:
        try:
            rows = read_simple_spreadsheet(path)
        except Exception as error:
            self.events.put(("load_received_sheet_error", {"name": name, "error": str(error)}))
            return
        self.events.put(("load_received_sheet_done", {"name": name, "rows": rows}))

    def _apply_loaded_received_sheet(self, name: str, rows: list[dict[str, object]]) -> None:
        review_rows = []
        for row in rows:
            review_rows.append(
                {
                    "cert_number": row.get("cert_number"),
                    "grader": row.get("grader"),
                    "card_title": row.get("card_title"),
                    "purchase_price": row.get("purchase_price"),
                    "card_ladder_value": row.get("card_ladder_value"),
                    "card_ladder_comps_average": row.get("card_ladder_comps_average"),
                    "cy_value": row.get("cy_value"),
                    "cy_confidence": row.get("cy_confidence"),
                    "card_ladder_comps": row.get("card_ladder_comps") or "",
                    "best_company": row.get("best_company") or "",
                    "estimated_payout": row.get("estimated_payout"),
                    "source": f"Received Sheet: {name}",
                    "sheet_source": name,
                    "status": "Received",
                    "notes": "Loaded from received sheet",
                }
            )
        added = self._append_review_rows(review_rows, schedule_recommendations=True)
        self.review_status.set(f"Loaded {len(added)} row(s) from {name}.")
        self.status_var.set(f"Loaded received sheet: {name}")

    def add_manual_review_row(self) -> int | None:
        added_rows = self._append_review_rows([
            {
                "cert_number": "",
                "grader": "",
                "card_title": "",
                "purchase_price": None,
                "source": "Manual",
                "notes": "Manual assignment",
            }
        ])
        if added_rows:
            row_id = str(added_rows[-1])
            target_tree = self.receive_tree if hasattr(self, "receive_tree") else self.review_tree
            target_tree.selection_set(row_id)
            target_tree.focus(row_id)
            target_tree.see(row_id)
            self.review_status.set("Manual row added. Double-click cells to edit it.")
            return added_rows[-1]
        return None

    def toggle_review_scanning(self) -> None:
        self.review_scanning_active = not self.review_scanning_active
        self._set_review_station_controls()
        if self.review_scanning_active:
            self.review_status.set("Receive scanning mode armed. Scan received certs now.")
            self._arm_review_scanner()
        else:
            self.review_status.set("Receive station is off.")

    def _set_review_station_controls(self) -> None:
        if not hasattr(self, "review_station_button"):
            return
        self.review_station_button.configure(text="Exit Receive Scanning Mode" if self.review_scanning_active else "Enter Receive Scanning Mode")
        if self.review_scan_entry is not None:
            self.review_scan_entry.configure(state=tk.NORMAL if self.review_scanning_active else tk.DISABLED)

    def add_review_scanned_row(self) -> None:
        if not self.review_scanning_active:
            self.review_status.set("Click Enter Receive Scanning Mode before scanning.")
            return
        cert = scan_to_cert(self.review_scan_cert.get())
        if not cert:
            self.review_status.set("No cert detected. Scan again.")
            self._arm_review_scanner()
            return
        self._append_review_rows([
            {
                "cert_number": cert,
                "grader": "",
                "card_title": "",
                "purchase_price": None,
                "source": "Receive Barcode",
                "notes": "Received",
            }
        ])
        self.review_scan_cert.set("")
        self.review_status.set(f"Received {cert}. Ready for next scan.")
        self._arm_review_scanner()

    def add_review_photos(self) -> None:
        paths = filedialog.askopenfilenames(
            title="Choose receive photos",
            filetypes=[("Images", "*.png *.jpg *.jpeg *.webp *.bmp"), ("All files", "*.*")],
        )
        self._add_review_photo_paths([Path(path) for path in paths])

    def clear_review_photos(self) -> None:
        if self.review_photo_worker and self.review_photo_worker.is_alive():
            messagebox.showinfo("Scan running", "Wait for the receive photo scan to finish before clearing photos.")
            return
        self.review_photo_paths = []
        self.review_photo_status.set("No receive photos selected.")

    def _add_review_photo_paths(self, paths: list[Path]) -> None:
        existing = {path.resolve() for path in self.review_photo_paths if path.exists()}
        added = 0
        for path in paths:
            if not path.exists() or path.resolve() in existing:
                continue
            self.review_photo_paths.append(path)
            existing.add(path.resolve())
            added += 1
        self.review_photo_status.set(f"{len(self.review_photo_paths)} receive photo(s) selected. Added {added}.")

    def scan_review_photos(self) -> None:
        if self.review_photo_worker and self.review_photo_worker.is_alive():
            messagebox.showinfo("Scan running", "Receive photo scan is already running.")
            return
        if not self.review_photo_paths:
            messagebox.showinfo("No photos", "Add receive photos before scanning.")
            return
        if genai is None or identify_cards_sync is None:
            messagebox.showerror("Missing dependency", "Photo OCR dependencies are not available.")
            return
        self._load_photo_env()
        api_key = os.environ.get("GOOGLE_API_KEY", "").strip()
        if not api_key:
            messagebox.showerror("Missing GOOGLE_API_KEY", "Create .env in the L.U.C.A.S project folder or set GOOGLE_API_KEY.")
            return
        self.photo_client = genai.Client(api_key=api_key)
        self.review_photo_status.set(f"Scanning 0/{len(self.review_photo_paths)} receive photo(s)...")
        self.review_photo_worker = threading.Thread(target=self._review_photo_scan_worker, daemon=True)
        self.review_photo_worker.start()

    def _review_photo_scan_worker(self) -> None:
        total = len(self.review_photo_paths)
        detected_total = 0
        for index, path in enumerate(list(self.review_photo_paths), start=1):
            try:
                self.events.put(("review_status", f"Scanning {index}/{total}: {path.name}..."))
                image_b64 = base64.b64encode(path.read_bytes()).decode("utf-8")
                cards = identify_cards_sync(
                    self.photo_client,
                    image_b64,
                    progress_callback=lambda message, i=index, n=total, p=path: self.events.put(
                        ("review_status", f"Scanning {i}/{n}: {p.name} - {message}")
                    ),
                )
                rows = [self._photo_card_to_review_row(path, card) for card in cards if self._photo_card_has_inventory(card)]
                detected_total += len(rows)
                self.events.put(("review_rows", rows))
                self.events.put(("review_status", f"Scanning {index}/{total}: {path.name} -> {len(rows)} receive row(s)."))
            except Exception as error:
                self.events.put(("review_status", f"{path.name}: {error}"))
        self.events.put(("review_status", f"Receive photo scan complete. Added {detected_total} row(s)."))

    def _photo_card_to_review_row(self, path: Path, card: dict) -> dict[str, object]:
        row = self._photo_card_to_row(path, card)
        row["source"] = f"Receive Photo: {path.name}"
        row["notes"] = "Received"
        return row

    def _append_review_rows(self, rows: list[dict[str, object]], schedule_recommendations: bool = False) -> list[int]:
        existing = list(self.review_rows)
        start = len(existing) + 2
        added_excel_rows: list[int] = []
        refreshed_incoming_index = False
        for offset, row in enumerate(rows):
            cert = scan_to_cert(row.get("cert_number"))
            match = self._incoming_match(cert)
            if cert and not match and not refreshed_incoming_index:
                self.refresh_incoming_index()
                refreshed_incoming_index = True
                match = self._incoming_match(cert)
            grader = str(row.get("grader") or match.get("grader") or infer_grader(str(row.get("card_title") or ""))).upper()
            card = str(row.get("card_title") or match.get("card_title") or "").strip()
            purchase_price = row.get("purchase_price") if row.get("purchase_price") is not None else match.get("purchase_price")
            card_ladder_value = row.get("card_ladder_value") if row.get("card_ladder_value") is not None else match.get("card_ladder_value")
            comps_average = row.get("card_ladder_comps_average") if row.get("card_ladder_comps_average") is not None else match.get("card_ladder_comps_average")
            cy_value = row.get("cy_value") if row.get("cy_value") is not None else match.get("cy_value")
            cy_confidence = row.get("cy_confidence") if row.get("cy_confidence") is not None else match.get("cy_confidence")
            comp_details = str(row.get("card_ladder_comps") or match.get("card_ladder_comps") or "")
            best_company = str(row.get("best_company") or match.get("best_company") or "").strip()
            estimated_payout = row.get("estimated_payout") if row.get("estimated_payout") is not None else match.get("estimated_payout")
            sheet_source = str(row.get("sheet_source") or match.get("sheet") or ("NO SHEET FOUND" if cert else ""))
            status = str(row.get("status") or ("Needs setup" if not cert else ("Received" if match else "Received - no incoming match")))
            excel_row = start + offset
            existing.append(
                WorkbookRow(
                    excel_row=excel_row,
                    cert_number=cert,
                    card_title=card,
                    grader=grader,
                    existing_value=purchase_price,
                    card_ladder_value=card_ladder_value,
                    card_ladder_comps_average=comps_average,
                    cy_value=cy_value,
                    cy_confidence=cy_confidence,
                    card_ladder_comps=comp_details,
                    best_company=best_company,
                    estimated_payout=estimated_payout,
                    status=status,
                    notes=str(row.get("notes") or ""),
                )
            )
            self.review_sources[excel_row] = str(row.get("source") or "")
            self.review_sheet_sources[excel_row] = sheet_source
            added_excel_rows.append(excel_row)
        self.review_rows = existing
        self._refresh_table(schedule_recommendations=schedule_recommendations)
        return added_excel_rows

    def _incoming_match(self, cert: str) -> dict[str, object]:
        return self.incoming_cert_index.get(scan_to_cert(cert), {})

    def _match_all_review_rows(self) -> None:
        for row in self.review_rows:
            match = self._incoming_match(row.cert_number)
            self.review_sheet_sources[row.excel_row] = str(match.get("sheet") or "NO SHEET FOUND")
            if match:
                if is_placeholder_title(row.card_title, row.grader) and match.get("card_title"):
                    row.card_title = str(match.get("card_title") or "")
                if not row.grader and match.get("grader"):
                    row.grader = str(match.get("grader") or "")
                if row.existing_value is None and match.get("purchase_price") is not None:
                    row.existing_value = match.get("purchase_price")
                if row.card_ladder_value is None and match.get("card_ladder_value") is not None:
                    row.card_ladder_value = match.get("card_ladder_value")
                if row.card_ladder_comps_average is None and match.get("card_ladder_comps_average") is not None:
                    row.card_ladder_comps_average = match.get("card_ladder_comps_average")
                if row.cy_value is None and match.get("cy_value") is not None:
                    row.cy_value = match.get("cy_value")
                if row.cy_confidence is None and match.get("cy_confidence") is not None:
                    row.cy_confidence = match.get("cy_confidence")
                if not row.card_ladder_comps and match.get("card_ladder_comps"):
                    row.card_ladder_comps = str(match.get("card_ladder_comps") or "")
                if not row.best_company and match.get("best_company"):
                    row.best_company = str(match.get("best_company") or "")
                if row.estimated_payout is None and match.get("estimated_payout") is not None:
                    row.estimated_payout = match.get("estimated_payout")
                row.status = "Received"
            elif row.status == "Received":
                row.status = "Received - no incoming match"

    def clear_review_rows(self) -> None:
        self.review_rows = []
        self.review_sources = {}
        self.review_sheet_sources = {}
        self._refresh_table()
        self.review_status.set("Receive/assignment rows cleared.")

    def delete_selected_review_rows(self) -> None:
        tree = self.review_tree
        if hasattr(self, "receive_tree") and self.receive_tree.selection():
            tree = self.receive_tree
        elif self.review_tree.selection():
            tree = self.review_tree
        deleted = self._delete_selected_rows(
            tree,
            self.review_rows,
            self.review_sources,
            self.review_sheet_sources,
        )
        if deleted:
            self.review_status.set(f"Deleted {deleted} receive/assignment row(s).")
            self.status_var.set(f"Deleted {deleted} receive/assignment row(s).")
        else:
            self.review_status.set("Select receive or assignment rows to delete.")

    def mark_review_received_in_sheets(self) -> None:
        certs = {scan_to_cert(row.cert_number) for row in self.review_rows if scan_to_cert(row.cert_number)}
        if not certs:
            messagebox.showinfo("No received certs", "Scan or load received cards in Receive before marking sheets.")
            return
        paths: list[Path] = []
        errors: list[str] = []
        for directory in (INCOMING_SHEETS_DIR, WORKING_SHEETS_DIR):
            try:
                directory.mkdir(parents=True, exist_ok=True)
                paths.extend(sorted(directory.glob("*.xlsx"), key=lambda path: path.name.lower()))
            except Exception as error:
                errors.append(f"{directory}: {error}")
        if not paths:
            messagebox.showinfo("No sheets found", "No incoming or working sheets were found to update.")
            return
        try:
            with shared_lock(CARD_PIPELINE_DIR, "receive-company-sheets", self.lucas_identity):
                result = mark_received_in_workbooks(paths, certs)
                errors.extend(result.get("errors") or [])
                rows_marked = int(result.get("rows_marked") or 0)
                files_updated = int(result.get("files_updated") or 0)
                certs_marked = len(result.get("certs_marked") or set())
                company_rows_added = 0
                company_rows_missing_company = 0
                inventory_rows_added = 0
                if rows_marked:
                    marked_certs = result.get("certs_marked", set())
                    company_rows = [
                        row
                        for row in self.review_rows
                        if row.company_pile and scan_to_cert(row.cert_number) in marked_certs
                    ]
                    inventory_rows = [
                        row
                        for row in self.review_rows
                        if not row.company_pile and scan_to_cert(row.cert_number) in marked_certs
                    ]
                    self._apply_recommendations_to_rows(company_rows, force=True)
                    eligible_company_rows = [row for row in company_rows if self._row_has_assignable_company(row)]
                    company_rows_missing_company = len(company_rows) - len(eligible_company_rows)
                    if eligible_company_rows:
                        company_result = append_company_sheet_rows(
                            COMPANY_SHEETS_DIR,
                            eligible_company_rows,
                            self.review_sources,
                            self.review_sheet_sources,
                        )
                        company_rows_added = int(company_result.get("rows_added") or 0)
                        self.record_profit_sales(list(company_result.get("added_records") or []))
                        errors.extend(company_result.get("errors") or [])
                    if inventory_rows:
                        inventory_records = [
                            self._inventory_record_from_row(
                                row,
                                person=self._assignment_person_for_row(row),
                                source_sheet=self.review_sheet_sources.get(row.excel_row, ""),
                                source=self.review_sources.get(row.excel_row, ""),
                                notes="Received without company pile",
                            )
                            for row in inventory_rows
                        ]
                        inventory_rows_added = self.add_inventory_records(inventory_records)
                moved_received = self._move_fully_received_sheets_to_received(paths)
                if moved_received:
                    self._save_sheet_markers()
        except Exception as error:
            messagebox.showerror("Shared folder busy", str(error))
            self.status_var.set(f"Receive update failed: {error}")
            return
        self.refresh_incoming_index()
        self.refresh_working_sheets()
        self.refresh_received_sheets()
        if rows_marked:
            self.review_status.set(f"Marked {rows_marked} row(s) received across {files_updated} sheet file(s).")
            self.status_var.set(f"Marked {certs_marked}/{len(certs)} received cert(s) in sheets.")
        else:
            self.review_status.set("No matching cert rows were found in incoming or working sheets.")
            self.status_var.set("No sheet rows marked received.")
        if moved_received:
            self.status_var.set(f"Moved {len(moved_received)} fully received sheet(s) to RECEIVED SHEETS.")
        if company_rows_added:
            self.status_var.set(f"Added {company_rows_added} card(s) to weekly company sheet(s).")
        if inventory_rows_added:
            self.status_var.set(f"Added {inventory_rows_added} received card(s) to active inventory.")
        elif company_rows_missing_company:
            self.status_var.set(f"{company_rows_missing_company} checked company pile card(s) had no Best Company.")
        self.refresh_home()
        summary_lines = [
            f"Marked rows: {rows_marked}",
            f"Updated sheet files: {files_updated}",
            f"Matched certs: {certs_marked}/{len(certs)}",
        ]
        if moved_received:
            summary_lines.append(f"Moved to received: {len(moved_received)}")
        if company_rows_added:
            summary_lines.append(f"Company sheet rows added: {company_rows_added}")
        if inventory_rows_added:
            summary_lines.append(f"Inventory rows added: {inventory_rows_added}")
        if company_rows_missing_company:
            summary_lines.append(f"Company pile rows missing Best Company: {company_rows_missing_company}")
        if errors:
            messagebox.showwarning("Mark received completed with warnings", "\n".join(summary_lines + ["", "Warnings:", *errors[:8]]))
        else:
            messagebox.showinfo("Mark received complete", "\n".join(summary_lines))

    def _apply_recommendations_to_rows(self, rows: list[WorkbookRow], force: bool = False) -> None:
        for row in rows:
            if not force and row.best_company and row.estimated_payout is not None:
                continue
            recommendation = self.assignment_engine.recommend(row, person=getattr(self, "_assignment_person_for_row", lambda _row: "")(row))
            if recommendation.payout is None:
                self._record_unassigned_player(row)
                if force:
                    row.best_company = NO_COMPANY_TAKES_LABEL
                    row.estimated_payout = None
                elif not row.best_company:
                    row.best_company = NO_COMPANY_TAKES_LABEL
                    row.estimated_payout = None
                continue
            row.best_company = recommendation.company
            row.estimated_payout = recommendation.payout

    def _row_has_assignable_company(self, row: WorkbookRow) -> bool:
        company = str(row.best_company or "").strip()
        return bool(company) and company.upper() != NO_COMPANY_TAKES_LABEL

    def _ensure_company_sheet_folders(self) -> None:
        try:
            COMPANY_SHEETS_DIR.mkdir(parents=True, exist_ok=True)
            for company in self.assignment_engine.companies:
                folder_name = self._safe_company_folder_name(company.name)
                if folder_name:
                    (COMPANY_SHEETS_DIR / folder_name).mkdir(parents=True, exist_ok=True)
        except Exception:
            pass

    def _weekly_company_sheet_timer(self) -> None:
        self._ensure_weekly_company_sheets_due()
        self.after(5 * 60 * 1000, self._weekly_company_sheet_timer)

    def _ensure_weekly_company_sheets_due(self, now: datetime | None = None) -> dict[str, object]:
        now = now or datetime.now()
        week_start_date = company_sheet_week_start_for_time(now)
        marker_key = week_start_date.isoformat()
        try:
            markers = json.loads(WEEKLY_COMPANY_SHEETS_PATH.read_text(encoding="utf-8")) if WEEKLY_COMPANY_SHEETS_PATH.exists() else {}
        except Exception:
            markers = {}
        weeks = markers.get("weeks") if isinstance(markers, dict) else {}
        if not isinstance(weeks, dict):
            weeks = {}
        companies = [company.name for company in self.assignment_engine.companies if company.name]
        company_names = sorted(companies, key=str.lower)
        marker = weeks.get(marker_key) if isinstance(weeks.get(marker_key), dict) else {}
        if marker.get("companies") == company_names:
            return {"created": [], "existing": [], "errors": [], "skipped": True, "week_start": marker_key}
        result = ensure_company_weekly_sheets(COMPANY_SHEETS_DIR, companies, week_start_date)
        errors = list(result.get("errors") or [])
        if not errors:
            weeks[marker_key] = {
                "created_at": now.isoformat(timespec="seconds"),
                "week_start": marker_key,
                "company_count": len(companies),
                "companies": company_names,
                "created_count": len(result.get("created") or []),
                "existing_count": len(result.get("existing") or []),
            }
            atomic_write_json(WEEKLY_COMPANY_SHEETS_PATH, {"weeks": weeks})
            created_count = len(result.get("created") or [])
            if created_count:
                self.status_var.set(f"Created {created_count} weekly company sheet tab(s) for week of {marker_key}.")
        return {**result, "skipped": False, "week_start": marker_key}

    def _safe_company_folder_name(self, name: str) -> str:
        return re.sub(r"[<>:\"/\\|?*]+", " ", str(name or "")).strip()[:140].strip()

    def _arm_review_scanner(self) -> None:
        if self.review_mode.get() != "Automatic Receive" or self.review_scan_entry is None:
            return
        try:
            self.review_scan_entry.focus_set()
            self.review_scan_entry.icursor(tk.END)
        except tk.TclError:
            pass

    def _cardladder_extension_warning(self) -> str:
        with self.state.lock:
            last_seen = self.state.last_seen_extension
            extension_version = self.state.extension_version
            manifest_version = self.state.extension_manifest_version
            extension_name = self.state.extension_name
            extension_url = self.state.extension_url
        if not last_seen:
            return ""
        loaded_label = extension_name or "Card Ladder extension"
        version_label = extension_version or "unversioned/old"
        manifest_label = manifest_version or "unknown"
        if extension_version != EXPECTED_CARDLADDER_EXTENSION_VERSION:
            return (
                f"{loaded_label} checked in at {last_seen}, but it is {version_label} "
                f"(manifest {manifest_label}).\n\n"
                f"Expected helper version: {EXPECTED_CARDLADDER_EXTENSION_VERSION} "
                f"(manifest {EXPECTED_CARDLADDER_MANIFEST_VERSION}).\n\n"
                "Open chrome://extensions, remove or disable the old Card Ladder Auto-Comp helper, "
                f"then Load unpacked from:\n{CARDLADDER_EXTENSION_DIR}\n\n"
                f"Chrome extension URL seen by the app: {extension_url or 'unknown'}"
            )
        return ""

    def run_all_comps(self) -> None:
        app_debug_log(f"run_all_comps_clicked rows={len(self.state.rows)}")
        if not self.state.rows:
            app_debug_log("run_all_comps_blocked reason=no_rows")
            messagebox.showinfo("No comp sheet loaded", "Choose and load a working sheet in the Comp tab first.")
            return
        requery_all = self.comp_scope_label.get() == COMP_SCOPE_ALL
        source_label = self.comp_source_label.get()
        run_card_ladder = source_label in {COMP_SOURCE_BOTH, COMP_SOURCE_CARD_LADDER}
        run_cy = source_label in {COMP_SOURCE_BOTH, COMP_SOURCE_CY}
        card_ladder_eligible = [
            row
            for row in self.state.rows
            if run_card_ladder and row.cert_number and row.grader and (requery_all or not row_has_comp_data(row))
        ]
        cy_eligible = [
            row
            for row in self.state.rows
            if run_cy and row.cert_number and row.grader and (requery_all or row.cy_value is None)
        ]
        if card_ladder_eligible:
            extension_warning = self._cardladder_extension_warning()
            if extension_warning:
                messagebox.showwarning("Reload Card Ladder extension", extension_warning)
                self.status_var.set("Reload the Card Ladder Chrome extension before comping.")
                return
        if not card_ladder_eligible and not cy_eligible:
            if requery_all:
                message = f"No rows have both a cert number and company ready for {source_label}."
            else:
                message = f"No rows are missing {source_label} data. Switch Run Scope to Recomp All if you want to refresh every row."
            app_debug_log(f"run_all_comps_blocked reason=no_eligible source={source_label} requery_all={requery_all}")
            messagebox.showinfo("No eligible rows", message)
            self.status_var.set(message)
            return
        self.state.set_comp_strategy(COMP_STRATEGY_DISPLAY.get(self.comp_strategy_label.get(), COMP_STRATEGY_AVERAGE))
        command_id = 0
        card_ladder_command_id = 0
        if card_ladder_eligible:
            card_ladder_command_id = self.state.start_all_comps(requery_all=requery_all)
            command_id = card_ladder_command_id
        if cy_eligible:
            command_id = self.state.start_cy_lookups(cy_eligible, defer=bool(card_ladder_eligible))
        self.comp_output_saved = False
        self._refresh_table()
        if card_ladder_eligible:
            self.after(12000, lambda queued_command_id=card_ladder_command_id: self._warn_if_extension_not_checked_in(queued_command_id))
        pieces = []
        if card_ladder_eligible:
            pieces.append(f"{len(card_ladder_eligible)} Card Ladder")
        if cy_eligible:
            pieces.append(f"{len(cy_eligible)} CY")
        self.status_var.set(f"Queued {' and '.join(pieces)} row(s) using {self.comp_scope_label.get()} as command #{command_id}.")
        app_debug_log(
            f"run_all_comps_queued command={command_id} source={source_label} "
            f"card_ladder={len(card_ladder_eligible)} cy={len(cy_eligible)}"
        )

    def _warn_if_extension_not_checked_in(self, command_id: int) -> None:
        extension_warning = self._cardladder_extension_warning()
        if extension_warning:
            messagebox.showwarning("Reload Card Ladder extension", extension_warning)
            self.status_var.set("Reload the Card Ladder Chrome extension before comping.")
            return
        with self.state.lock:
            command_pending = bool(self.state.command and self.state.command.get("id") == command_id)
        if not command_pending:
            return
        messagebox.showwarning(
            "Card Ladder extension not connected",
            "The rows were queued, but the Card Ladder Chrome extension has not checked in. Make sure Chrome is open, logged into Card Ladder, and the bundled extension is loaded.",
        )

    def stop_comp_run(self) -> None:
        self.state.request_cancel()
        self.comp_output_saved = False
        self._refresh_table()
        self.status_var.set("Stop requested. Card Ladder will stop after the current row.")

    def clear_comp_rows(self) -> None:
        if self.state.rows and not self.comp_output_saved:
            confirmed = messagebox.askyesno(
                "Clear unsaved comp rows?",
                "These comp rows have not been saved as an output. Clear them anyway?",
                icon=messagebox.WARNING,
            )
            if not confirmed:
                self.status_var.set("Clear comp rows cancelled.")
                return
        self.state.set_rows([])
        self.row_sources = {}
        self.comp_sheet_sources = {}
        self.selected_working_sheet.set("")
        self.comp_output_saved = True
        self._cancel_cell_edit()
        try:
            self.working_sheet_list.selection_clear(0, tk.END)
        except tk.TclError:
            pass
        self._refresh_table()
        self.status_var.set("Comp rows cleared.")

    def recalculate_comp_method(self, _event=None) -> None:
        strategy = COMP_STRATEGY_DISPLAY.get(self.comp_strategy_label.get(), COMP_STRATEGY_AVERAGE)
        self.state.set_comp_strategy(strategy)
        updated = 0
        with self.state.lock:
            for row in self.state.rows:
                comps = parse_formatted_comps(row.card_ladder_comps)
                if not comps:
                    continue
                row.card_ladder_comps_average = comp_price(comps, strategy)
                row.card_ladder_comps = format_comps(comps, strategy)
                updated += 1
        if updated:
            self.comp_output_saved = False
        self._refresh_table(schedule_recommendations=bool(updated))
        if updated:
            self.status_var.set(f"Recalculated {updated} comp row(s) with {self.comp_strategy_label.get()}.")
        elif self.state.rows:
            self.status_var.set("Comp method updated. No stored comp details were available to recalculate.")
        else:
            self.status_var.set("Comp method updated.")

    def save_output(self) -> None:
        if not self.state.rows:
            messagebox.showinfo("No rows", "Load or scan cards before saving.")
            return
        self._apply_recommendations()
        default = default_output_path(ROOT)
        path = filedialog.asksaveasfilename(
            title="Save pipeline workbook",
            initialdir=str(default.parent),
            initialfile=default.name,
            defaultextension=".xlsx",
            filetypes=[("Excel workbook", "*.xlsx")],
        )
        if not path:
            return
        try:
            with shared_lock(CARD_PIPELINE_DIR, "workbook-writes", self.lucas_identity):
                write_pipeline_output(Path(path), self.state.rows, self.row_sources)
        except Exception as error:
            messagebox.showerror("Save failed", str(error))
            return
        self.comp_output_saved = True
        self.status_var.set(f"Saved {path}")

    def save_working_sheet(self) -> None:
        if not self.intake_rows:
            messagebox.showinfo("No create rows", "Scan or load cards in Create before saving a working sheet.")
            return
        title = self.working_sheet_title.get().strip()
        if not title:
            messagebox.showinfo("Title required", "Enter a working sheet title first.")
            return
        self.apply_create_seller_terms(show_status=False)
        seller = self.seller_terms_seller_var.get().strip() if self._network_mode_enabled() and hasattr(self, "seller_terms_seller_var") else ""
        try:
            with shared_lock(CARD_PIPELINE_DIR, "workbook-writes", self.lucas_identity):
                path = working_sheet_path(WORKING_SHEETS_DIR, title)
                write_working_sheet(path, self.intake_rows, self.intake_sources)
            if seller:
                self._assign_sheet_to_seller("Working", path.name, seller)
        except Exception as error:
            messagebox.showerror("Save failed", str(error))
            return
        seller_note = f" Assigned to {seller} for payouts." if seller else ""
        self.status_var.set(f"Saved working sheet: {path}.{seller_note}")
        self.intake_rows = []
        self.intake_sources = {}
        self.intake_sheet_sources = {}
        self.working_sheet_title.set("")
        self._refresh_table()
        self.refresh_home()

    def refresh_pipeline(self) -> None:
        self.refresh_working_sheets()
        self.refresh_home()
        self._refresh_table()

    def refresh_working_sheets(self) -> None:
        try:
            CARD_PIPELINE_DIR.mkdir(parents=True, exist_ok=True)
            WORKING_SHEETS_DIR.mkdir(parents=True, exist_ok=True)
            paths = sorted(WORKING_SHEETS_DIR.glob("*.xlsx"), key=lambda path: path.stat().st_mtime, reverse=True)
        except Exception as error:
            self.working_sheet_paths = {}
            if hasattr(self, "working_sheet_list"):
                self.working_sheet_list.delete(0, tk.END)
            self.status_var.set(f"Working sheets unavailable: {error}")
            return
        self.working_sheet_paths = {path.name: path for path in paths}
        if hasattr(self, "working_sheet_list"):
            self.working_sheet_list.delete(0, tk.END)
            for name in self.working_sheet_paths:
                self.working_sheet_list.insert(tk.END, name)
        if paths and self.selected_working_sheet.get() not in self.working_sheet_paths:
            self.selected_working_sheet.set(paths[0].name)
        self._select_working_sheet_in_list()
        self.status_var.set(f"Found {len(paths)} working sheet(s).")

    def load_selected_working_sheet(self) -> None:
        name = self._selected_working_sheet_name()
        path = self.working_sheet_paths.get(name)
        if not path:
            messagebox.showinfo("Choose sheet", "Choose a working sheet first.")
            return
        self.status_var.set(f"Loading working sheet: {name}...")
        threading.Thread(target=self._load_working_sheet_worker, args=(name, path), daemon=True).start()

    def _load_working_sheet_worker(self, name: str, path: Path) -> None:
        try:
            rows = read_simple_spreadsheet(path)
        except Exception as error:
            self.events.put(("load_working_sheet_error", {"name": name, "error": str(error)}))
            return
        self.events.put(("load_working_sheet_done", {"name": name, "rows": rows}))

    def _apply_loaded_working_sheet(self, name: str, rows: list[dict[str, object]]) -> None:
        workbook_rows: list[WorkbookRow] = []
        sources: dict[int, str] = {}
        for offset, row in enumerate(rows, start=2):
            cert = str(row.get("cert_number") or "")
            grader = str(row.get("grader") or infer_grader(str(row.get("card_title") or "")) or "PSA").upper()
            card = str(row.get("card_title") or "").strip()
            workbook_rows.append(
                WorkbookRow(
                    excel_row=offset,
                    cert_number=cert,
                    card_title=card,
                    grader=grader,
                    existing_value=row.get("purchase_price"),
                    card_ladder_value=row.get("card_ladder_value"),
                    card_ladder_comps_average=row.get("card_ladder_comps_average"),
                    cy_value=row.get("cy_value"),
                    cy_confidence=row.get("cy_confidence"),
                    card_ladder_comps=str(row.get("card_ladder_comps") or ""),
                    best_company=str(row.get("best_company") or ""),
                    estimated_payout=row.get("estimated_payout"),
                    status=str(row.get("status") or ("Ready" if cert and grader else "Needs setup")),
                    notes=str(row.get("notes") or ""),
                )
            )
            sources[offset] = str(row.get("source") or name)
        self.state.set_rows(workbook_rows)
        self.row_sources = sources
        self.comp_sheet_sources = {}
        self.comp_output_saved = True
        self._refresh_table(schedule_recommendations=any(row.card_ladder_comps_average is not None for row in workbook_rows))
        self.selected_working_sheet.set(name)
        self._select_working_sheet_in_list()
        self.status_var.set(f"Loaded working sheet: {name}")

    def _selected_working_sheet_name(self) -> str:
        if hasattr(self, "working_sheet_list"):
            selected = self.working_sheet_list.curselection()
            if selected:
                return str(self.working_sheet_list.get(selected[0]))
        return self.selected_working_sheet.get()

    def _select_working_sheet_in_list(self) -> None:
        if not hasattr(self, "working_sheet_list"):
            return
        target = self.selected_working_sheet.get()
        self.working_sheet_list.selection_clear(0, tk.END)
        for index, name in enumerate(self.working_sheet_list.get(0, tk.END)):
            if name == target:
                self.working_sheet_list.selection_set(index)
                self.working_sheet_list.see(index)
                break

    def clear_rows(self) -> None:
        self.intake_rows = []
        self.intake_sources = {}
        self.intake_sheet_sources = {}
        self._refresh_table()
        self.status_var.set("Create rows cleared.")

    def delete_selected_intake_rows(self) -> None:
        deleted = self._delete_selected_rows(
            self.intake_tree,
            self.intake_rows,
            self.intake_sources,
            self.intake_sheet_sources,
        )
        if deleted:
            self.status_var.set(f"Deleted {deleted} create row(s).")
        else:
            self.status_var.set("Select create rows to delete.")

    def _append_rows(self, rows: list[dict[str, object]]) -> list[int]:
        existing = list(self.intake_rows)
        start = len(existing) + 2
        added_excel_rows: list[int] = []
        for offset, row in enumerate(rows):
            cert = str(row.get("cert_number") or "")
            grader = str(row.get("grader") or infer_grader(str(row.get("card_title") or ""))).upper()
            card = str(row.get("card_title") or "").strip()
            status = "Ready" if cert and grader else "Needs setup"
            notes = str(row.get("notes") or "")
            excel_row = start + offset
            existing.append(
                WorkbookRow(
                    excel_row=excel_row,
                    cert_number=cert,
                    card_title=card,
                    grader=grader,
                    existing_value=row.get("purchase_price"),
                    status=status,
                    notes=notes,
                )
            )
            self.intake_sources[excel_row] = str(row.get("source") or "")
            self.intake_sheet_sources[excel_row] = ""
            added_excel_rows.append(excel_row)
        self.intake_rows = existing
        self.apply_create_seller_terms(show_status=False)
        self._refresh_table()
        return added_excel_rows

    def _apply_recommendations(self) -> None:
        for row in [*self.state.rows, *self.review_rows]:
            recommendation = self.assignment_engine.recommend(row, person=getattr(self, "_assignment_person_for_row", lambda _row: "")(row))
            if recommendation.payout is None:
                self._record_unassigned_player(row)
                row.best_company = NO_COMPANY_TAKES_LABEL
                row.estimated_payout = None
                continue
            row.best_company = recommendation.company
            row.estimated_payout = recommendation.payout

    def _load_player_overrides(self) -> int:
        return assignment_engine.load_player_sport_overrides(PLAYER_OVERRIDES_PATH)

    def _load_unassigned_players(self) -> dict[str, dict[str, object]]:
        try:
            raw = json.loads(UNASSIGNED_PLAYERS_PATH.read_text(encoding="utf-8"))
        except Exception:
            return {}
        entries = raw.get("entries", raw) if isinstance(raw, dict) else {}
        return entries if isinstance(entries, dict) else {}

    def _save_unassigned_players(self, entries: dict[str, dict[str, object]]) -> None:
        UNASSIGNED_PLAYERS_PATH.parent.mkdir(parents=True, exist_ok=True)
        atomic_write_json(UNASSIGNED_PLAYERS_PATH, {"entries": entries})

    def _record_unassigned_players(self, rows: list[WorkbookRow]) -> None:
        for row in rows:
            self._record_unassigned_player(row)

    def _record_unassigned_player(self, row: WorkbookRow) -> None:
        best_company = str(row.best_company or "").strip()
        if row.estimated_payout is not None or (best_company and best_company.upper() != NO_COMPANY_TAKES_LABEL):
            return
        if row.card_ladder_comps_average is None and row.card_ladder_value is None:
            return
        title = str(row.card_title or "").strip()
        if not title:
            return
        player_guess = self._guess_unassigned_player(row)
        if not player_guess:
            return
        key = re.sub(r"[^a-z0-9]+", " ", player_guess.lower()).strip() or re.sub(r"[^a-z0-9]+", " ", title.lower()).strip()
        if not key:
            return
        try:
            with shared_lock(CARD_PIPELINE_DIR, "unassigned-players", self.lucas_identity):
                entries = self._load_unassigned_players()
                existing = dict(entries.get(key, {}))
                existing["player"] = existing.get("player") or player_guess
                existing["last_title"] = title
                existing["sample_titles"] = self._append_unique_sample(existing.get("sample_titles"), title)
                existing["count"] = int(existing.get("count") or 0) + 1
                existing["last_seen"] = datetime.now().isoformat(timespec="seconds")
                existing["source"] = str(row.cert_number or "")
                entries[key] = existing
                self._save_unassigned_players(entries)
        except Exception:
            pass

    def _append_unique_sample(self, samples: object, value: str) -> list[str]:
        result = [str(item) for item in samples] if isinstance(samples, list) else []
        if value not in result:
            result.append(value)
        return result[-5:]

    def _guess_unassigned_player(self, row: WorkbookRow) -> str:
        title = str(row.card_title or "").strip()
        if not title:
            return ""
        grader_pattern = r"\b(?:PSA|BGS|SGC|CGC)\b"
        before_grader = re.split(grader_pattern, title, flags=re.I)[0].strip()
        before_grader = self._strip_card_variant_tail(before_grader)
        number_match = re.search(r"\b\d+[A-Za-z]?\s+([A-Z][A-Za-z'.-]+(?:\s+[A-Z][A-Za-z'.-]+){1,3})\s*$", before_grader)
        if number_match:
            return number_match.group(1).strip()
        words = re.findall(r"[A-Z][A-Za-z'.-]+", before_grader)
        stop_words = {
            "Topps", "Panini", "Bowman", "Donruss", "Optic", "Prizm", "Select", "Mosaic", "Contenders",
            "Chrome", "Finest", "Upper", "Deck", "Fleer", "Score", "Absolute", "Elite", "Stadium",
            "Club", "Uniformity", "Dominance", "Rookie", "Prospect", "Autographs", "Variation",
        }
        candidates = [word for word in words if word not in stop_words and not re.fullmatch(r"[IVX]+", word)]
        return " ".join(candidates[-2:]).strip() if len(candidates) >= 2 else ""

    def _strip_card_variant_tail(self, value: str) -> str:
        text = str(value or "").strip()
        tail_terms = (
            "chrome-refractor",
            "chrome refractor",
            "refractor",
            "silver",
            "prizm",
            "mosaic",
            "green",
            "blue",
            "red",
            "gold",
            "orange",
            "purple",
            "black",
            "white",
            "pink",
            "aqua",
            "teal",
            "auto",
            "autograph",
            "rookie",
        )
        changed = True
        while changed:
            changed = False
            for term in tail_terms:
                pattern = rf"(?:[- ]+{re.escape(term)})$"
                stripped = re.sub(pattern, "", text, flags=re.I).strip()
                if stripped != text:
                    text = stripped
                    changed = True
        return text

    def _auto_categorize_unassigned_players(self) -> dict[str, object]:
        entries = self._load_unassigned_players()
        resolved: list[tuple[str, str, str, str]] = []
        unresolved: list[str] = []
        errors: list[str] = []
        for key, entry in list(entries.items()):
            player = str(entry.get("player") or "").strip()
            if not player:
                unresolved.append(str(key))
                continue
            try:
                category, evidence = self._search_unassigned_player_category(entry)
            except Exception as error:
                errors.append(f"{player}: {error}")
                continue
            if category:
                resolved.append((str(key), player, category, evidence))
            else:
                unresolved.append(str(key))
        if resolved:
            for _key, player, category, _evidence in resolved:
                self._write_player_category_override(player, category)
            with shared_lock(CARD_PIPELINE_DIR, "unassigned-players", self.lucas_identity):
                latest = self._load_unassigned_players()
                for key, _player, _category, _evidence in resolved:
                    latest.pop(key, None)
                self._save_unassigned_players(latest)
            self.assignment_engine = AssignmentEngine.load()
        return {
            "resolved": len(resolved),
            "unresolved": len(unresolved),
            "errors": errors,
            "details": resolved,
        }

    def _search_unassigned_player_category(self, entry: dict[str, object]) -> tuple[str, str]:
        player = str(entry.get("player") or "").strip()
        title = str(entry.get("last_title") or "").strip()
        query = " ".join(part for part in (player, title, "sports cards category") if part)
        text = " ".join(part for part in (player, title, self._category_research_text(player, query)) if part)
        return self._infer_category_from_web_text(text)

    def _category_research_text(self, player: str, query: str) -> str:
        parts: list[str] = []
        for fetcher in (
            lambda: self._wikipedia_search_text(player),
            lambda: self._wikidata_search_text(player),
            lambda: self._duckduckgo_search_text(query),
        ):
            try:
                text = fetcher()
            except Exception:
                continue
            if text:
                parts.append(text)
        return " ".join(parts)

    def _wikipedia_search_text(self, player: str) -> str:
        search_url = "https://en.wikipedia.org/w/api.php?" + urllib.parse.urlencode({
            "action": "query",
            "list": "search",
            "srsearch": player,
            "format": "json",
            "srlimit": "3",
        })
        payload = self._read_json_url(search_url)
        snippets: list[str] = []
        for item in ((payload.get("query") or {}).get("search") or [])[:3]:
            title = str(item.get("title") or "")
            snippet = re.sub(r"<[^>]+>", " ", str(item.get("snippet") or ""))
            snippets.append(f"{title} {snippet}")
            if title:
                summary_url = "https://en.wikipedia.org/api/rest_v1/page/summary/" + urllib.parse.quote(title.replace(" ", "_"), safe="")
                try:
                    summary = self._read_json_url(summary_url)
                    snippets.append(" ".join(str(summary.get(key) or "") for key in ("title", "description", "extract")))
                except Exception:
                    pass
        return " ".join(snippets)

    def _wikidata_search_text(self, player: str) -> str:
        search_url = "https://www.wikidata.org/w/api.php?" + urllib.parse.urlencode({
            "action": "wbsearchentities",
            "search": player,
            "language": "en",
            "format": "json",
            "limit": "5",
        })
        payload = self._read_json_url(search_url)
        return " ".join(
            " ".join(str(item.get(key) or "") for key in ("label", "description"))
            for item in payload.get("search") or []
        )

    def _duckduckgo_search_text(self, query: str) -> str:
        return self._html_search_text(f"https://duckduckgo.com/html/?q={urllib.parse.quote_plus(query)}")

    def _web_search_text(self, query: str) -> str:
        return self._duckduckgo_search_text(query)

    def _read_json_url(self, url: str) -> dict[str, object]:
        request = urllib.request.Request(
            url,
            headers={
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) LUCAS/1.0",
                "Accept": "application/json,text/plain,*/*",
            },
        )
        with urllib.request.urlopen(request, timeout=12) as response:
            return json.loads(response.read(200000).decode("utf-8", errors="replace"))

    def _html_search_text(self, url: str) -> str:
        request = urllib.request.Request(
            url,
            headers={
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) LUCAS/1.0",
                "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            },
        )
        with urllib.request.urlopen(request, timeout=12) as response:
            raw = response.read(200000).decode("utf-8", errors="replace")
        text = re.sub(r"<script\b.*?</script>|<style\b.*?</style>", " ", raw, flags=re.I | re.S)
        text = re.sub(r"<[^>]+>", " ", text)
        return re.sub(r"\s+", " ", html.unescape(text)).strip()

    def _infer_category_from_web_text(self, text: str) -> tuple[str, str]:
        haystack = re.sub(r"\s+", " ", str(text or "").lower())
        scores: dict[str, int] = {}
        for category, terms in ASSIGNMENT_CATEGORY_WEB_SIGNALS.items():
            score = 0
            for term in terms:
                cleaned = str(term).lower()
                matches = len(re.findall(rf"\b{re.escape(cleaned)}\b", haystack))
                if not matches:
                    continue
                score += matches * (4 if cleaned == category else 2 if " " in cleaned else 1)
            if score:
                scores[category] = score
        if not scores:
            return "", ""
        ranked = sorted(scores.items(), key=lambda item: item[1], reverse=True)
        best_category, best_score = ranked[0]
        second_score = ranked[1][1] if len(ranked) > 1 else 0
        if best_score < 2 or (second_score and best_score < second_score + 2):
            return "", ""
        return best_category, f"score {best_score}"

    def open_unassigned_players_dialog(self) -> None:
        entries = self._load_unassigned_players()
        popup = tk.Toplevel(self)
        popup.title("Unassigned Players")
        popup.configure(bg="#1f1f1f")
        popup.transient(self)
        popup.grab_set()
        popup.geometry("980x560")
        popup.minsize(860, 480)

        frame = ttk.Frame(popup, style="Panel.TFrame", padding=(16, 14))
        frame.pack(fill=tk.BOTH, expand=True)
        ttk.Label(frame, text="Unassigned Players", style="Panel.TLabel", font=("Segoe UI Semibold", 13)).grid(row=0, column=0, columnspan=4, sticky="w")
        ttk.Label(frame, text="Search the player, choose a category, then save to teach assignment matching.", style="Muted.TLabel").grid(row=1, column=0, columnspan=4, sticky="w", pady=(2, 12))

        tree = ttk.Treeview(frame, columns=("player", "count", "last_seen", "title"), show="headings", selectmode="browse")
        headings = {"player": "Player", "count": "Count", "last_seen": "Last Seen", "title": "Sample Card"}
        widths = {"player": 180, "count": 60, "last_seen": 150, "title": 510}
        for column in headings:
            tree.heading(column, text=headings[column], anchor=tk.W)
            tree.column(column, width=widths[column], minwidth=50, stretch=column == "title")
        tree.grid(row=2, column=0, columnspan=4, sticky="nsew")
        scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=tree.yview)
        scrollbar.grid(row=2, column=4, sticky="ns")
        tree.configure(yscrollcommand=scrollbar.set)

        player_var = tk.StringVar()
        category_var = tk.StringVar(value=ASSIGNMENT_CATEGORY_OPTIONS[0])
        selected_key = tk.StringVar()
        ttk.Label(frame, text="Player", style="Panel.TLabel").grid(row=3, column=0, sticky="w", pady=(12, 4))
        player_entry = ttk.Entry(frame, textvariable=player_var)
        player_entry.grid(row=4, column=0, sticky="ew", padx=(0, 10))
        ttk.Label(frame, text="Category", style="Panel.TLabel").grid(row=3, column=1, sticky="w", pady=(12, 4))
        category_combo = ttk.Combobox(frame, textvariable=category_var, values=ASSIGNMENT_CATEGORY_OPTIONS, state="readonly", width=18)
        category_combo.grid(row=4, column=1, sticky="ew", padx=(0, 10))

        entry_by_iid: dict[str, dict[str, object]] = {}

        def refresh_tree() -> None:
            tree.delete(*tree.get_children())
            entry_by_iid.clear()
            latest = self._load_unassigned_players()
            for key, entry in sorted(latest.items(), key=lambda item: str(item[1].get("last_seen") or ""), reverse=True):
                iid = str(key)
                entry_by_iid[iid] = entry
                tree.insert("", tk.END, iid=iid, values=(
                    entry.get("player") or "",
                    entry.get("count") or 0,
                    entry.get("last_seen") or "",
                    entry.get("last_title") or "",
                ))

        def select_entry(_event=None) -> None:
            selected = tree.selection()
            if not selected:
                return
            iid = selected[0]
            selected_key.set(iid)
            entry = entry_by_iid.get(iid, {})
            player_var.set(str(entry.get("player") or ""))

        def search_selected() -> None:
            key = selected_key.get()
            entry = entry_by_iid.get(key, {})
            query = " ".join(part for part in (player_var.get().strip(), "sports cards", str(entry.get("last_title") or "")) if part)
            if not query.strip():
                return
            webbrowser.open(f"https://www.google.com/search?q={urllib.parse.quote_plus(query)}")

        def save_mapping() -> None:
            player = player_var.get().strip()
            category = category_var.get().strip()
            key = selected_key.get()
            if not player or not category or not key:
                messagebox.showinfo("Unassigned player", "Choose an entry, enter the player name, and choose a category.")
                return
            self.save_player_category_override(player, category, key)
            refresh_tree()

        def remove_entry() -> None:
            key = selected_key.get()
            if not key:
                return
            with shared_lock(CARD_PIPELINE_DIR, "unassigned-players", self.lucas_identity):
                latest = self._load_unassigned_players()
                latest.pop(key, None)
                self._save_unassigned_players(latest)
            selected_key.set("")
            player_var.set("")
            refresh_tree()

        def auto_categorize_all() -> None:
            auto_button.configure(state=tk.DISABLED)
            self.status_var.set("Auto-categorizing unassigned players...")

            def worker() -> None:
                try:
                    result = self._auto_categorize_unassigned_players()
                except Exception as error:
                    self.after(0, lambda: finish_auto({"resolved": 0, "unresolved": 0, "errors": [str(error)]}))
                    return
                self.after(0, lambda: finish_auto(result))

            def finish_auto(result: dict[str, object]) -> None:
                auto_button.configure(state=tk.NORMAL)
                refresh_tree()
                resolved = int(result.get("resolved") or 0)
                unresolved = int(result.get("unresolved") or 0)
                errors = list(result.get("errors") or [])
                self.assignment_config_status.set(self._assignment_config_status())
                if resolved:
                    self._schedule_assignment_recommendations(delay_ms=50)
                if errors:
                    self.status_var.set(f"Auto-categorized {resolved}; {unresolved} left; {len(errors)} search issue(s).")
                    messagebox.showwarning("Auto Categorize", "\n".join([f"Resolved: {resolved}", f"Left: {unresolved}", "", "Issues:", *[str(error) for error in errors[:8]]]))
                else:
                    self.status_var.set(f"Auto-categorized {resolved} unassigned player(s); {unresolved} left.")
                    messagebox.showinfo("Auto Categorize", f"Resolved: {resolved}\nLeft unresolved: {unresolved}")

            threading.Thread(target=worker, daemon=True).start()

        tree.bind("<<TreeviewSelect>>", select_entry)
        ttk.Button(frame, text="Web Search", command=search_selected, style="Soft.TButton").grid(row=4, column=2, sticky="ew", padx=(0, 10))
        ttk.Button(frame, text="Save Category", command=save_mapping, style="Primary.TButton").grid(row=4, column=3, sticky="ew")
        buttons = ttk.Frame(frame, style="Panel.TFrame")
        buttons.grid(row=5, column=0, columnspan=4, sticky="e", pady=(12, 0))
        auto_button = ttk.Button(buttons, text="Auto Categorize All", command=auto_categorize_all, style="Primary.TButton")
        auto_button.pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(buttons, text="Remove", command=remove_entry, style="Soft.TButton").pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(buttons, text="Close", command=popup.destroy, style="Soft.TButton").pack(side=tk.LEFT)
        frame.columnconfigure(0, weight=1)
        frame.columnconfigure(3, weight=1)
        frame.rowconfigure(2, weight=1)
        refresh_tree()
        if tree.get_children():
            first = tree.get_children()[0]
            tree.selection_set(first)
            tree.focus(first)
            select_entry()

    def save_player_category_override(self, player: str, category: str, unassigned_key: str = "") -> None:
        self._write_player_category_override(player, category)
        self.assignment_engine = AssignmentEngine.load()
        if unassigned_key:
            with shared_lock(CARD_PIPELINE_DIR, "unassigned-players", self.lucas_identity):
                entries = self._load_unassigned_players()
                entries.pop(unassigned_key, None)
                self._save_unassigned_players(entries)
        self.assignment_config_status.set(self._assignment_config_status())
        self._schedule_assignment_recommendations(delay_ms=50)
        self.status_var.set(f"Saved {player} as {category}. Assignment will recalculate.")

    def _write_player_category_override(self, player: str, category: str) -> None:
        with shared_lock(CARD_PIPELINE_DIR, "assignment-player-overrides", self.lucas_identity):
            try:
                payload = json.loads(PLAYER_OVERRIDES_PATH.read_text(encoding="utf-8"))
            except Exception:
                payload = {}
            players = payload.get("players") if isinstance(payload, dict) else {}
            if not isinstance(players, dict):
                players = {}
            players[player] = {"sport": category, "displayName": player}
            payload = {"players": players}
            PLAYER_OVERRIDES_PATH.parent.mkdir(parents=True, exist_ok=True)
            atomic_write_json(PLAYER_OVERRIDES_PATH, payload)
        assignment_engine.add_player_sport_hint(player, category, display_name=player)

    def _queue_assignment_recommendations(self) -> None:
        rows = [*self.state.rows, *self.review_rows]
        if not rows or not self.assignment_engine.companies:
            self.assignment_progress_value.set(0)
            return
        self.assignment_recommendation_after_id = None
        self.assignment_recommendation_job += 1
        job_id = self.assignment_recommendation_job
        self.assignment_recommendation_running = True
        self.assignment_progress_value.set(0)
        total = len(rows)
        self.review_status.set(f"Calculating assignment recommendations: 0/{total}...")
        self.status_var.set("Calculating assignment recommendations...")
        threading.Thread(target=self._assignment_recommendations_worker, args=(job_id, rows), daemon=True).start()

    def _schedule_assignment_recommendations(self, delay_ms: int = 700) -> None:
        if not self.assignment_engine.companies:
            self.assignment_progress_value.set(0)
            return
        if self.assignment_recommendation_after_id is not None:
            try:
                self.after_cancel(self.assignment_recommendation_after_id)
            except tk.TclError:
                pass
        self.assignment_recommendation_after_id = self.after(delay_ms, self._queue_assignment_recommendations)

    def _assignment_recommendations_worker(self, job_id: int, rows: list[WorkbookRow]) -> None:
        total = len(rows)
        results: list[tuple[int, str, float | None]] = []
        progress_step = max(1, total // 25)
        try:
            for index, row in enumerate(rows, start=1):
                recommendation = self.assignment_engine.recommend(row, person=getattr(self, "_assignment_person_for_row", lambda _row: "")(row))
                results.append((id(row), recommendation.company, recommendation.payout))
                if index == total or index % progress_step == 0:
                    self.events.put(("assignment_recommendations_progress", {"job_id": job_id, "done": index, "total": total}))
        except Exception as error:
            self.events.put(("assignment_recommendations_error", {"job_id": job_id, "error": str(error)}))
            return
        self.events.put(("assignment_recommendations_done", {"job_id": job_id, "total": total, "results": results}))

    def _assignment_person_for_row(self, row: WorkbookRow) -> str:
        selected_working_sheet = getattr(self, "selected_working_sheet", None)
        source_sheet = (
            getattr(self, "comp_sheet_sources", {}).get(row.excel_row)
            or getattr(self, "review_sheet_sources", {}).get(row.excel_row)
            or getattr(self, "intake_sheet_sources", {}).get(row.excel_row)
            or getattr(self, "row_sources", {}).get(row.excel_row)
            or getattr(self, "review_sources", {}).get(row.excel_row)
            or getattr(self, "intake_sources", {}).get(row.excel_row)
            or (selected_working_sheet.get() if selected_working_sheet is not None else "")
            or ""
        )
        source_name = Path(str(source_sheet or "")).name
        if not source_name or source_name == "NO SHEET FOUND":
            return ""
        for stage in ("Working", "Incoming", "Received"):
            marker = self.home_sheet_markers.get(self._home_sheet_key(stage, source_name), {})
            person = str(marker.get("assigned_person") or "").strip()
            if person:
                return person
        for key, marker in self.home_sheet_markers.items():
            _stage, name = self._split_home_sheet_key(key)
            if Path(name).name == source_name:
                person = str(marker.get("assigned_person") or "").strip()
                if person:
                    return person
        return ""

    def _apply_assignment_recommendation_results(self, payload: dict[str, object]) -> None:
        if int(payload.get("job_id") or 0) != self.assignment_recommendation_job:
            return
        results = {
            int(row_id): (str(company or ""), payout)
            for row_id, company, payout in list(payload.get("results") or [])
        }
        filled = 0
        comp_rows_updated = False
        unresolved_rows: list[WorkbookRow] = []
        state_row_ids = {id(row) for row in self.state.rows}
        for row in [*self.state.rows, *self.review_rows]:
            company, payout = results.get(id(row), ("", None))
            if payout is not None:
                row.best_company = company
                row.estimated_payout = payout
                filled += 1
                if id(row) in state_row_ids:
                    comp_rows_updated = True
            else:
                row.best_company = NO_COMPANY_TAKES_LABEL
                row.estimated_payout = None
                unresolved_rows.append(row)
        self._record_unassigned_players(unresolved_rows)
        if comp_rows_updated:
            self.comp_output_saved = False
        total = int(payload.get("total") or 0)
        self.assignment_recommendation_running = False
        self.assignment_progress_value.set(100 if total else 0)
        self.review_status.set(f"Assignment recommendations complete: {filled}/{total} row(s) populated.")
        self.status_var.set(f"Assignment recommendations complete: {filled}/{total} row(s) populated.")
        self._refresh_table(schedule_recommendations=False)

    def _update_assignment_recommendation_progress(self, payload: dict[str, object]) -> None:
        if int(payload.get("job_id") or 0) != self.assignment_recommendation_job:
            return
        done = int(payload.get("done") or 0)
        total = int(payload.get("total") or 0)
        percent = (done / total * 100) if total else 0
        self.assignment_progress_value.set(percent)
        self.review_status.set(f"Calculating assignment recommendations: {done}/{total}...")

    def _handle_assignment_recommendation_error(self, payload: dict[str, object]) -> None:
        if int(payload.get("job_id") or 0) != self.assignment_recommendation_job:
            return
        error = str(payload.get("error") or "Unknown error")
        self.assignment_recommendation_running = False
        self.assignment_progress_value.set(0)
        self.review_status.set(f"Assignment recommendations failed: {error}")
        self.status_var.set(f"Assignment recommendations failed: {error}")

    def reload_assignment_rules(self) -> None:
        self._load_player_overrides()
        self.assignment_engine = AssignmentEngine.load()
        self._refresh_keep_source_registry()
        self._ensure_company_sheet_folders()
        self._ensure_weekly_company_sheets_due()
        self.assignment_config_status.set(self._assignment_config_status())
        self._refresh_table(schedule_recommendations=True)
        self.review_status.set("Assignment rules reloaded.")
        self.status_var.set("Assignment rules reloaded.")

    def open_assignment_rules(self) -> None:
        open_assignment_rules_dialog(self, CARD_PIPELINE_DIR, self.reload_assignment_rules)

    def _assignment_config_status(self) -> str:
        if self.assignment_engine.error:
            return f"Assignment config error: {self.assignment_engine.error}"
        count = len(self.assignment_engine.companies)
        if not count:
            return "Assignment companies: none configured. Add assignment_companies.json to enable best-company payouts."
        return f"Assignment companies loaded: {count}"

    def _refresh_table(self, schedule_recommendations: bool = False) -> None:
        self._render_rows(self.intake_tree, self.intake_rows, self.intake_sources)
        self._render_rows(self.comp_tree, self.state.rows, self.row_sources, self.comp_sheet_sources)
        self._render_rows(self.receive_tree, self.review_rows, self.review_sources, self.review_sheet_sources)
        self._render_rows(self.review_tree, self.review_rows, self.review_sources, self.review_sheet_sources)
        completed = sum(1 for row in self.state.rows if row.card_ladder_value is not None)
        self.summary_var.set(f"{len(self.intake_rows)} create rows | Loaded comp rows: {len(self.state.rows)} | Card Ladder values: {completed}")
        if schedule_recommendations:
            self._schedule_assignment_recommendations()

    def _render_rows(self, tree: ttk.Treeview, rows: list[WorkbookRow], sources: dict[int, str], sheet_sources: dict[int, str] | None = None) -> None:
        self._remember_column_widths(tree)
        tree.delete(*tree.get_children())
        duplicate_certs = self._duplicate_certs(rows)
        columns = self._tree_columns(tree)
        for row in rows:
            tags = []
            if row.cert_number and row.cert_number in duplicate_certs:
                tags.append("duplicate_cert")
            if (sheet_sources or {}).get(row.excel_row) == "NO SHEET FOUND":
                tags.append("no_sheet_found")
            tree.insert(
                "",
                tk.END,
                iid=str(row.excel_row),
                tags=tuple(tags),
                values=tuple(self._row_display_value(row, col, sources, sheet_sources) for col in columns),
            )
        add_row_iid = ""
        should_show_add_row = False
        if tree is self.intake_tree and self.input_mode.get() == "Manual Entry":
            add_row_iid = ADD_INTAKE_ROW_IID
            should_show_add_row = True
        elif self._is_receive_tree(tree) and self.review_mode.get() == "Manual Receive":
            add_row_iid = ADD_REVIEW_ROW_IID
            should_show_add_row = True
        if should_show_add_row:
            add_values = []
            for col in columns:
                if col == "excel_row":
                    add_values.append("+")
                elif col == "card_title":
                    add_values.append("Add row")
                else:
                    add_values.append("")
            tree.insert(
                "",
                tk.END,
                iid=add_row_iid,
                tags=("add_review_row",),
                values=tuple(add_values),
            )
        self._restore_column_widths(tree)

    def _tree_columns(self, tree: ttk.Treeview) -> tuple[str, ...]:
        return tuple(getattr(tree, "_display_columns", DISPLAY_COLUMNS))

    def _is_receive_tree(self, tree: ttk.Treeview) -> bool:
        return hasattr(self, "receive_tree") and tree is self.receive_tree

    def _is_review_row_tree(self, tree: ttk.Treeview) -> bool:
        return self._is_receive_tree(tree) or (hasattr(self, "review_tree") and tree is self.review_tree)

    def _row_display_value(
        self,
        row: WorkbookRow,
        column: str,
        sources: dict[int, str],
        sheet_sources: dict[int, str] | None,
    ) -> object:
        if column == "excel_row":
            return row.excel_row
        if column == "source":
            return sources.get(row.excel_row, "")
        if column == "sheet_source":
            return (sheet_sources or {}).get(row.excel_row, "")
        if column == "cert_number":
            return row.cert_number
        if column == "grader":
            return row.grader
        if column == "card_title":
            return row.card_title
        if column == "purchase_price":
            return format_money(row.existing_value if isinstance(row.existing_value, (int, float)) else None)
        if column == "card_ladder_value":
            return format_money(row.card_ladder_value)
        if column == "card_ladder_comps_average":
            return format_money(row.card_ladder_comps_average)
        if column == "cy_value":
            return format_money(row.cy_value)
        if column == "cy_confidence":
            return "" if row.cy_confidence is None else row.cy_confidence
        if column == "best_company":
            return row.best_company
        if column == "estimated_payout":
            return format_money(row.estimated_payout)
        if column == "status":
            return row.status
        if column == "company_pile":
            return "[x]" if row.company_pile else "[ ]"
        return ""

    def _delete_selected_table_rows(self, event) -> str | None:
        tree = event.widget
        if tree is self.intake_tree:
            self.delete_selected_intake_rows()
            return "break"
        if self._is_review_row_tree(tree):
            self.delete_selected_review_rows()
            return "break"
        return None

    def _delete_selected_rows(
        self,
        tree: ttk.Treeview,
        rows: list[WorkbookRow],
        sources: dict[int, str],
        sheet_sources: dict[int, str],
    ) -> int:
        selected_rows = {
            int(iid)
            for iid in tree.selection()
            if str(iid).isdigit() and str(iid) not in {ADD_INTAKE_ROW_IID, ADD_REVIEW_ROW_IID}
        }
        if not selected_rows:
            return 0
        remaining: list[WorkbookRow] = []
        new_sources: dict[int, str] = {}
        new_sheet_sources: dict[int, str] = {}
        for next_excel_row, row in enumerate((row for row in rows if row.excel_row not in selected_rows), start=2):
            old_excel_row = row.excel_row
            row.excel_row = next_excel_row
            remaining.append(row)
            if old_excel_row in sources:
                new_sources[next_excel_row] = sources[old_excel_row]
            if old_excel_row in sheet_sources:
                new_sheet_sources[next_excel_row] = sheet_sources[old_excel_row]
        if tree is self.intake_tree:
            self.intake_rows = remaining
            self.intake_sources = new_sources
            self.intake_sheet_sources = new_sheet_sources
        elif self._is_review_row_tree(tree):
            self.review_rows = remaining
            self.review_sources = new_sources
            self.review_sheet_sources = new_sheet_sources
        else:
            return 0
        self._cancel_cell_edit()
        self._refresh_table(schedule_recommendations=(tree is self.comp_tree or tree is self.review_tree))
        return len(selected_rows)

    def _duplicate_certs(self, rows: list[WorkbookRow]) -> set[str]:
        counts: dict[str, int] = {}
        for row in rows:
            cert = str(row.cert_number or "").strip().upper()
            if not cert:
                continue
            counts[cert] = counts.get(cert, 0) + 1
        return {cert for cert, count in counts.items() if count > 1}

    def _remember_column_widths(self, tree: ttk.Treeview) -> None:
        widths = self.column_widths_by_tree.setdefault(id(tree), {})
        for col in self._tree_columns(tree):
            try:
                widths[col] = int(tree.column(col, "width"))
            except tk.TclError:
                pass

    def _restore_column_widths(self, tree: ttk.Treeview) -> None:
        widths = self.column_widths_by_tree.get(id(tree), {})
        for col in self._tree_columns(tree):
            if col in widths:
                tree.column(col, width=widths[col])

    def _select_excel_row(self, excel_row: int) -> None:
        iid = str(excel_row)
        if self.intake_tree.exists(iid):
            self.intake_tree.selection_set(iid)
            self.intake_tree.focus(iid)
            self.intake_tree.see(iid)

    def _handle_table_click(self, event):
        tree = event.widget
        row_id = tree.identify_row(event.y)
        column_id = tree.identify_column(event.x)
        if tree is self.intake_tree and row_id == ADD_INTAKE_ROW_IID:
            self.add_manual_intake_row()
            return "break"
        if self._is_receive_tree(tree) and row_id == ADD_REVIEW_ROW_IID:
            self.add_manual_review_row()
            return "break"
        if self._is_receive_tree(tree) and row_id and column_id:
            column_index = int(column_id.replace("#", "")) - 1
            columns = self._tree_columns(tree)
            if 0 <= column_index < len(columns) and columns[column_index] == "company_pile":
                self._toggle_company_pile(row_id)
                return "break"
        return None

    def _toggle_company_pile(self, row_id: str) -> None:
        if not str(row_id).isdigit():
            return
        excel_row = int(row_id)
        for row in self.review_rows:
            if row.excel_row != excel_row:
                continue
            row.company_pile = not row.company_pile
            self._refresh_table()
            self.review_status.set("Company pile checked." if row.company_pile else "Company pile unchecked.")
            return

    def _begin_cell_edit(self, event) -> None:
        tree = event.widget
        row_id = tree.identify_row(event.y)
        column_id = tree.identify_column(event.x)
        if tree is self.intake_tree and row_id == ADD_INTAKE_ROW_IID:
            self.add_manual_intake_row()
            return
        if self._is_receive_tree(tree) and row_id == ADD_REVIEW_ROW_IID:
            self.add_manual_review_row()
            return
        if not row_id or not column_id:
            return
        column_index = int(column_id.replace("#", "")) - 1
        columns = self._tree_columns(tree)
        if column_index < 0 or column_index >= len(columns):
            return
        column = columns[column_index]
        if column not in EDITABLE_COLUMNS:
            return
        bbox = tree.bbox(row_id, column_id)
        if not bbox:
            return
        self._cancel_cell_edit()
        x, y, width, height = bbox
        current = tree.set(row_id, column)
        editor = ttk.Entry(tree)
        editor.insert(0, current)
        editor.select_range(0, tk.END)
        editor.place(x=x, y=y, width=width, height=height)
        editor.focus_set()
        self.cell_editor = editor
        self.cell_edit = (tree, row_id, column)
        editor.bind("<Return>", lambda _event: self._commit_cell_edit())
        editor.bind("<KP_Enter>", lambda _event: self._commit_cell_edit())
        editor.bind("<Escape>", lambda _event: self._cancel_cell_edit())
        editor.bind("<FocusOut>", lambda _event: self._commit_cell_edit())

    def _commit_cell_edit(self) -> None:
        if not self.cell_editor or not self.cell_edit:
            return
        tree, row_id, column = self.cell_edit
        value = self.cell_editor.get()
        current = tree.set(row_id, column)
        self._destroy_cell_editor()
        if value.strip() == str(current or "").strip():
            return
        excel_row = int(row_id)
        self._apply_cell_value(tree, excel_row, column, value)
        if tree is self.comp_tree:
            self.comp_output_saved = False
        self._refresh_table(schedule_recommendations=self._edit_affects_assignment(tree, column))
        if tree.exists(row_id):
            tree.selection_set(row_id)
            tree.focus(row_id)
            tree.see(row_id)
        self.status_var.set(f"Updated row {excel_row}.")

    def _edit_affects_assignment(self, tree: ttk.Treeview, column: str) -> bool:
        if tree is not self.comp_tree and tree is not self.review_tree:
            return False
        return column in {
            "cert_number",
            "grader",
            "card_title",
            "card_ladder_value",
            "card_ladder_comps_average",
            "cy_value",
        }

    def _cancel_cell_edit(self) -> None:
        self._destroy_cell_editor()

    def _destroy_cell_editor(self) -> None:
        if self.cell_editor is not None:
            try:
                self.cell_editor.destroy()
            except tk.TclError:
                pass
        self.cell_editor = None
        self.cell_edit = None

    def _apply_cell_value(self, tree: ttk.Treeview, excel_row: int, column: str, value: str) -> None:
        clean_value = value.strip()
        if tree is self.comp_tree:
            target_rows = self.state.rows
            target_sources = self.row_sources
            target_sheet_sources = self.comp_sheet_sources
        elif self._is_review_row_tree(tree):
            target_rows = self.review_rows
            target_sources = self.review_sources
            target_sheet_sources = self.review_sheet_sources
        else:
            target_rows = self.intake_rows
            target_sources = self.intake_sources
            target_sheet_sources = self.intake_sheet_sources
        if column == "source":
            target_sources[excel_row] = clean_value
            return
        if column == "sheet_source":
            target_sheet_sources[excel_row] = clean_value
            return
        for row in target_rows:
            if row.excel_row != excel_row:
                continue
            previous_cert = scan_to_cert(row.cert_number)
            if column == "cert_number":
                row.cert_number = scan_to_cert(clean_value)
            elif column == "grader":
                row.grader = normalize_grader(clean_value) or clean_value.upper()
            elif column == "card_title":
                row.card_title = clean_value
                inferred = infer_grader(row.card_title)
                if inferred:
                    row.grader = inferred
            elif column == "purchase_price":
                row.existing_value = self._parse_money_text(clean_value)
            elif column == "card_ladder_value":
                row.card_ladder_value = self._parse_money_text(clean_value)
            elif column == "card_ladder_comps_average":
                row.card_ladder_comps_average = self._parse_money_text(clean_value)
            elif column == "cy_value":
                row.cy_value = self._parse_money_text(clean_value)
            elif column == "cy_confidence":
                row.cy_confidence = clean_value
            row.status = "Ready" if row.cert_number and row.grader else "Needs setup"
            if tree is self.intake_tree:
                if column == "purchase_price":
                    setattr(row, "_seller_terms_base_purchase", row.existing_value)
                if column in {"purchase_price", "card_ladder_value", "card_ladder_comps_average", "cy_value"}:
                    self.apply_create_seller_terms(show_status=False)
            if self._is_review_row_tree(tree) and column == "cert_number" and scan_to_cert(row.cert_number) != previous_cert:
                match = self._incoming_match(row.cert_number)
                target_sheet_sources[excel_row] = str(match.get("sheet") or "NO SHEET FOUND")
                if match:
                    row.status = "Received"
                    if is_placeholder_title(row.card_title, row.grader) and match.get("card_title"):
                        row.card_title = str(match.get("card_title") or "")
                    if row.existing_value is None and match.get("purchase_price") is not None:
                        row.existing_value = match.get("purchase_price")
                    if row.card_ladder_value is None and match.get("card_ladder_value") is not None:
                        row.card_ladder_value = match.get("card_ladder_value")
                    if row.card_ladder_comps_average is None and match.get("card_ladder_comps_average") is not None:
                        row.card_ladder_comps_average = match.get("card_ladder_comps_average")
                    if row.cy_value is None and match.get("cy_value") is not None:
                        row.cy_value = match.get("cy_value")
                    if row.cy_confidence is None and match.get("cy_confidence") is not None:
                        row.cy_confidence = match.get("cy_confidence")
                    if not row.card_ladder_comps and match.get("card_ladder_comps"):
                        row.card_ladder_comps = str(match.get("card_ladder_comps") or "")
                elif row.cert_number:
                    target_sheet_sources[excel_row] = "NO SHEET FOUND"
                    row.status = "Received - no incoming match"
            if not row.cert_number:
                row.notes = "Missing cert"
            elif not row.grader:
                row.notes = "Missing grader"
            elif row.notes in {"Missing cert", "Missing grader", "Missing cert or grader"}:
                row.notes = ""
            return

    def _poll_events(self) -> None:
        try:
            while True:
                event = self.events.get_nowait()
                if event == "refresh":
                    self._refresh_table()
                elif event == "comp_refresh":
                    self.comp_output_saved = False
                    self._refresh_table(schedule_recommendations=True)
                elif isinstance(event, tuple):
                    kind, payload = event
                    if kind == "photo_rows":
                        self._append_rows(payload)
                    elif kind == "photo_status":
                        self.photo_status.set(str(payload))
                        self.status_var.set(str(payload))
                    elif kind == "review_rows":
                        self._append_review_rows(payload)
                    elif kind == "review_status":
                        self.review_photo_status.set(str(payload))
                        self.review_status.set(str(payload))
                        self.status_var.set(str(payload))
                    elif kind == "startup_refresh":
                        self._apply_startup_refresh(payload)
                    elif kind == "load_working_sheet_done":
                        self._apply_loaded_working_sheet(str(payload.get("name") or ""), list(payload.get("rows") or []))
                    elif kind == "load_working_sheet_error":
                        self.status_var.set(f"Working sheet load failed: {payload.get('error')}")
                        messagebox.showerror("Load failed", str(payload.get("error") or "Unknown error"))
                    elif kind == "load_received_sheet_done":
                        self._apply_loaded_received_sheet(str(payload.get("name") or ""), list(payload.get("rows") or []))
                    elif kind == "load_received_sheet_error":
                        self.review_status.set(f"Received sheet load failed: {payload.get('error')}")
                        self.status_var.set(f"Received sheet load failed: {payload.get('error')}")
                        messagebox.showerror("Load failed", str(payload.get("error") or "Unknown error"))
                    elif kind == "assignment_recommendations_progress":
                        self._update_assignment_recommendation_progress(payload)
                    elif kind == "assignment_recommendations_done":
                        self._apply_assignment_recommendation_results(payload)
                    elif kind == "assignment_recommendations_error":
                        self._handle_assignment_recommendation_error(payload)
                    elif kind == "inventory_refresh":
                        self.refresh_inventory_tab(enrich=True)
                        self.status_var.set(str(payload))
                    elif kind == "profit_refresh":
                        self.refresh_profit_tab()
                        self.refresh_payouts_tab()
                        self.status_var.set(str(payload))
        except queue.Empty:
            pass
        self.after(200, self._poll_events)

    def _parse_money_text(self, value: str) -> float | None:
        text = value.strip().replace("$", "").replace(",", "")
        if not text:
            return None
        try:
            return float(text)
        except ValueError:
            return None

    def _arm_scanner(self) -> None:
        if self.input_mode.get() != "Barcode Scanner" or self.scan_entry is None:
            return
        try:
            self.scan_entry.focus_set()
            self.scan_entry.icursor(tk.END)
        except tk.TclError:
            pass

    def destroy(self) -> None:
        self.bridge.stop()
        super().destroy()


def first_number(*values: float | None) -> float | None:
    for value in values:
        if value is not None:
            return value
    return None


def is_placeholder_title(card_title: str, grader: str) -> bool:
    title = str(card_title or "").strip()
    company = str(grader or "").strip()
    if not title:
        return True
    return bool(company and title.upper() == company.upper())


if __name__ == "__main__":
    CardPipelineApp().mainloop()
