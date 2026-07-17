from __future__ import annotations

import csv
import json
import os
import re
import subprocess
import sys
import time
import tkinter as tk
import urllib.parse
import webbrowser
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog, ttk
from typing import Any, Callable

from assignment_engine import (
    CONFIG_PATH,
    gsheet_shortcut_url,
    is_google_keep_url,
    keep_note_cache_path,
    load_gsheet_shortcut,
    normalize_source_value,
    read_source_text,
    safe_filename,
)
from google_sheets_import import LAST_OAUTH_DIAGNOSTICS, TOKEN_PATH, authorize_google_sheets
from lucas_diagnostics import diagnostic_json, google_status_lines


GRADE_COMPANIES = ("psa", "bgs", "sgc", "cgc")
SPORT_OPTIONS = (
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
RULE_SOURCE_LABELS = {
    "manual": "Manual rules",
    "keep_file": "Google Keep note / local text",
    "sheet_file": "Google Sheets file / URL",
}
PAYOUT_SOURCE_LABELS = {
    "manual": "Manual payout tiers",
    "file": "Local payout file",
}
VALUE_SOURCE_LABELS = {
    "comps": "Comps",
    "card_ladder": "Card Ladder value",
    "cy_estimate": "CY Estimate",
}
COMPANY_RESET_WEEKDAYS = ("Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday")
DEFAULT_COMPANY_RESET_WEEKDAY = "Monday"
DEFAULT_COMPANY_RESET_TIME = "20:00"
SELLER_TERMS_FIELDS = ("Seller", "Sheet Type", "Seller Rate", "Deduction")
SELLER_TERMS_FIELD_LABELS = {
    "Seller Rate": "Seller Rate %",
    "Deduction": "Deduction %",
}


def build_scrollable_dialog_body(parent: tk.Misc, style_name: str, bg: str = "#121212", padding: int = 16) -> ttk.Frame:
    outer = ttk.Frame(parent, style=style_name)
    outer.pack(fill=tk.BOTH, expand=True)
    outer.columnconfigure(0, weight=1)
    outer.rowconfigure(0, weight=1)
    canvas = tk.Canvas(outer, bg=bg, highlightthickness=0, borderwidth=0)
    canvas.grid(row=0, column=0, sticky="nsew")
    y_scroll = ttk.Scrollbar(outer, orient=tk.VERTICAL, command=canvas.yview)
    y_scroll.grid(row=0, column=1, sticky="ns")
    x_scroll = ttk.Scrollbar(outer, orient=tk.HORIZONTAL, command=canvas.xview)
    x_scroll.grid(row=1, column=0, sticky="ew")
    canvas.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)
    body = ttk.Frame(canvas, style=style_name, padding=padding)
    window_id = canvas.create_window((0, 0), window=body, anchor="nw")

    def update_scrollregion(_event: tk.Event | None = None) -> None:
        canvas.configure(scrollregion=canvas.bbox("all"))

    def resize_window(event: tk.Event) -> None:
        required = body.winfo_reqwidth()
        canvas.itemconfigure(window_id, width=max(event.width, required))
        update_scrollregion()

    def on_mousewheel(event: tk.Event) -> str:
        delta = int(getattr(event, "delta", 0) or 0)
        if delta:
            canvas.yview_scroll(int(-1 * (delta / 120)), "units")
        return "break"

    def on_shift_mousewheel(event: tk.Event) -> str:
        delta = int(getattr(event, "delta", 0) or 0)
        if delta:
            canvas.xview_scroll(int(-1 * (delta / 120)), "units")
        return "break"

    body.bind("<Configure>", update_scrollregion)
    canvas.bind("<Configure>", resize_window)
    for widget in (outer, canvas, body):
        widget.bind("<MouseWheel>", on_mousewheel, add="+")
        widget.bind("<Shift-MouseWheel>", on_shift_mousewheel, add="+")
    return body


def bind_single_paste(widget: tk.Widget) -> tk.Widget:
    def handle_paste(_event: tk.Event) -> str:
        now = time.monotonic()
        try:
            text = widget.clipboard_get()
        except tk.TclError:
            return "break"
        last = getattr(widget, "_lucas_last_paste", None)
        if last and now - last[0] < 0.15 and text == last[1]:
            return "break"
        setattr(widget, "_lucas_last_paste", (now, text))
        try:
            state = str(widget.cget("state") or "")
            if state == "disabled":
                return "break"
            if state == "readonly" and isinstance(widget, ttk.Combobox):
                values = [str(value) for value in widget.cget("values")]
                if text in values:
                    widget.set(text)
                return "break"
            try:
                widget.delete(tk.SEL_FIRST, tk.SEL_LAST)
            except tk.TclError:
                pass
            widget.insert(tk.INSERT, text)
        except tk.TclError:
            pass
        return "break"

    widget.bind("<<Paste>>", handle_paste)
    widget.bind("<Command-v>", handle_paste)
    widget.bind("<Control-v>", handle_paste)
    return widget


def read_seller_terms_rows(seller_terms_path: Path) -> list[dict[str, str]]:
    if not seller_terms_path.exists():
        return []
    with seller_terms_path.open("r", encoding="utf-8-sig", newline="") as handle:
        rows = []
        for row in csv.DictReader(handle):
            normalized = {re.sub(r"[^a-z0-9]+", "", str(key or "").lower()): value for key, value in row.items()}
            rows.append(
                {
                    "Seller": str(normalized.get("seller") or normalized.get("person") or normalized.get("name") or "").strip(),
                    "Sheet Type": str(normalized.get("sheettype") or normalized.get("type") or normalized.get("company") or "").strip(),
                    "Seller Rate": str(normalized.get("sellerrate") or normalized.get("rate") or normalized.get("payout") or normalized.get("percentage") or "").strip(),
                    "Deduction": str(normalized.get("deduction") or normalized.get("sellerdeduction") or normalized.get("deductionpercent") or normalized.get("deductionpercentage") or "").strip(),
                }
            )
    return rows


def write_seller_terms_rows(seller_terms_path: Path, rows: list[dict[str, str]]) -> None:
    seller_terms_path.parent.mkdir(parents=True, exist_ok=True)
    with seller_terms_path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=list(SELLER_TERMS_FIELDS))
        writer.writeheader()
        for row in rows:
            writer.writerow({field: str(row.get(field) or "").strip() for field in SELLER_TERMS_FIELDS})


def seller_terms_percent_display(value: object) -> str:
    raw = str(value or "").strip()
    if not raw:
        return ""
    has_percent = raw.endswith("%")
    text = raw[:-1].strip() if has_percent else raw
    try:
        numeric = float(text)
    except ValueError:
        return raw
    if not has_percent and 0 <= numeric <= 1:
        numeric *= 100
    return f"{numeric:g}"


def seller_terms_percent_input_is_number(value: object) -> bool:
    raw = str(value or "").strip()
    if not raw:
        return True
    if "%" in raw:
        return False
    return bool(re.fullmatch(r"\d+(?:\.\d+)?", raw))


def seller_terms_rate(value: object) -> float | None:
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


def normalize_company_reset_weekday(value: object) -> str:
    text = str(value or DEFAULT_COMPANY_RESET_WEEKDAY).strip().title()
    return text if text in COMPANY_RESET_WEEKDAYS else DEFAULT_COMPANY_RESET_WEEKDAY


def normalize_company_reset_time(value: object) -> str:
    text = str(value or DEFAULT_COMPANY_RESET_TIME).strip()
    if not text:
        text = DEFAULT_COMPANY_RESET_TIME
    for pattern in ("%H:%M", "%H%M", "%I:%M %p", "%I %p", "%I:%M%p", "%I%p"):
        try:
            from datetime import datetime

            return datetime.strptime(text.upper(), pattern).time().strftime("%H:%M")
        except ValueError:
            continue
    raise ValueError("Use a time like 20:00 or 8:00 PM.")


def seller_terms_health_lines(seller_terms_path: Path, companies: list[dict[str, Any]]) -> list[str]:
    active_companies = {
        str(company.get("name") or "").strip().lower(): str(company.get("name") or "").strip()
        for company in companies
        if str(company.get("name") or "").strip() and company.get("active") is not False
    }
    inactive_companies = {
        str(company.get("name") or "").strip().lower(): str(company.get("name") or "").strip()
        for company in companies
        if str(company.get("name") or "").strip() and company.get("active") is False
    }
    if not seller_terms_path.exists():
        return [
            "People Rules: not found",
            f"Path: {seller_terms_path}",
            "Open People Rules when Network Mode seller payouts are needed.",
        ]
    try:
        rows = read_seller_terms_rows(seller_terms_path)
    except Exception as error:
        return [
            "People Rules: unreadable",
            f"Path: {seller_terms_path}",
            f"Error: {error}",
        ]

    issues: list[tuple[str, str]] = []
    parsed: list[str] = []
    seen: dict[tuple[str, str], int] = {}
    valid_count = 0
    for index, row in enumerate(rows, start=2):
        seller = str(row.get("Seller") or "").strip()
        sheet_type = str(row.get("Sheet Type") or "").strip()
        rate_raw = row.get("Seller Rate")
        deduction_raw = row.get("Deduction")
        rate = seller_terms_rate(rate_raw)
        deduction = seller_terms_rate(deduction_raw)
        row_errors: list[str] = []
        row_warnings: list[str] = []
        if not seller:
            row_errors.append("missing Seller")
        if not sheet_type:
            row_errors.append("missing Sheet Type")
        if str(rate_raw or "").strip() and rate is None:
            row_errors.append(f"invalid Seller Rate {rate_raw!r}")
        if str(deduction_raw or "").strip() and deduction is None:
            row_errors.append(f"invalid Deduction {deduction_raw!r}")
        if rate is None and deduction is None:
            row_errors.append("missing Seller Rate or Deduction")
        if rate is not None and deduction is not None:
            row_errors.append("use Seller Rate or Deduction, not both")
        if rate is not None and rate > 1:
            row_errors.append(f"Seller Rate parses above 100% ({rate:.0%})")
        if deduction is not None and deduction > 1:
            row_errors.append(f"Deduction parses above 100% ({deduction:.0%})")
        type_key = sheet_type.lower()
        if sheet_type and type_key not in active_companies:
            if type_key in inactive_companies:
                row_warnings.append(f"Sheet Type is inactive assignment company {inactive_companies[type_key]!r}")
            else:
                row_errors.append(f"Sheet Type {sheet_type!r} is not an assignment company")
        duplicate_key = (seller.lower(), type_key)
        if seller and sheet_type:
            previous_row = seen.get(duplicate_key)
            if previous_row:
                row_warnings.append(f"duplicate Seller/Sheet Type also appears on row {previous_row}")
            else:
                seen[duplicate_key] = index
        label = f"row {index}: {seller or '<missing seller>'} / {sheet_type or '<missing sheet type>'}"
        if row_errors:
            issues.append(("ERROR", f"{label} - {'; '.join(row_errors)}"))
            continue
        if row_warnings:
            issues.append(("WARN", f"{label} - {'; '.join(row_warnings)}"))
        valid_count += 1
        parts = []
        if rate is not None:
            parts.append(f"rate {rate:.0%}")
        if deduction is not None:
            parts.append(f"deduction {deduction:.0%}")
        parsed.append(f"{seller} / {sheet_type}: {', '.join(parts)}")

    errors = sum(1 for level, _message in issues if level == "ERROR")
    warnings = sum(1 for level, _message in issues if level == "WARN")
    lines = [
        f"People Rules: {valid_count} valid row(s), {warnings} warning(s), {errors} error(s)",
        f"Path: {seller_terms_path}",
    ]
    if not rows:
        lines.append("People Rules are empty.")
    if parsed:
        lines.append("Parsed terms:")
        lines.extend(f"- {line}" for line in parsed[:10])
        if len(parsed) > 10:
            lines.append(f"- ... {len(parsed) - 10} more")
    if issues:
        lines.append("Issues:")
        lines.extend(f"- {level}: {message}" for level, message in issues[:12])
        if len(issues) > 12:
            lines.append(f"- ... {len(issues) - 12} more")
    elif rows:
        lines.append("No seller term issues found.")
    return lines


class AssignmentRulesDialog(tk.Toplevel):
    def __init__(self, parent: tk.Tk, pipeline_root: Path, on_saved: Callable[[], None]) -> None:
        super().__init__(parent)
        self.title("Company Rules")
        self.geometry("1380x920")
        self.minsize(900, 560)
        self.transient(parent)
        self.configure(bg="#121212")
        self.on_saved = on_saved
        self.pipeline_root = Path(pipeline_root)
        self.rules_dir = self.pipeline_root / "ASSIGNMENT RULES"
        self.company_sheets_dir = self.pipeline_root / "COMPANY SHEETS"
        self.config_path = CONFIG_PATH
        self.companies = self._load_config()
        self.selected_index: int | None = None
        self.company_rows: list[tk.Frame] = []
        self.rule_rows: list[dict[str, Any]] = []
        self.payout_rows: list[dict[str, tk.StringVar]] = []

        self.company_name = tk.StringVar()
        self.company_filter_text = tk.StringVar()
        self.company_filter_state = tk.StringVar(value="all")
        self.value_source = tk.StringVar(value="comps")
        self.rule_source_mode = tk.StringVar(value="manual")
        self.rule_source_path = tk.StringVar()
        self.link_payouts_to_rule_source = tk.BooleanVar(value=False)
        self.payout_source_mode = tk.StringVar(value="manual")
        self.payout_source_path = tk.StringVar()
        self.reset_weekday = tk.StringVar(value=DEFAULT_COMPANY_RESET_WEEKDAY)
        self.reset_time = tk.StringVar(value=DEFAULT_COMPANY_RESET_TIME)
        self.manual_min_year = tk.StringVar()
        self.manual_max_year = tk.StringVar()
        self.status = tk.StringVar(value="Create or edit a company, then save.")
        self.preview_status = tk.StringVar(value="No source file selected.")
        self.google_status = tk.StringVar(value="")
        self.seller_terms_status = tk.StringVar(value="")
        self.rule_materialized_source: dict[str, Any] | None = None
        self.payout_materialized_source: dict[str, Any] | None = None

        self._configure_styles()
        self._build_ui()
        self.company_filter_text.trace_add("write", lambda *_args: self._refresh_company_list())
        self._refresh_company_list()
        if self.companies:
            self._select_company(0)
        else:
            self._new_company()

    def _configure_styles(self) -> None:
        style = ttk.Style(self)
        palette = {
            "bg": "#121212",
            "panel": "#1f1f1f",
            "panel_high": "#242424",
            "field": "#2a2a2a",
            "border": "#333333",
            "muted": "#b3b3b3",
            "text": "#ffffff",
            "button": "#1ed760",
            "button_hover": "#1fdf64",
            "soft": "#2a2a2a",
        }
        style.configure("Assign.TFrame", background=palette["bg"])
        style.configure("AssignPanel.TFrame", background=palette["panel"])
        style.configure("AssignHeader.TLabel", background=palette["bg"], foreground=palette["text"], font=("Segoe UI Semibold", 18))
        style.configure("AssignTitle.TLabel", background=palette["panel"], foreground=palette["text"], font=("Segoe UI Semibold", 11))
        style.configure("Assign.TLabel", background=palette["panel"], foreground=palette["text"])
        style.configure("AssignMuted.TLabel", background=palette["panel"], foreground=palette["muted"])
        style.configure("AssignBgMuted.TLabel", background=palette["bg"], foreground=palette["muted"])
        style.configure("Assign.TCheckbutton", background=palette["panel"], foreground=palette["text"])
        style.map("Assign.TCheckbutton", background=[("active", palette["panel"])], foreground=[("active", palette["text"])])
        style.configure("Assign.TRadiobutton", background=palette["panel"], foreground=palette["text"])
        style.map("Assign.TRadiobutton", background=[("active", palette["panel"])], foreground=[("active", palette["text"])])
        style.configure("Assign.TEntry", fieldbackground=palette["field"], foreground=palette["text"], bordercolor=palette["border"], padding=(8, 6))
        style.configure("Assign.TCombobox", fieldbackground=palette["field"], foreground=palette["text"], bordercolor=palette["border"], padding=(8, 5))
        style.configure("AssignPrimary.TButton", background=palette["button"], foreground="#000000", padding=(14, 8), font=("Segoe UI Semibold", 10))
        style.map("AssignPrimary.TButton", background=[("active", palette["button_hover"])])
        style.configure("AssignSoft.TButton", background=palette["soft"], foreground=palette["text"], padding=(12, 8), font=("Segoe UI Semibold", 10))
        style.map("AssignSoft.TButton", background=[("active", "#3a3a3a")])
        style.configure("Assign.TLabelframe", background=palette["panel"], foreground=palette["text"], bordercolor=palette["border"])
        style.configure("Assign.TLabelframe.Label", background=palette["panel"], foreground=palette["text"], font=("Segoe UI Semibold", 10))

    def _build_ui(self) -> None:
        shell = build_scrollable_dialog_body(self, "Assign.TFrame", padding=16)
        shell.columnconfigure(1, weight=1)
        shell.rowconfigure(1, weight=1)

        ttk.Label(shell, text="Company Rules", style="AssignHeader.TLabel").grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 12))

        side = ttk.Frame(shell, style="AssignPanel.TFrame", padding=12)
        side.grid(row=1, column=0, sticky="ns", padx=(0, 12))
        ttk.Label(side, text="Companies", style="AssignTitle.TLabel").pack(anchor=tk.W)
        filter_entry = bind_single_paste(ttk.Entry(side, textvariable=self.company_filter_text, style="Assign.TEntry"))
        filter_entry.pack(fill=tk.X, pady=(8, 6))
        filter_modes = ttk.Frame(side, style="AssignPanel.TFrame")
        filter_modes.pack(fill=tk.X)
        for value, label in (("all", "All"), ("active", "Active"), ("inactive", "Inactive")):
            ttk.Radiobutton(
                filter_modes,
                text=label,
                value=value,
                variable=self.company_filter_state,
                command=self._refresh_company_list,
                style="Assign.TRadiobutton",
            ).pack(side=tk.LEFT, padx=(0, 8))
        self.company_list = tk.Frame(
            side,
            width=280,
            height=560,
            bg="#1f1f1f",
            highlightthickness=1,
            highlightbackground="#333333",
            relief=tk.FLAT,
            borderwidth=0
        )
        self.company_list.pack(fill=tk.BOTH, expand=True, pady=(8, 10))
        self.company_list.pack_propagate(False)
        ttk.Button(side, text="New Company", command=self._new_company, style="AssignPrimary.TButton").pack(fill=tk.X, pady=(0, 8))
        ttk.Button(side, text="Delete Company", command=self._delete_company, style="AssignSoft.TButton").pack(fill=tk.X)

        main = ttk.Frame(shell, style="Assign.TFrame")
        main.grid(row=1, column=1, sticky="nsew")
        main.columnconfigure(0, weight=1)
        main.rowconfigure(2, weight=1)

        details = ttk.Frame(main, style="AssignPanel.TFrame", padding=12)
        details.grid(row=0, column=0, sticky="ew")
        details.columnconfigure(1, weight=1)
        ttk.Label(details, text="Company Name", style="Assign.TLabel").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        bind_single_paste(ttk.Entry(details, textvariable=self.company_name, style="Assign.TEntry")).grid(row=0, column=1, sticky="ew")
        ttk.Label(details, text="Assignment Value", style="Assign.TLabel").grid(row=1, column=0, sticky=tk.W, padx=(0, 10), pady=(8, 0))
        value_source_frame = ttk.Frame(details, style="AssignPanel.TFrame")
        value_source_frame.grid(row=1, column=1, sticky=tk.W, pady=(8, 0))
        for index, (value, label) in enumerate(VALUE_SOURCE_LABELS.items()):
            ttk.Radiobutton(
                value_source_frame,
                text=label,
                value=value,
                variable=self.value_source,
                style="Assign.TRadiobutton",
            ).grid(row=0, column=index, sticky=tk.W, padx=(0, 16))
        ttk.Label(details, text="Company Sheet Reset", style="Assign.TLabel").grid(row=2, column=0, sticky=tk.W, padx=(0, 10), pady=(8, 0))
        reset_frame = ttk.Frame(details, style="AssignPanel.TFrame")
        reset_frame.grid(row=2, column=1, sticky=tk.W, pady=(8, 0))
        bind_single_paste(ttk.Combobox(
            reset_frame,
            textvariable=self.reset_weekday,
            values=COMPANY_RESET_WEEKDAYS,
            width=14,
            state="readonly",
            style="Assign.TCombobox",
        )).grid(row=0, column=0, sticky=tk.W, padx=(0, 8))
        bind_single_paste(ttk.Entry(reset_frame, textvariable=self.reset_time, width=12, style="Assign.TEntry")).grid(row=0, column=1, sticky=tk.W)
        ttk.Label(
            reset_frame,
            text="Starts the next weekly company sheet at this weekday/time.",
            style="AssignMuted.TLabel",
        ).grid(row=0, column=2, sticky=tk.W, padx=(10, 0))

        sources = ttk.Frame(main, style="AssignPanel.TFrame", padding=12)
        sources.grid(row=1, column=0, sticky="ew", pady=(12, 0))
        sources.columnconfigure(0, weight=1)
        sources.columnconfigure(1, weight=1)
        self._build_rule_source_panel(sources).grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        self._build_payout_source_panel(sources).grid(row=0, column=1, sticky="nsew", padx=(8, 0))
        status_panels = ttk.Frame(main, style="Assign.TFrame")
        status_panels.grid(row=2, column=0, sticky="ew", pady=(12, 0))
        status_panels.columnconfigure(0, weight=1)
        status_panels.columnconfigure(1, weight=1)
        self._build_google_status_panel(status_panels).grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        self._build_seller_terms_panel(status_panels).grid(row=0, column=1, sticky="nsew", padx=(8, 0))

        self.body = ttk.Frame(main, style="Assign.TFrame")
        self.body.grid(row=3, column=0, sticky="nsew", pady=(12, 0))
        self.body.columnconfigure(0, weight=1)
        self.body.columnconfigure(1, weight=1)
        self.body.rowconfigure(0, weight=1)

        self.manual_rule_panel = ttk.Frame(self.body, style="AssignPanel.TFrame", padding=12)
        self.manual_rule_panel.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        self.manual_rule_panel.columnconfigure(0, weight=1)
        self.manual_rule_panel.rowconfigure(3, weight=1)
        rule_header = ttk.Frame(self.manual_rule_panel, style="AssignPanel.TFrame")
        rule_header.grid(row=0, column=0, sticky="ew")
        ttk.Label(rule_header, text="Manual Rule Builder", style="AssignTitle.TLabel").pack(side=tk.LEFT)
        ttk.Button(rule_header, text="Add Rule", command=self._add_rule_row, style="AssignSoft.TButton").pack(side=tk.RIGHT)
        ttk.Label(self.manual_rule_panel, text="Used when Rule Source is Manual rules.", style="AssignMuted.TLabel").grid(row=1, column=0, sticky=tk.W, pady=(3, 8))
        year_frame = ttk.Frame(self.manual_rule_panel, style="AssignPanel.TFrame")
        year_frame.grid(row=2, column=0, sticky="ew", pady=(0, 10))
        ttk.Label(year_frame, text="Company Card Year Range", style="Assign.TLabel").grid(row=0, column=0, sticky=tk.W, padx=(0, 8))
        ttk.Label(year_frame, text="Min Year", style="AssignMuted.TLabel").grid(row=0, column=1, sticky=tk.W, padx=(0, 6))
        bind_single_paste(ttk.Entry(year_frame, textvariable=self.manual_min_year, width=10, style="Assign.TEntry")).grid(row=0, column=2, sticky=tk.W)
        ttk.Label(year_frame, text="Max Year", style="AssignMuted.TLabel").grid(row=0, column=3, sticky=tk.W, padx=(14, 6))
        bind_single_paste(ttk.Entry(year_frame, textvariable=self.manual_max_year, width=10, style="Assign.TEntry")).grid(row=0, column=4, sticky=tk.W)
        ttk.Label(year_frame, text="Example: Fanatics min year 1990.", style="AssignMuted.TLabel").grid(row=0, column=5, sticky=tk.W, padx=(14, 0))
        rules_view = ttk.Frame(self.manual_rule_panel, style="AssignPanel.TFrame")
        rules_view.grid(row=3, column=0, sticky="nsew")
        rules_view.columnconfigure(0, weight=1)
        rules_view.rowconfigure(0, weight=1)
        self.rules_canvas = tk.Canvas(rules_view, bg="#1f1f1f", highlightthickness=0, borderwidth=0)
        self.rules_canvas.grid(row=0, column=0, sticky="nsew")
        rules_scrollbar = ttk.Scrollbar(rules_view, orient=tk.VERTICAL, command=self.rules_canvas.yview)
        rules_scrollbar.grid(row=0, column=1, sticky="ns")
        self.rules_canvas.configure(yscrollcommand=rules_scrollbar.set)
        self.rules_frame = ttk.Frame(self.rules_canvas, style="AssignPanel.TFrame")
        self.rules_canvas_window = self.rules_canvas.create_window((0, 0), window=self.rules_frame, anchor="nw")
        self.rules_frame.bind("<Configure>", self._update_rules_scrollregion)
        self.rules_frame.bind("<Enter>", lambda _event: self.rules_canvas.bind_all("<MouseWheel>", self._on_rules_mousewheel))
        self.rules_frame.bind("<Leave>", lambda _event: self.rules_canvas.unbind_all("<MouseWheel>"))
        self.rules_canvas.bind("<Configure>", self._resize_rules_canvas_window)
        self.rules_canvas.bind("<Enter>", lambda _event: self.rules_canvas.bind_all("<MouseWheel>", self._on_rules_mousewheel))
        self.rules_canvas.bind("<Leave>", lambda _event: self.rules_canvas.unbind_all("<MouseWheel>"))

        self.right_panel = ttk.Frame(self.body, style="Assign.TFrame")
        self.right_panel.grid(row=0, column=1, sticky="nsew", padx=(8, 0))
        self.right_panel.columnconfigure(0, weight=1)
        self.right_panel.rowconfigure(0, weight=1)
        self.payout_panel = self._build_payout_panel(self.right_panel)
        self.payout_panel.grid(row=0, column=0, sticky="nsew")

        footer = ttk.Frame(shell, style="Assign.TFrame")
        footer.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(12, 0))
        footer.columnconfigure(0, weight=1)
        ttk.Label(footer, textvariable=self.status, style="AssignBgMuted.TLabel").grid(row=0, column=0, sticky=tk.W)
        ttk.Button(footer, text="Save Company", command=self._save_company, style="AssignSoft.TButton").grid(row=0, column=1, padx=(8, 0))
        ttk.Button(footer, text="Save & Reload", command=self._save_and_reload, style="AssignPrimary.TButton").grid(row=0, column=2, padx=(8, 0))
        ttk.Button(footer, text="Close", command=self.destroy, style="AssignSoft.TButton").grid(row=0, column=3, padx=(8, 0))

    def _build_rule_source_panel(self, parent: ttk.Frame) -> ttk.Frame:
        frame = ttk.LabelFrame(parent, text="Rule Source", style="Assign.TLabelframe", padding=10)
        for row, (value, label) in enumerate(RULE_SOURCE_LABELS.items()):
            ttk.Radiobutton(frame, text=label, value=value, variable=self.rule_source_mode, command=self._on_source_mode_change, style="Assign.TRadiobutton").grid(row=row, column=0, sticky=tk.W)
        path_row = ttk.Frame(frame, style="AssignPanel.TFrame")
        path_row.grid(row=3, column=0, sticky="ew", pady=(8, 0))
        path_row.columnconfigure(0, weight=1)
        self.rule_path_entry = bind_single_paste(ttk.Entry(path_row, textvariable=self.rule_source_path, style="Assign.TEntry"))
        self.rule_path_entry.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        ttk.Button(path_row, text="Browse", command=self._browse_rule_source, style="AssignSoft.TButton").grid(row=0, column=1)
        actions = ttk.Frame(frame, style="AssignPanel.TFrame")
        actions.grid(row=4, column=0, sticky="ew", pady=(8, 0))
        actions.columnconfigure(0, weight=1)
        actions.columnconfigure(1, weight=1)
        ttk.Button(actions, text="Preview Source", command=self._preview_rule_source, style="AssignSoft.TButton").grid(row=0, column=0, sticky="ew", padx=(0, 4))
        ttk.Button(actions, text="Open Keep Note", command=self._open_keep_note, style="AssignSoft.TButton").grid(row=0, column=1, sticky="ew", padx=(4, 0))
        ttk.Button(actions, text="Sync Google Keep", command=self._sync_keep_note, style="AssignSoft.TButton").grid(row=1, column=0, sticky="ew", padx=(0, 4), pady=(6, 0))
        ttk.Button(actions, text="Connect Google", command=self._connect_google, style="AssignSoft.TButton").grid(row=1, column=1, sticky="ew", padx=(4, 0), pady=(6, 0))
        self.link_payouts_check = ttk.Checkbutton(
            frame,
            text="Link Payouts to Same File",
            variable=self.link_payouts_to_rule_source,
            command=self._on_source_mode_change,
            style="Assign.TCheckbutton",
        )
        self.link_payouts_check.grid(row=5, column=0, sticky=tk.W, pady=(8, 0))
        ttk.Label(frame, textvariable=self.preview_status, style="AssignMuted.TLabel", wraplength=420).grid(row=6, column=0, sticky="ew", pady=(8, 0))
        return frame

    def _build_google_status_panel(self, parent: ttk.Frame) -> ttk.Frame:
        frame = ttk.LabelFrame(parent, text="Google Status", style="Assign.TLabelframe", padding=10)
        frame.columnconfigure(0, weight=1)
        self.google_status.set("\n".join(google_status_lines()))
        ttk.Label(frame, textvariable=self.google_status, style="AssignMuted.TLabel", wraplength=520, justify=tk.LEFT).grid(row=0, column=0, columnspan=4, sticky="ew")
        ttk.Button(frame, text="Refresh Status", command=self._refresh_google_status, style="AssignSoft.TButton").grid(row=1, column=0, sticky="ew", pady=(8, 0), padx=(0, 4))
        ttk.Button(frame, text="Reconnect Google", command=self._connect_google, style="AssignSoft.TButton").grid(row=1, column=1, sticky="ew", pady=(8, 0), padx=4)
        ttk.Button(frame, text="Open Token Folder", command=self._open_token_folder, style="AssignSoft.TButton").grid(row=1, column=2, sticky="ew", pady=(8, 0), padx=4)
        ttk.Button(frame, text="Copy Details", command=self._copy_google_details, style="AssignSoft.TButton").grid(row=1, column=3, sticky="ew", pady=(8, 0), padx=(4, 0))
        return frame

    def _build_seller_terms_panel(self, parent: ttk.Frame) -> ttk.Frame:
        frame = ttk.LabelFrame(parent, text="People Rules Health", style="Assign.TLabelframe", padding=10)
        frame.columnconfigure(0, weight=1)
        self._refresh_seller_terms_status()
        ttk.Label(frame, textvariable=self.seller_terms_status, style="AssignMuted.TLabel", wraplength=520, justify=tk.LEFT).grid(row=0, column=0, columnspan=4, sticky="ew")
        ttk.Button(frame, text="Edit People Rules", command=self._open_people_rules, style="AssignPrimary.TButton").grid(row=1, column=0, sticky="ew", pady=(8, 0), padx=(0, 4))
        ttk.Button(frame, text="Refresh", command=self._refresh_seller_terms_status, style="AssignSoft.TButton").grid(row=1, column=1, sticky="ew", pady=(8, 0), padx=4)
        ttk.Button(frame, text="Open Folder", command=self._open_seller_terms_folder, style="AssignSoft.TButton").grid(row=1, column=2, sticky="ew", pady=(8, 0), padx=4)
        ttk.Button(frame, text="Copy Details", command=self._copy_seller_terms_details, style="AssignSoft.TButton").grid(row=1, column=3, sticky="ew", pady=(8, 0), padx=(4, 0))
        return frame

    def _build_payout_source_panel(self, parent: ttk.Frame) -> ttk.Frame:
        frame = ttk.LabelFrame(parent, text="Payout Source", style="Assign.TLabelframe", padding=10)
        for row, (value, label) in enumerate(PAYOUT_SOURCE_LABELS.items()):
            ttk.Radiobutton(frame, text=label, value=value, variable=self.payout_source_mode, command=self._on_source_mode_change, style="Assign.TRadiobutton").grid(row=row, column=0, sticky=tk.W)
        path_row = ttk.Frame(frame, style="AssignPanel.TFrame")
        path_row.grid(row=2, column=0, sticky="ew", pady=(8, 0))
        path_row.columnconfigure(0, weight=1)
        self.payout_path_entry = bind_single_paste(ttk.Entry(path_row, textvariable=self.payout_source_path, style="Assign.TEntry"))
        self.payout_path_entry.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        ttk.Button(path_row, text="Browse", command=self._browse_payout_source, style="AssignSoft.TButton").grid(row=0, column=1)
        return frame

    def _build_payout_panel(self, parent: ttk.Frame) -> ttk.Frame:
        frame = ttk.Frame(parent, style="AssignPanel.TFrame", padding=12)
        header = ttk.Frame(frame, style="AssignPanel.TFrame")
        header.pack(fill=tk.X)
        ttk.Label(header, text="Manual Payout Tiers", style="AssignTitle.TLabel").pack(side=tk.LEFT)
        ttk.Button(header, text="Add Tier", command=self._add_payout_row, style="AssignSoft.TButton").pack(side=tk.RIGHT)
        ttk.Label(frame, text="Used when Payout Source is Manual payout tiers.", style="AssignMuted.TLabel").pack(anchor=tk.W, pady=(3, 8))
        self.payout_frame = ttk.Frame(frame, style="AssignPanel.TFrame")
        self.payout_frame.pack(fill=tk.X)
        return frame

    def _sync_source_visibility(self) -> None:
        rule_linked = self.rule_source_mode.get() != "manual"
        payout_same_file = self.link_payouts_to_rule_source.get() and rule_linked
        manual_rule_payouts = self.rule_source_mode.get() == "manual" and self.payout_source_mode.get() == "manual"
        payout_linked = self.payout_source_mode.get() == "file"
        self.rule_path_entry.configure(state=tk.NORMAL if rule_linked else tk.DISABLED)
        self.link_payouts_check.configure(state=tk.NORMAL if rule_linked else tk.DISABLED)
        if not rule_linked and self.link_payouts_to_rule_source.get():
            self.link_payouts_to_rule_source.set(False)
            payout_same_file = False
        payout_source_state = tk.DISABLED if payout_same_file else (tk.NORMAL if payout_linked else tk.DISABLED)
        self.payout_path_entry.configure(state=payout_source_state)
        if rule_linked:
            self.manual_rule_panel.grid_remove()
        else:
            self.manual_rule_panel.grid(row=0, column=0, columnspan=2, sticky="nsew")
        show_payout_panel = rule_linked and not payout_linked and not payout_same_file
        if not show_payout_panel:
            self.right_panel.grid_remove()
            self.payout_panel.grid_remove()
        else:
            self.right_panel.grid(row=0, column=0, columnspan=2, sticky="nsew", padx=(0, 0))
            self.payout_panel.grid(row=0, column=0, sticky="nsew")

    def _reset_preview_status(self) -> None:
        self.preview_status.set("No source file selected.")

    def _on_source_mode_change(self) -> None:
        self._reset_preview_status()
        self._sync_source_visibility()

    def _update_rules_scrollregion(self, _event=None) -> None:
        if hasattr(self, "rules_canvas"):
            self.rules_canvas.configure(scrollregion=self.rules_canvas.bbox("all"))

    def _resize_rules_canvas_window(self, event) -> None:
        if hasattr(self, "rules_canvas_window"):
            self.rules_canvas.itemconfigure(self.rules_canvas_window, width=event.width)

    def _on_rules_mousewheel(self, event) -> None:
        if hasattr(self, "rules_canvas"):
            self.rules_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _load_config(self) -> list[dict[str, Any]]:
        if not self.config_path.exists():
            return []
        try:
            raw = json.loads(self.config_path.read_text(encoding="utf-8"))
        except Exception:
            return []
        companies = raw.get("companies", raw) if isinstance(raw, dict) else raw
        return [company for company in companies if isinstance(company, dict)]

    def _write_config(self) -> None:
        self.config_path.write_text(json.dumps({"companies": self.companies}, indent=2), encoding="utf-8")

    def _refresh_company_list(self) -> None:
        for child in self.company_list.winfo_children():
            child.destroy()
        self.company_rows = []
        visible_count = 0
        for index, company in enumerate(self.companies):
            if not self._company_matches_filter(company):
                continue
            self._add_company_row(index, company)
            visible_count += 1
        if not visible_count:
            empty_label = tk.Label(
                self.company_list,
                text="No matching companies",
                bg="#1f1f1f",
                fg="#b3b3b3",
                anchor=tk.W,
                padx=10,
                pady=10,
            )
            empty_label.pack(fill=tk.X)

    def _company_matches_filter(self, company: dict[str, Any]) -> bool:
        needle = self.company_filter_text.get().strip().lower()
        if needle and needle not in str(company.get("name") or "Untitled").lower():
            return False
        state = self.company_filter_state.get()
        active = company.get("active") is not False
        if state == "active" and not active:
            return False
        if state == "inactive" and active:
            return False
        return True

    def _add_company_row(self, index: int, company: dict[str, Any]) -> None:
        active = company.get("active") is not False
        selected = index == self.selected_index
        bg = "#242424" if selected else "#1f1f1f"
        row = tk.Frame(self.company_list, bg=bg, padx=5, pady=4)
        row.pack(fill=tk.X, padx=4, pady=(4 if not self.company_rows else 0, 0))
        active_button = tk.Button(
            row,
            text="Active" if active else "Inactive",
            command=lambda row_index=index: self._toggle_company_active(row_index),
            bg="#1ed760" if active else "#3a3a3a",
            fg="#000000" if active else "#d9d9d9",
            activebackground="#1fdf64" if active else "#4a4a4a",
            activeforeground="#000000" if active else "#ffffff",
            relief=tk.FLAT,
            borderwidth=0,
            font=("Segoe UI Semibold", 8),
            width=8,
            padx=6,
            pady=4,
        )
        active_button.pack(side=tk.LEFT, padx=(0, 6))
        name_button = tk.Button(
            row,
            text=str(company.get("name") or "Untitled"),
            command=lambda row_index=index: self._select_company(row_index),
            anchor=tk.W,
            bg=bg,
            fg="#ffffff" if selected else "#d9d9d9",
            activebackground="#242424",
            activeforeground="#ffffff",
            relief=tk.FLAT,
            borderwidth=0,
            font=("Segoe UI Semibold" if selected else "Segoe UI", 10),
            padx=4,
            pady=4,
        )
        name_button.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.company_rows.append(row)

    def _toggle_company_active(self, index: int) -> None:
        if not (0 <= index < len(self.companies)):
            return
        company = self.companies[index]
        company["active"] = not (company.get("active") is not False)
        self._write_config()
        self._refresh_company_list()
        self._refresh_seller_terms_status()
        state = "active" if company.get("active") is not False else "inactive"
        self.status.set(f"{company.get('name') or 'Company'} is now {state}. Save & Reload to apply in Assignment.")

    def _select_company(self, index: int) -> None:
        self.selected_index = index
        self._refresh_company_list()
        company = self.companies[index]
        self.company_name.set(str(company.get("name") or ""))
        self.value_source.set(normalize_value_source(company.get("value_source") or company.get("valueSource")))
        self.reset_weekday.set(normalize_company_reset_weekday(company.get("reset_weekday") or company.get("resetWeekday")))
        try:
            self.reset_time.set(normalize_company_reset_time(company.get("reset_time") or company.get("resetTime")))
        except ValueError:
            self.reset_time.set(DEFAULT_COMPANY_RESET_TIME)
        self.rule_source_mode.set(str(company.get("rules_source_kind") or source_kind_for_path(company.get("rules"))))
        self.rule_source_path.set(display_source_path(company.get("rules")))
        self.rule_materialized_source = company.get("rules") if isinstance(company.get("rules"), dict) else None
        linked_payouts = is_same_file_payout_source(company.get("payout"))
        self.link_payouts_to_rule_source.set(linked_payouts)
        payout_kind = str(company.get("payout_source_kind") or "")
        if linked_payouts or payout_kind == "same_file":
            self.payout_source_mode.set("file")
            self.link_payouts_to_rule_source.set(True)
        else:
            self.payout_source_mode.set(str(payout_kind or ("file" if company.get("payout") and not is_generated_payout_path(company.get("payout")) else "manual")))
        self.payout_source_path.set(display_source_path(company.get("payout")))
        self.payout_materialized_source = company.get("payout") if isinstance(company.get("payout"), dict) else None
        rules_payload = self._load_json_source(company.get("rules") or company.get("rules_source") or company.get("rulesSource"))
        payout_payload = self._load_json_source(company.get("payout") or company.get("payout_source") or company.get("payoutSource"))
        self.manual_min_year.set(str((rules_payload or {}).get("minYear") or (rules_payload or {}).get("min_year") or ""))
        self.manual_max_year.set(str((rules_payload or {}).get("maxYear") or (rules_payload or {}).get("max_year") or ""))
        self._set_rule_rows(rules_payload.get("rules") if isinstance(rules_payload, dict) else [])
        self._set_payout_rows(payout_payload.get("tiers") if isinstance(payout_payload, dict) else [])
        if self.rule_source_mode.get() == "manual" and self.payout_source_mode.get() == "manual":
            self._apply_payout_tiers_to_rule_rows(payout_payload.get("tiers") if isinstance(payout_payload, dict) else [])
        self._sync_source_visibility()
        self._reset_preview_status()
        self.status.set(f"Editing {company.get('name') or 'company'}.")

    def _load_json_source(self, source: Any) -> dict[str, Any]:
        try:
            text = read_source_text(source, self.config_path.parent)
            payload = json.loads(text) if text.strip() else {}
            return payload if isinstance(payload, dict) else {}
        except Exception:
            return {}

    def _new_company(self) -> None:
        self.selected_index = None
        self._refresh_company_list()
        self.company_name.set("")
        self.value_source.set("comps")
        self.reset_weekday.set(DEFAULT_COMPANY_RESET_WEEKDAY)
        self.reset_time.set(DEFAULT_COMPANY_RESET_TIME)
        self.rule_source_mode.set("manual")
        self.rule_source_path.set("")
        self.link_payouts_to_rule_source.set(False)
        self.payout_source_mode.set("manual")
        self.payout_source_path.set("")
        self.rule_materialized_source = None
        self.payout_materialized_source = None
        self.manual_min_year.set("")
        self.manual_max_year.set("")
        self._reset_preview_status()
        self._set_rule_rows([])
        self._set_payout_rows([])
        self._sync_source_visibility()
        self.status.set("New company ready.")

    def _delete_company(self) -> None:
        if self.selected_index is None:
            return
        name = self.companies[self.selected_index].get("name") or "this company"
        if not messagebox.askyesno("Delete company", f"Delete {name}?"):
            return
        del self.companies[self.selected_index]
        self.selected_index = None
        self._write_config()
        self._refresh_company_list()
        self._new_company()
        self.status.set(f"Deleted {name}.")

    def _browse_rule_source(self) -> None:
        path = filedialog.askopenfilename(
            parent=self,
            title="Choose local rules file",
            initialdir=str(self.pipeline_root if self.pipeline_root.exists() else Path.home()),
            filetypes=[
                ("Rule files", "*.txt *.md *.json *.csv *.xlsx *.xlsm *.gsheet"),
                ("All files", "*.*"),
            ],
        )
        if path:
            self.rule_source_path.set(path)
            self.rule_materialized_source = None
            self._preview_rule_source()

    def _browse_payout_source(self) -> None:
        path = filedialog.askopenfilename(
            parent=self,
            title="Choose local payout file",
            initialdir=str(self.pipeline_root if self.pipeline_root.exists() else Path.home()),
            filetypes=[
                ("Payout files", "*.txt *.md *.json *.csv *.xlsx *.xlsm *.gsheet"),
                ("All files", "*.*"),
            ],
        )
        if path:
            self.payout_source_path.set(path)
            self.payout_materialized_source = None

    def _preview_rule_source(self) -> None:
        path = self.rule_source_path.get().strip()
        if not path:
            self.preview_status.set("No source file selected.")
            return
        try:
            source = self._materialize_source_if_needed(path)
            if isinstance(source, dict):
                self.rule_materialized_source = source
                self.rule_source_path.set(str(source.get("path") or ""))
            else:
                self.rule_materialized_source = None
                self.rule_source_path.set(source)
            text = read_source_text(source, self.config_path.parent, interactive_google=True)
        except Exception as error:
            self.preview_status.set(f"Could not read source: {error}")
            return
        lines = [line for line in text.splitlines() if line.strip()]
        if not lines:
            self.preview_status.set(
                f"No rows found in {Path(path).name}. If this is a native Google Sheet shortcut, choose a synced/exported .xlsx or .csv file instead."
            )
            return
        self.preview_status.set(f"Read {len(lines)} non-empty line(s) from {Path(path).name}.")
        self._refresh_google_status()

    def _set_rule_rows(self, rules: Any) -> None:
        for child in self.rules_frame.winfo_children():
            child.destroy()
        self.rule_rows = []
        if not isinstance(rules, list) or not rules:
            self._add_rule_row()
            return
        for rule in rules:
            self._add_rule_row(rule if isinstance(rule, dict) else None)

    def _add_rule_row(self, data: dict[str, Any] | None = None) -> None:
        index = len(self.rule_rows)
        frame = ttk.LabelFrame(self.rules_frame, text=f"Rule {index + 1}", style="Assign.TLabelframe", padding=10)
        frame.pack(fill=tk.X, pady=(0, 10))
        frame.columnconfigure(0, weight=1)

        card_header = ttk.Frame(frame, style="AssignPanel.TFrame")
        card_header.grid(row=0, column=0, sticky="ew")
        ttk.Label(card_header, text="Rule Details", style="Assign.TLabel").pack(side=tk.LEFT)
        ttk.Button(card_header, text="Remove Rule", command=lambda: self._remove_rule_row(frame), style="AssignSoft.TButton").pack(side=tk.RIGHT)

        sports = set(data.get("sports") or split_values(data.get("sport")) if data else [])
        sport_vars = {sport: tk.BooleanVar(value=sport in sports) for sport in SPORT_OPTIONS}
        ttk.Label(frame, text="Category", style="Assign.TLabel").grid(row=1, column=0, sticky=tk.W, pady=(8, 0))
        sport_frame = ttk.Frame(frame, style="AssignPanel.TFrame")
        sport_frame.grid(row=2, column=0, sticky="ew", pady=(6, 12))
        sport_frame.columnconfigure(0, weight=1)
        sport_frame.columnconfigure(1, weight=1)
        for option_index, sport in enumerate(SPORT_OPTIONS):
            row = option_index // 2
            col = option_index % 2
            ttk.Checkbutton(
                sport_frame,
                text=category_label(sport),
                variable=sport_vars[sport],
                style="Assign.TCheckbutton",
            ).grid(row=row, column=col, sticky=tk.W, padx=(0, 18), pady=2)

        price_frame = ttk.Frame(frame, style="AssignPanel.TFrame")
        price_frame.grid(row=3, column=0, sticky="ew")
        price_frame.columnconfigure(1, weight=1)
        price_frame.columnconfigure(3, weight=1)
        price_frame.columnconfigure(5, weight=1)
        price_ranges = data.get("priceRanges") if data else None
        first_range = price_ranges[0] if isinstance(price_ranges, list) and price_ranges else {}
        min_var = tk.StringVar(value=str(first_range.get("min") or ""))
        max_var = tk.StringVar(value=str(first_range.get("max") or ""))
        payout_value = (data or {}).get("payout") or (data or {}).get("rate") or first_range.get("payout") or first_range.get("rate") or ""
        payout_var = tk.StringVar(value=str(payout_value))
        ttk.Label(price_frame, text="Price Range & Payout Percentage", style="Assign.TLabel").grid(row=0, column=0, columnspan=6, sticky=tk.W, pady=(0, 6))
        ttk.Label(price_frame, text="Min", style="Assign.TLabel").grid(row=1, column=0, sticky=tk.W, padx=(0, 6))
        bind_single_paste(ttk.Entry(price_frame, textvariable=min_var, width=12, style="Assign.TEntry")).grid(row=1, column=1, sticky=tk.W)
        ttk.Label(price_frame, text="Max", style="Assign.TLabel").grid(row=1, column=2, sticky=tk.W, padx=(14, 6))
        bind_single_paste(ttk.Entry(price_frame, textvariable=max_var, width=12, style="Assign.TEntry")).grid(row=1, column=3, sticky=tk.W)
        ttk.Label(price_frame, text="Payout Percentage", style="Assign.TLabel").grid(row=1, column=4, sticky=tk.W, padx=(14, 6))
        bind_single_paste(ttk.Entry(price_frame, textvariable=payout_var, width=12, style="Assign.TEntry")).grid(row=1, column=5, sticky=tk.W)

        grades_frame = ttk.Frame(frame, style="AssignPanel.TFrame")
        grades_frame.grid(row=4, column=0, sticky="ew", pady=(12, 0))
        grades_frame.columnconfigure(0, minsize=62)
        grades_frame.columnconfigure(1, minsize=92)
        grades_frame.columnconfigure(2, minsize=86)
        grades_frame.columnconfigure(3, minsize=86)
        ttk.Label(grades_frame, text="Grade Ranges", style="Assign.TLabel").grid(row=0, column=0, columnspan=4, sticky=tk.W, pady=(0, 6))
        ttk.Label(grades_frame, text="Company", style="AssignMuted.TLabel").grid(row=1, column=0, sticky=tk.W)
        ttk.Label(grades_frame, text="Allowed", style="AssignMuted.TLabel").grid(row=1, column=1, sticky=tk.W)
        ttk.Label(grades_frame, text="Min", style="AssignMuted.TLabel").grid(row=1, column=2, sticky=tk.W)
        ttk.Label(grades_frame, text="Max", style="AssignMuted.TLabel").grid(row=1, column=3, sticky=tk.W)
        grade_payload = data.get("grades") if data else {}
        grade_vars: dict[str, dict[str, Any]] = {}
        for grade_index, company in enumerate(GRADE_COMPANIES, start=2):
            payload = grade_payload.get(company) if isinstance(grade_payload, dict) else {}
            allowed = tk.BooleanVar(value=(payload or {}).get("allowed") is not False)
            min_grade = tk.StringVar(value=str((payload or {}).get("min") or ""))
            max_grade = tk.StringVar(value=str((payload or {}).get("max") or ""))
            ttk.Label(grades_frame, text=company.upper(), style="Assign.TLabel").grid(row=grade_index, column=0, sticky=tk.W, pady=3)
            ttk.Checkbutton(grades_frame, text="", variable=allowed, style="Assign.TCheckbutton").grid(row=grade_index, column=1, sticky=tk.W, pady=3)
            bind_single_paste(ttk.Entry(grades_frame, textvariable=min_grade, width=8, style="Assign.TEntry")).grid(row=grade_index, column=2, sticky=tk.W, pady=3)
            bind_single_paste(ttk.Entry(grades_frame, textvariable=max_grade, width=8, style="Assign.TEntry")).grid(row=grade_index, column=3, sticky=tk.W, pady=3)
            grade_vars[company] = {"allowed": allowed, "min": min_grade, "max": max_grade}

        self.rule_rows.append({"frame": frame, "sports": sport_vars, "min": min_var, "max": max_var, "payout": payout_var, "grades": grade_vars})
        self._renumber_rule_rows()

    def _remove_rule_row(self, frame: ttk.Frame) -> None:
        self.rule_rows = [row for row in self.rule_rows if row["frame"] is not frame]
        frame.destroy()
        if not self.rule_rows:
            self._add_rule_row()
        self._renumber_rule_rows()

    def _renumber_rule_rows(self) -> None:
        for index, row in enumerate(self.rule_rows, start=1):
            row["frame"].configure(text=f"Rule {index}")

    def _set_payout_rows(self, tiers: Any) -> None:
        for child in self.payout_frame.winfo_children():
            child.destroy()
        self.payout_rows = []
        if not isinstance(tiers, list) or not tiers:
            self._add_payout_row({"min": "", "max": "", "rate": ""})
            return
        for tier in tiers:
            self._add_payout_row(tier if isinstance(tier, dict) else None)

    def _apply_payout_tiers_to_rule_rows(self, tiers: Any) -> None:
        if not isinstance(tiers, list):
            return
        unused_tiers = [tier for tier in tiers if isinstance(tier, dict)]
        for row in self.rule_rows:
            if row["payout"].get().strip():
                continue
            row_min = row["min"].get().strip()
            row_max = row["max"].get().strip()
            match = next(
                (
                    tier
                    for tier in unused_tiers
                    if str(tier.get("min") or "").strip() == row_min and str(tier.get("max") or "").strip() == row_max
                ),
                None,
            )
            if match is None and unused_tiers:
                match = unused_tiers[0]
            if match is None:
                continue
            row["payout"].set(str(match.get("rate") or ""))
            unused_tiers.remove(match)

    def _add_payout_row(self, data: dict[str, Any] | None = None) -> None:
        row_index = len(self.payout_rows)
        min_var = tk.StringVar(value=str((data or {}).get("min") or ""))
        max_var = tk.StringVar(value=str((data or {}).get("max") or ""))
        rate_var = tk.StringVar(value=str((data or {}).get("rate") or ""))
        ttk.Label(self.payout_frame, text="Min", style="Assign.TLabel").grid(row=row_index, column=0, sticky=tk.W, padx=(0, 6), pady=4)
        bind_single_paste(ttk.Entry(self.payout_frame, textvariable=min_var, width=9, style="Assign.TEntry")).grid(row=row_index, column=1, sticky=tk.W, pady=4)
        ttk.Label(self.payout_frame, text="Max", style="Assign.TLabel").grid(row=row_index, column=2, sticky=tk.W, padx=(10, 6), pady=4)
        bind_single_paste(ttk.Entry(self.payout_frame, textvariable=max_var, width=9, style="Assign.TEntry")).grid(row=row_index, column=3, sticky=tk.W, pady=4)
        ttk.Label(self.payout_frame, text="Rate", style="Assign.TLabel").grid(row=row_index, column=4, sticky=tk.W, padx=(10, 6), pady=4)
        bind_single_paste(ttk.Entry(self.payout_frame, textvariable=rate_var, width=9, style="Assign.TEntry")).grid(row=row_index, column=5, sticky=tk.W, pady=4)
        self.payout_rows.append({"min": min_var, "max": max_var, "rate": rate_var})

    def _save_company(self) -> bool:
        name = self.company_name.get().strip()
        if not name:
            messagebox.showinfo("Company name", "Name the company before saving.")
            return False
        reset_weekday = normalize_company_reset_weekday(self.reset_weekday.get())
        try:
            reset_time = normalize_company_reset_time(self.reset_time.get())
        except ValueError as error:
            messagebox.showinfo("Company sheet reset", str(error))
            return False

        rule_source = self._save_or_link_rules(name)
        if not rule_source:
            return False
        payout_source = self._save_or_link_payout(name, rule_source)
        if not payout_source:
            return False

        company = {
            "name": name,
            "active": self.companies[self.selected_index].get("active", True) if self.selected_index is not None else True,
            "value_source": normalize_value_source(self.value_source.get()),
            "reset_weekday": reset_weekday,
            "reset_time": reset_time,
            "rules": rule_source,
            "rules_source_kind": self.rule_source_mode.get(),
            "payout": payout_source,
            "payout_source_kind": "same_file" if self.link_payouts_to_rule_source.get() else self.payout_source_mode.get(),
        }
        if self.selected_index is None:
            self.companies.append(company)
            self.selected_index = len(self.companies) - 1
        else:
            self.companies[self.selected_index] = company
        self._write_config()
        folder_error = self._ensure_company_sheet_folder(name)
        self._refresh_company_list()
        self._select_company(self.selected_index)
        self._refresh_seller_terms_status()
        if folder_error:
            self.status.set(f"Saved {name}. Company sheet folder failed: {folder_error}")
        else:
            self.status.set(f"Saved {name}. Company sheet folder ready.")
        return True

    def _ensure_company_sheet_folder(self, name: str) -> str:
        try:
            folder_name = safe_filename(name)
            self.company_sheets_dir.mkdir(parents=True, exist_ok=True)
            (self.company_sheets_dir / folder_name).mkdir(parents=True, exist_ok=True)
            return ""
        except Exception as error:
            return str(error)

    def _save_or_link_rules(self, name: str) -> str | dict[str, Any]:
        mode = self.rule_source_mode.get()
        if mode != "manual":
            path = self.rule_source_path.get().strip()
            if not path:
                messagebox.showinfo("Rule source", "Choose the local rules file before saving.")
                return ""
            try:
                return self._materialize_source_if_needed(path)
            except Exception as error:
                messagebox.showerror("Rule source", f"Could not prepare rule source: {error}")
                return ""
        self.rules_dir.mkdir(parents=True, exist_ok=True)
        rules_path = self.rules_dir / f"{safe_stem(name)}-rules.json"
        try:
            rules_payload = self._rules_payload()
        except ValueError as error:
            messagebox.showinfo("Manual rules", str(error))
            return ""
        rules_path.write_text(json.dumps(rules_payload, indent=2), encoding="utf-8")
        return str(rules_path)

    def _save_or_link_payout(self, name: str, rule_source: str | dict[str, Any]) -> str | dict[str, Any]:
        if self.link_payouts_to_rule_source.get() and self.rule_source_mode.get() != "manual":
            return self._same_file_payout_source(rule_source)
        mode = self.payout_source_mode.get()
        if mode == "file":
            path = self.payout_source_path.get().strip()
            if not path:
                messagebox.showinfo("Payout source", "Choose the local payout file before saving.")
                return ""
            try:
                return self._materialize_source_if_needed(path, is_payout=True)
            except Exception as error:
                messagebox.showerror("Payout source", f"Could not prepare payout source: {error}")
                return ""
        self.rules_dir.mkdir(parents=True, exist_ok=True)
        payout_path = self.rules_dir / f"{safe_stem(name)}-payout.json"
        payout_path.write_text(json.dumps(self._payout_payload(), indent=2), encoding="utf-8")
        return str(payout_path)

    def _same_file_payout_source(self, rule_source: str | dict[str, Any]) -> dict[str, Any]:
        if isinstance(rule_source, dict):
            source = dict(rule_source)
        else:
            source = {"path": rule_source}
        source["sheet_name"] = "Payouts"
        source["linked_payouts_to_rule_source"] = True
        return source

    def _materialize_source_if_needed(self, path_text: str, is_payout: bool = False) -> str | dict[str, Any]:
        normalized = normalize_source_value(path_text)
        cached = self.payout_materialized_source if is_payout else self.rule_materialized_source
        if isinstance(cached, dict) and (
            normalize_source_value(cached.get("path")) == normalized
            or normalize_source_value(cached.get("url")) == normalized
        ):
            return cached
        if is_google_keep_url(normalized):
            source_name = keep_note_display_name(normalized)
            source = {
                "kind": "google_keep",
                "url": normalized,
                "path": str(keep_note_cache_path(normalized, self.pipeline_root, source_name)),
                "name": source_name,
                "refresh_on_load": True,
            }
            if is_payout:
                self.payout_materialized_source = source
            else:
                self.rule_materialized_source = source
            return source
        if is_google_sheet_url(normalized):
            source_name = google_sheet_display_name(normalized)
            source = {
                "kind": "google_sheet",
                "url": normalized,
                "path": str(self.rules_dir / "SHEET EXPORTS" / f"{safe_filename(source_name)}.xlsx"),
                "name": source_name,
                "refresh_on_load": True,
            }
            if is_payout:
                self.payout_materialized_source = source
            else:
                self.rule_materialized_source = source
            return source
        path = Path(normalized).expanduser()
        if path.suffix.lower() != ".gsheet":
            return normalized
        export_dir = self.rules_dir / "SHEET EXPORTS"
        source_url = ""
        source_name = path.stem or "google-sheet"
        try:
            shortcut = load_gsheet_shortcut(path)
            source_url = gsheet_shortcut_url(shortcut)
            source_name = str(shortcut.get("name") or source_name)
        except Exception as error:
            url = simpledialog.askstring(
                "Google Sheet URL",
                (
                    "Google Drive did not expose this .gsheet shortcut as a readable local file.\n\n"
                    "Paste the Google Sheet URL and L.U.C.A.S will read it with the authenticated Google Sheets connection."
                ),
                parent=self,
            )
            if not url:
                raise error
            source_url = url.strip()
        if not source_url:
            raise ValueError("This .gsheet shortcut does not contain a Google Sheet URL or document id.")
        exported = export_dir / f"{safe_filename(source_name)}.xlsx"
        source = {
            "kind": "google_sheet",
            "url": source_url,
            "path": str(exported),
            "name": source_name,
            "refresh_on_load": True,
        }
        if is_payout:
            self.payout_materialized_source = source
        else:
            self.rule_materialized_source = source
        return source

    def _connect_google(self) -> None:
        try:
            authorize_google_sheets(interactive=True)
        except Exception as error:
            details = self._google_diagnostic_payload("Connect Google failed", str(error))
            self.clipboard_clear()
            self.clipboard_append(diagnostic_json(details))
            messagebox.showerror("Connect Google", f"{error}\n\nDiagnostic details were copied to the clipboard.")
            self._refresh_google_status()
            return
        self.preview_status.set("Google Sheets connected. Preview the source to read the live sheet.")
        self._refresh_google_status()

    def _refresh_google_status(self) -> None:
        sheet_status = self._selected_sheet_status()
        keep_status = self._selected_keep_status()
        self.google_status.set("\n".join(google_status_lines(sheet_status=sheet_status, keep_status=keep_status)))

    def _selected_sheet_status(self) -> str:
        source = self.rule_materialized_source if isinstance(self.rule_materialized_source, dict) else self.rule_source_path.get().strip()
        raw = normalize_source_value(source.get("url") if isinstance(source, dict) else source)
        if not is_google_sheet_url(raw):
            return "no Google Sheet source selected"
        try:
            text = read_source_text(source, self.config_path.parent, interactive_google=False)
            return f"yes ({len([line for line in text.splitlines() if line.strip()])} non-empty line(s))"
        except Exception as error:
            return f"no ({error})"

    def _selected_keep_status(self) -> str:
        source = self.rule_materialized_source if isinstance(self.rule_materialized_source, dict) else None
        raw = normalize_source_value((source or {}).get("url") or self.rule_source_path.get().strip())
        if not is_google_keep_url(raw):
            return "no Google Keep source selected"
        path = Path(str((source or {}).get("path") or keep_note_cache_path(raw, self.pipeline_root)))
        if not path.exists():
            return f"not synced ({path})"
        try:
            mtime = path.stat().st_mtime
        except OSError:
            return f"cache unreadable ({path})"
        return f"{Path(path).name} at {__import__('datetime').datetime.fromtimestamp(mtime).strftime('%Y-%m-%d %H:%M:%S')}"

    def _open_token_folder(self) -> None:
        folder = TOKEN_PATH.parent
        try:
            if os.name == "nt":
                os.startfile(str(folder))  # type: ignore[attr-defined]
            else:
                webbrowser.open(folder.as_uri())
        except Exception as error:
            messagebox.showerror("Open Token Folder", str(error))

    def _copy_google_details(self) -> None:
        details = self._google_diagnostic_payload("Google status", "")
        self.clipboard_clear()
        self.clipboard_append(diagnostic_json(details))
        self.status.set("Copied Google diagnostic details.")

    def _seller_terms_path(self) -> Path:
        return self.rules_dir / "seller_terms.csv"

    def _refresh_seller_terms_status(self) -> None:
        self.seller_terms_status.set("\n".join(seller_terms_health_lines(self._seller_terms_path(), self.companies)))

    def _open_people_rules(self) -> None:
        open_people_rules_dialog(self, self.pipeline_root, self.companies, self._refresh_seller_terms_status)

    def _open_seller_terms_folder(self) -> None:
        folder = self.rules_dir
        try:
            folder.mkdir(parents=True, exist_ok=True)
            if os.name == "nt":
                os.startfile(str(folder))  # type: ignore[attr-defined]
            else:
                webbrowser.open(folder.as_uri())
        except Exception as error:
            messagebox.showerror("Open Terms Folder", str(error))

    def _copy_seller_terms_details(self) -> None:
        self._refresh_seller_terms_status()
        details = {
            "title": "People Rules Health",
            "seller_terms_path": str(self._seller_terms_path()),
            "seller_terms_status": self.seller_terms_status.get(),
            "assignment_companies": [
                {
                    "name": str(company.get("name") or ""),
                    "active": company.get("active") is not False,
                    "value_source": company.get("value_source") or company.get("valueSource") or "",
                }
                for company in self.companies
            ],
        }
        self.clipboard_clear()
        self.clipboard_append(diagnostic_json(details))
        self.status.set("Copied seller terms diagnostic details.")

    def _google_diagnostic_payload(self, title: str, error: str) -> dict[str, Any]:
        return {
            "title": title,
            "error": error,
            "oauth": dict(LAST_OAUTH_DIAGNOSTICS),
            "google_status": self.google_status.get(),
            "rule_source": self.rule_materialized_source or self.rule_source_path.get(),
            "payout_source": self.payout_materialized_source or self.payout_source_path.get(),
        }

    def _open_keep_note(self) -> None:
        path = self.rule_source_path.get().strip()
        source = self.rule_materialized_source if isinstance(self.rule_materialized_source, dict) else None
        url = str((source or {}).get("url") or path).strip()
        if not is_google_keep_url(url):
            messagebox.showinfo("Open Keep Note", "Paste or select a Google Keep note URL first.")
            return
        webbrowser.open(url)
        self.preview_status.set("Opened Google Keep. Keep the note open; L.U.C.A.S will use the latest synced cache.")

    def _current_keep_sources(self) -> list[dict[str, Any]]:
        sources: list[dict[str, Any]] = []
        seen: set[str] = set()

        def add_source(source: object, is_payout: bool = False) -> None:
            prepared: dict[str, Any] | None = None
            if isinstance(source, dict) and str(source.get("kind") or "").strip() == "google_keep":
                prepared = dict(source)
            else:
                raw = normalize_source_value(source.get("url") if isinstance(source, dict) else source)
                if is_google_keep_url(raw):
                    prepared = dict(self._materialize_source_if_needed(raw, is_payout=is_payout))  # type: ignore[arg-type]
            if not prepared:
                return
            url = str(prepared.get("url") or "").strip()
            if not url or url in seen:
                return
            seen.add(url)
            sources.append(prepared)

        rule_source = self.rule_materialized_source if isinstance(self.rule_materialized_source, dict) else self.rule_source_path.get().strip()
        add_source(rule_source)
        if not self.link_payouts_to_rule_source.get():
            payout_source = self.payout_materialized_source if isinstance(self.payout_materialized_source, dict) else self.payout_source_path.get().strip()
            add_source(payout_source, is_payout=True)
        return sources

    def _sync_keep_note(self) -> None:
        sources = self._current_keep_sources()
        if not sources:
            messagebox.showinfo("Sync Keep", "Select a Google Keep note URL for this company first.")
            return
        parent_state = getattr(getattr(self, "master", None), "state", None)
        if parent_state is not None and hasattr(parent_state, "register_keep_note_sources"):
            try:
                parent_state.register_keep_note_sources(sources)
            except Exception:
                pass
        opened = 0
        for source in sources:
            url = str(source.get("url") or "").strip()
            if url and self._open_google_keep_source_url(url):
                opened += 1
        if opened:
            note_word = "note" if opened == 1 else "notes"
            self.preview_status.set(f"Opened {opened} Google Keep {note_word} for this company. The extension will sync after each note loads.")
            self._refresh_google_status()
            return
        messagebox.showinfo("Sync Keep", "L.U.C.A.S could not open this company's Google Keep note.")

    def _open_google_keep_source_url(self, url: str) -> bool:
        if sys.platform == "win32":
            chrome_candidates = [
                Path(os.environ.get("ProgramFiles", "")) / "Google" / "Chrome" / "Application" / "chrome.exe",
                Path(os.environ.get("ProgramFiles(x86)", "")) / "Google" / "Chrome" / "Application" / "chrome.exe",
                Path(os.environ.get("LocalAppData", "")) / "Google" / "Chrome" / "Application" / "chrome.exe",
            ]
            for chrome_path in chrome_candidates:
                try:
                    if chrome_path.exists():
                        subprocess.Popen([str(chrome_path), url], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                        return True
                except Exception:
                    continue
        if sys.platform == "darwin":
            try:
                result = subprocess.run(
                    ["open", "-a", "Google Chrome", url],
                    check=False,
                    capture_output=True,
                    text=True,
                    timeout=10,
                )
                if result.returncode == 0:
                    return True
            except Exception:
                pass
        return bool(webbrowser.open(url))

    def _save_and_reload(self) -> None:
        if self._save_company():
            self.on_saved()
            self.status.set("Saved and reloaded Assignment rules.")

    def _rules_payload(self) -> dict[str, Any]:
        min_year = self._validated_year(self.manual_min_year.get(), "Min Year")
        max_year = self._validated_year(self.manual_max_year.get(), "Max Year")
        if min_year and max_year and min_year > max_year:
            raise ValueError("Min Year cannot be later than Max Year.")
        return {
            "rules": [self._rule_payload(row) for row in self.rule_rows],
            "blocks": [],
            "minYear": str(min_year) if min_year else "",
            "maxYear": str(max_year) if max_year else "",
        }

    def _validated_year(self, value: str, label: str) -> int | None:
        text = str(value or "").strip()
        if not text:
            return None
        if not re.fullmatch(r"(?:19|20)\d{2}", text):
            raise ValueError(f"{label} must be a four-digit year, such as 1990.")
        year = int(text)
        if year < 1900 or year > 2099:
            raise ValueError(f"{label} must be between 1900 and 2099.")
        return year

    def _rule_payload(self, row: dict[str, Any]) -> dict[str, Any]:
        return {
            "sports": [sport for sport, var in row["sports"].items() if var.get()],
            "priceRanges": [{"min": row["min"].get().strip(), "max": row["max"].get().strip()}],
            "payout": row["payout"].get().strip(),
            "grades": {
                company: {
                    "allowed": values["allowed"].get(),
                    "min": values["min"].get().strip(),
                    "max": values["max"].get().strip(),
                }
                for company, values in row["grades"].items()
            },
        }

    def _payout_payload(self) -> dict[str, Any]:
        manual_rule_tiers = self._manual_rule_payout_tiers()
        if manual_rule_tiers:
            return {"tiers": manual_rule_tiers}
        return {
            "tiers": [
                {"min": row["min"].get().strip(), "max": row["max"].get().strip(), "rate": row["rate"].get().strip()}
                for row in self.payout_rows
                if row["rate"].get().strip()
            ]
        }

    def _manual_rule_payout_tiers(self) -> list[dict[str, str]]:
        if self.rule_source_mode.get() != "manual" or self.payout_source_mode.get() != "manual":
            return []
        return [
            {"min": row["min"].get().strip(), "max": row["max"].get().strip(), "rate": row["payout"].get().strip()}
            for row in self.rule_rows
            if row["payout"].get().strip()
        ]


def open_assignment_rules_dialog(parent: tk.Tk, pipeline_root: Path, on_saved: Callable[[], None]) -> None:
    dialog = AssignmentRulesDialog(parent, pipeline_root, on_saved)
    dialog.focus_set()
    dialog.grab_set()


class PeopleRulesDialog(tk.Toplevel):
    def __init__(self, parent: tk.Misc, pipeline_root: Path, companies: list[dict[str, Any]], on_saved: Callable[[], None] | None = None) -> None:
        super().__init__(parent)
        self.title("People Rules")
        self.geometry("980x620")
        self.minsize(760, 420)
        self.transient(parent)
        self.configure(bg="#121212")
        self.pipeline_root = Path(pipeline_root)
        self.rules_dir = self.pipeline_root / "ASSIGNMENT RULES"
        self.seller_terms_path = self.rules_dir / "seller_terms.csv"
        self.companies = companies
        self.on_saved = on_saved
        self.row_vars: list[dict[str, tk.StringVar]] = []
        self.status = tk.StringVar(value="People rules control Network Mode seller payouts.")
        AssignmentRulesDialog._configure_styles(self)  # type: ignore[misc]
        self._build_ui()
        self._load_rows()

    def _active_sheet_types(self) -> list[str]:
        names = [
            str(company.get("name") or "").strip()
            for company in self.companies
            if str(company.get("name") or "").strip() and company.get("active") is not False
        ]
        return sorted(set(names), key=str.lower)

    def _build_ui(self) -> None:
        shell = build_scrollable_dialog_body(self, "Assign.TFrame", padding=16)
        shell.columnconfigure(0, weight=1)
        ttk.Label(shell, text="People Rules", style="AssignHeader.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 8))
        ttk.Label(
            shell,
            text="Add seller payout terms here. Sheet Type must match an active company rule.",
            style="AssignBgMuted.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(0, 12))
        table = ttk.Frame(shell, style="AssignPanel.TFrame", padding=12)
        table.grid(row=2, column=0, sticky="nsew")
        shell.rowconfigure(2, weight=1)
        self.rows_frame = table
        headings = tuple(SELLER_TERMS_FIELD_LABELS.get(field, field) for field in SELLER_TERMS_FIELDS) + ("",)
        widths = (26, 24, 14, 14, 10)
        for column, (heading, width) in enumerate(zip(headings, widths)):
            ttk.Label(table, text=heading, style="AssignTitle.TLabel").grid(row=0, column=column, sticky="w", padx=(0, 8), pady=(0, 8))
            table.columnconfigure(column, weight=1 if column in {0, 1} else 0, minsize=width * 8)
        actions = ttk.Frame(shell, style="Assign.TFrame")
        actions.grid(row=3, column=0, sticky="ew", pady=(12, 0))
        actions.columnconfigure(0, weight=1)
        ttk.Label(actions, textvariable=self.status, style="AssignBgMuted.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Button(actions, text="Add Rule", command=self._add_row, style="AssignSoft.TButton").grid(row=0, column=1, padx=(8, 0))
        ttk.Button(actions, text="Save", command=self._save, style="AssignPrimary.TButton").grid(row=0, column=2, padx=(8, 0))
        ttk.Button(actions, text="Close", command=self.destroy, style="AssignSoft.TButton").grid(row=0, column=3, padx=(8, 0))

    def _load_rows(self) -> None:
        try:
            rows = read_seller_terms_rows(self.seller_terms_path)
        except Exception as error:
            rows = []
            self.status.set(f"Could not read People Rules: {error}")
        for row in rows:
            self._add_row(row)
        if not rows:
            self._add_row()

    def _add_row(self, row: dict[str, str] | None = None) -> None:
        values = {field: str((row or {}).get(field) or "") for field in SELLER_TERMS_FIELDS}
        for field in ("Seller Rate", "Deduction"):
            values[field] = seller_terms_percent_display(values[field])
        vars_by_field = {field: tk.StringVar(value=values[field]) for field in SELLER_TERMS_FIELDS}
        self.row_vars.append(vars_by_field)
        self._render_rows()

    def _delete_row(self, index: int) -> None:
        if 0 <= index < len(self.row_vars):
            self.row_vars.pop(index)
        if not self.row_vars:
            self._add_row()
            return
        self._render_rows()

    def _bind_rate_deduction_exclusivity(self, vars_by_field: dict[str, tk.StringVar]) -> None:
        rate_var = vars_by_field["Seller Rate"]
        deduction_var = vars_by_field["Deduction"]
        if getattr(rate_var, "_lucas_exclusive_bound", False):
            return
        updating = {"active": False}

        def clear_deduction(*_args) -> None:
            if updating["active"] or not rate_var.get().strip():
                return
            updating["active"] = True
            deduction_var.set("")
            updating["active"] = False

        def clear_rate(*_args) -> None:
            if updating["active"] or not deduction_var.get().strip():
                return
            updating["active"] = True
            rate_var.set("")
            updating["active"] = False

        rate_var.trace_add("write", clear_deduction)
        deduction_var.trace_add("write", clear_rate)
        setattr(rate_var, "_lucas_exclusive_bound", True)
        setattr(deduction_var, "_lucas_exclusive_bound", True)

    def _render_rows(self) -> None:
        for child in self.rows_frame.grid_slaves():
            row = int(child.grid_info().get("row") or 0)
            if row > 0:
                child.destroy()
        sheet_types = self._active_sheet_types()
        for index, vars_by_field in enumerate(self.row_vars, start=1):
            self._bind_rate_deduction_exclusivity(vars_by_field)
            bind_single_paste(ttk.Entry(self.rows_frame, textvariable=vars_by_field["Seller"], style="Assign.TEntry", width=26)).grid(row=index, column=0, sticky="ew", padx=(0, 8), pady=(0, 8))
            bind_single_paste(ttk.Combobox(
                self.rows_frame,
                textvariable=vars_by_field["Sheet Type"],
                values=sheet_types,
                style="Assign.TCombobox",
                width=24,
            )).grid(row=index, column=1, sticky="ew", padx=(0, 8), pady=(0, 8))
            bind_single_paste(ttk.Entry(self.rows_frame, textvariable=vars_by_field["Seller Rate"], style="Assign.TEntry", width=14)).grid(row=index, column=2, sticky="ew", padx=(0, 8), pady=(0, 8))
            bind_single_paste(ttk.Entry(self.rows_frame, textvariable=vars_by_field["Deduction"], style="Assign.TEntry", width=14)).grid(row=index, column=3, sticky="ew", padx=(0, 8), pady=(0, 8))
            ttk.Button(self.rows_frame, text="Delete", command=lambda row_index=index - 1: self._delete_row(row_index), style="AssignSoft.TButton").grid(row=index, column=4, sticky="ew", pady=(0, 8))

    def _validated_rows(self) -> list[dict[str, str]] | None:
        active_types = {name.lower(): name for name in self._active_sheet_types()}
        rows: list[dict[str, str]] = []
        seen: set[tuple[str, str]] = set()
        for index, vars_by_field in enumerate(self.row_vars, start=1):
            row = {field: var.get().strip() for field, var in vars_by_field.items()}
            if not any(row.values()):
                continue
            if not row["Seller"]:
                self.status.set(f"Row {index}: Seller is required.")
                return None
            if not row["Sheet Type"]:
                self.status.set(f"Row {index}: Sheet Type is required.")
                return None
            if row["Sheet Type"].lower() not in active_types:
                self.status.set(f"Row {index}: Sheet Type must match an active Company Rule.")
                return None
            if not row["Seller Rate"] and not row["Deduction"]:
                self.status.set(f"Row {index}: enter Seller Rate or Deduction.")
                return None
            if row["Seller Rate"] and row["Deduction"]:
                self.status.set(f"Row {index}: use Seller Rate or Deduction, not both.")
                return None
            if row["Seller Rate"] and not seller_terms_percent_input_is_number(row["Seller Rate"]):
                self.status.set(f"Row {index}: Seller Rate % must be a number only.")
                return None
            if row["Deduction"] and not seller_terms_percent_input_is_number(row["Deduction"]):
                self.status.set(f"Row {index}: Deduction % must be a number only.")
                return None
            if row["Seller Rate"] and seller_terms_rate(row["Seller Rate"]) is None:
                self.status.set(f"Row {index}: Seller Rate % is invalid.")
                return None
            if row["Deduction"] and seller_terms_rate(row["Deduction"]) is None:
                self.status.set(f"Row {index}: Deduction % is invalid.")
                return None
            if row["Seller Rate"] and (seller_terms_rate(row["Seller Rate"]) or 0) > 1:
                self.status.set(f"Row {index}: Seller Rate % cannot be above 100.")
                return None
            if row["Deduction"] and (seller_terms_rate(row["Deduction"]) or 0) > 1:
                self.status.set(f"Row {index}: Deduction % cannot be above 100.")
                return None
            key = (row["Seller"].lower(), row["Sheet Type"].lower())
            if key in seen:
                self.status.set(f"Row {index}: duplicate Seller and Sheet Type.")
                return None
            seen.add(key)
            row["Sheet Type"] = active_types[row["Sheet Type"].lower()]
            rows.append(row)
        return rows

    def _save(self) -> None:
        rows = self._validated_rows()
        if rows is None:
            return
        try:
            write_seller_terms_rows(self.seller_terms_path, rows)
        except Exception as error:
            messagebox.showerror("Save People Rules", str(error))
            return
        self.status.set(f"Saved {len(rows)} People Rule(s).")
        if self.on_saved:
            self.on_saved()


def open_people_rules_dialog(parent: tk.Misc, pipeline_root: Path, companies: list[dict[str, Any]], on_saved: Callable[[], None] | None = None) -> None:
    dialog = PeopleRulesDialog(parent, pipeline_root, companies, on_saved)
    dialog.focus_set()
    dialog.grab_set()


def source_kind_for_path(value: Any) -> str:
    if isinstance(value, dict):
        if value.get("kind") == "google_keep":
            return "keep_file"
        if value.get("kind") == "google_sheet":
            return "sheet_file"
        value = value.get("path") or value.get("file") or value.get("url")
    path = str(value or "").lower()
    if not path:
        return "manual"
    if path.endswith((".xlsx", ".xlsm", ".csv", ".gsheet")):
        return "sheet_file"
    if path.endswith((".txt", ".md")):
        return "keep_file"
    return "manual" if is_generated_rules_path(path) else "keep_file"


def is_generated_rules_path(value: Any) -> bool:
    return str(value or "").lower().endswith("-rules.json")


def is_generated_payout_path(value: Any) -> bool:
    if isinstance(value, dict):
        value = value.get("path") or value.get("file") or value.get("url")
    return str(value or "").lower().endswith("-payout.json")


def is_same_file_payout_source(value: Any) -> bool:
    return isinstance(value, dict) and (
        bool(value.get("linked_payouts_to_rule_source"))
        or str(value.get("sheet_name") or value.get("sheet") or "").strip().lower() == "payouts"
    )


def display_source_path(value: Any) -> str:
    if isinstance(value, dict):
        if value.get("kind") == "google_keep":
            return str(value.get("url") or value.get("path") or value.get("file") or "")
        return str(value.get("path") or value.get("file") or value.get("url") or "")
    return str(value or "")


def safe_stem(value: str) -> str:
    stem = re.sub(r"[^a-z0-9]+", "-", value.lower()).strip("-")
    return stem or "company"


def keep_note_display_name(url: str) -> str:
    match = re.search(r"/notes/([^/?#]+)", str(url or ""))
    return f"Google Keep {match.group(1)}" if match else "Google Keep note"


def is_google_sheet_url(value: object) -> bool:
    parsed = urllib.parse.urlparse(str(value or "").strip())
    return parsed.scheme in {"http", "https"} and parsed.netloc.lower().endswith("docs.google.com") and "/spreadsheets/" in parsed.path


def google_sheet_display_name(url: str) -> str:
    match = re.search(r"/spreadsheets/d/([^/?#]+)", str(url or ""))
    return f"Google Sheet {match.group(1)}" if match else "Google Sheet"


def title_case(value: str) -> str:
    return " ".join(part.capitalize() for part in value.split())


def category_label(value: str) -> str:
    labels = {
        "wwe": "WWE",
        "f1": "F1",
        "ufc": "UFC",
        "star wars": "Star Wars",
    }
    return labels.get(str(value or "").lower(), title_case(value))


def split_values(value: Any) -> list[str]:
    if value is None:
        return []
    if isinstance(value, list):
        return [str(item).strip() for item in value if str(item).strip()]
    return [item.strip() for item in str(value).split(",") if item.strip()]


def normalize_value_source(value: Any) -> str:
    raw = str(value or "").strip().lower()
    if raw in {"card_ladder", "cardladder", "cl", "card ladder", "card_ladder_value"}:
        return "card_ladder"
    if raw in {"cy_estimate", "cyestimate", "cy", "cy value", "cy_value", "cy estimate", "courtyard", "courtyard_estimate", "court yard", "estimate"}:
        return "cy_estimate"
    return "comps"
