from __future__ import annotations

import json
import re
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog, ttk
from typing import Any, Callable

from assignment_engine import (
    CONFIG_PATH,
    gsheet_shortcut_url,
    load_gsheet_shortcut,
    normalize_source_value,
    read_source_text,
    safe_filename,
)
from google_sheets_import import authorize_google_sheets


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
    "keep_file": "Google Keep local file",
    "sheet_file": "Google Sheets local file",
}
PAYOUT_SOURCE_LABELS = {
    "manual": "Manual payout tiers",
    "file": "Local payout file",
}


class AssignmentRulesDialog(tk.Toplevel):
    def __init__(self, parent: tk.Tk, pipeline_root: Path, on_saved: Callable[[], None]) -> None:
        super().__init__(parent)
        self.title("Assignment Rules")
        self.geometry("1380x920")
        self.minsize(1240, 860)
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
        self.use_card_ladder_value = tk.BooleanVar(value=False)
        self.rule_source_mode = tk.StringVar(value="manual")
        self.rule_source_path = tk.StringVar()
        self.link_payouts_to_rule_source = tk.BooleanVar(value=False)
        self.payout_source_mode = tk.StringVar(value="manual")
        self.payout_source_path = tk.StringVar()
        self.status = tk.StringVar(value="Create or edit a company, then save.")
        self.preview_status = tk.StringVar(value="No source file selected.")
        self.rule_materialized_source: dict[str, Any] | None = None
        self.payout_materialized_source: dict[str, Any] | None = None

        self._configure_styles()
        self._build_ui()
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
        shell = ttk.Frame(self, style="Assign.TFrame", padding=16)
        shell.pack(fill=tk.BOTH, expand=True)
        shell.columnconfigure(1, weight=1)
        shell.rowconfigure(1, weight=1)

        ttk.Label(shell, text="Assignment Rules", style="AssignHeader.TLabel").grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 12))

        side = ttk.Frame(shell, style="AssignPanel.TFrame", padding=12)
        side.grid(row=1, column=0, sticky="ns", padx=(0, 12))
        ttk.Label(side, text="Companies", style="AssignTitle.TLabel").pack(anchor=tk.W)
        self.company_list = tk.Frame(
            side,
            width=280,
            height=620,
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
        ttk.Entry(details, textvariable=self.company_name, style="Assign.TEntry").grid(row=0, column=1, sticky="ew")
        ttk.Checkbutton(
            details,
            text="Require Card Ladder value instead of Comps",
            variable=self.use_card_ladder_value,
            style="Assign.TCheckbutton",
        ).grid(row=1, column=1, sticky=tk.W, pady=(8, 0))

        sources = ttk.Frame(main, style="AssignPanel.TFrame", padding=12)
        sources.grid(row=1, column=0, sticky="ew", pady=(12, 0))
        sources.columnconfigure(0, weight=1)
        sources.columnconfigure(1, weight=1)
        self._build_rule_source_panel(sources).grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        self._build_payout_source_panel(sources).grid(row=0, column=1, sticky="nsew", padx=(8, 0))

        self.body = ttk.Frame(main, style="Assign.TFrame")
        self.body.grid(row=2, column=0, sticky="nsew", pady=(12, 0))
        self.body.columnconfigure(0, weight=1)
        self.body.columnconfigure(1, weight=1)
        self.body.rowconfigure(0, weight=1)

        self.manual_rule_panel = ttk.Frame(self.body, style="AssignPanel.TFrame", padding=12)
        self.manual_rule_panel.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        self.manual_rule_panel.columnconfigure(0, weight=1)
        self.manual_rule_panel.rowconfigure(2, weight=1)
        rule_header = ttk.Frame(self.manual_rule_panel, style="AssignPanel.TFrame")
        rule_header.grid(row=0, column=0, sticky="ew")
        ttk.Label(rule_header, text="Manual Rule Builder", style="AssignTitle.TLabel").pack(side=tk.LEFT)
        ttk.Button(rule_header, text="Add Rule", command=self._add_rule_row, style="AssignSoft.TButton").pack(side=tk.RIGHT)
        ttk.Label(self.manual_rule_panel, text="Used when Rule Source is Manual rules.", style="AssignMuted.TLabel").grid(row=1, column=0, sticky=tk.W, pady=(3, 8))
        rules_view = ttk.Frame(self.manual_rule_panel, style="AssignPanel.TFrame")
        rules_view.grid(row=2, column=0, sticky="nsew")
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
        self.rule_path_entry = ttk.Entry(path_row, textvariable=self.rule_source_path, style="Assign.TEntry")
        self.rule_path_entry.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        ttk.Button(path_row, text="Browse", command=self._browse_rule_source, style="AssignSoft.TButton").grid(row=0, column=1)
        actions = ttk.Frame(frame, style="AssignPanel.TFrame")
        actions.grid(row=4, column=0, sticky="ew", pady=(8, 0))
        actions.columnconfigure(0, weight=1)
        actions.columnconfigure(1, weight=1)
        ttk.Button(actions, text="Preview Source", command=self._preview_rule_source, style="AssignSoft.TButton").grid(row=0, column=0, sticky="ew", padx=(0, 4))
        ttk.Button(actions, text="Connect Google", command=self._connect_google, style="AssignSoft.TButton").grid(row=0, column=1, sticky="ew", padx=(4, 0))
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

    def _build_payout_source_panel(self, parent: ttk.Frame) -> ttk.Frame:
        frame = ttk.LabelFrame(parent, text="Payout Source", style="Assign.TLabelframe", padding=10)
        for row, (value, label) in enumerate(PAYOUT_SOURCE_LABELS.items()):
            ttk.Radiobutton(frame, text=label, value=value, variable=self.payout_source_mode, command=self._on_source_mode_change, style="Assign.TRadiobutton").grid(row=row, column=0, sticky=tk.W)
        path_row = ttk.Frame(frame, style="AssignPanel.TFrame")
        path_row.grid(row=2, column=0, sticky="ew", pady=(8, 0))
        path_row.columnconfigure(0, weight=1)
        self.payout_path_entry = ttk.Entry(path_row, textvariable=self.payout_source_path, style="Assign.TEntry")
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
        for index, company in enumerate(self.companies):
            self._add_company_row(index, company)

    def _add_company_row(self, index: int, company: dict[str, Any]) -> None:
        active = company.get("active") is not False
        selected = index == self.selected_index
        bg = "#242424" if selected else "#1f1f1f"
        row = tk.Frame(self.company_list, bg=bg, padx=5, pady=4)
        row.pack(fill=tk.X, padx=4, pady=(4 if index == 0 else 0, 0))
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
        state = "active" if company.get("active") is not False else "inactive"
        self.status.set(f"{company.get('name') or 'Company'} is now {state}. Save & Reload to apply in Assignment.")

    def _select_company(self, index: int) -> None:
        self.selected_index = index
        self._refresh_company_list()
        company = self.companies[index]
        self.company_name.set(str(company.get("name") or ""))
        self.use_card_ladder_value.set(str(company.get("value_source") or company.get("valueSource") or "").strip().lower() in {"card_ladder", "cardladder", "cl", "card ladder"})
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
        self.use_card_ladder_value.set(False)
        self.rule_source_mode.set("manual")
        self.rule_source_path.set("")
        self.link_payouts_to_rule_source.set(False)
        self.payout_source_mode.set("manual")
        self.payout_source_path.set("")
        self.rule_materialized_source = None
        self.payout_materialized_source = None
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
        ttk.Entry(price_frame, textvariable=min_var, width=12, style="Assign.TEntry").grid(row=1, column=1, sticky=tk.W)
        ttk.Label(price_frame, text="Max", style="Assign.TLabel").grid(row=1, column=2, sticky=tk.W, padx=(14, 6))
        ttk.Entry(price_frame, textvariable=max_var, width=12, style="Assign.TEntry").grid(row=1, column=3, sticky=tk.W)
        ttk.Label(price_frame, text="Payout Percentage", style="Assign.TLabel").grid(row=1, column=4, sticky=tk.W, padx=(14, 6))
        ttk.Entry(price_frame, textvariable=payout_var, width=12, style="Assign.TEntry").grid(row=1, column=5, sticky=tk.W)

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
            ttk.Entry(grades_frame, textvariable=min_grade, width=8, style="Assign.TEntry").grid(row=grade_index, column=2, sticky=tk.W, pady=3)
            ttk.Entry(grades_frame, textvariable=max_grade, width=8, style="Assign.TEntry").grid(row=grade_index, column=3, sticky=tk.W, pady=3)
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
        ttk.Entry(self.payout_frame, textvariable=min_var, width=9, style="Assign.TEntry").grid(row=row_index, column=1, sticky=tk.W, pady=4)
        ttk.Label(self.payout_frame, text="Max", style="Assign.TLabel").grid(row=row_index, column=2, sticky=tk.W, padx=(10, 6), pady=4)
        ttk.Entry(self.payout_frame, textvariable=max_var, width=9, style="Assign.TEntry").grid(row=row_index, column=3, sticky=tk.W, pady=4)
        ttk.Label(self.payout_frame, text="Rate", style="Assign.TLabel").grid(row=row_index, column=4, sticky=tk.W, padx=(10, 6), pady=4)
        ttk.Entry(self.payout_frame, textvariable=rate_var, width=9, style="Assign.TEntry").grid(row=row_index, column=5, sticky=tk.W, pady=4)
        self.payout_rows.append({"min": min_var, "max": max_var, "rate": rate_var})

    def _save_company(self) -> bool:
        name = self.company_name.get().strip()
        if not name:
            messagebox.showinfo("Company name", "Name the company before saving.")
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
            "value_source": "card_ladder" if self.use_card_ladder_value.get() else "comps",
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
        rules_path.write_text(json.dumps(self._rules_payload(), indent=2), encoding="utf-8")
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
        if isinstance(cached, dict) and normalize_source_value(cached.get("path")) == normalized:
            return cached
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
            messagebox.showerror("Connect Google", str(error))
            return
        self.preview_status.set("Google Sheets connected. Preview the source to read the live sheet.")

    def _save_and_reload(self) -> None:
        if self._save_company():
            self.on_saved()
            self.status.set("Saved and reloaded Assignment rules.")

    def _rules_payload(self) -> dict[str, Any]:
        return {
            "rules": [self._rule_payload(row) for row in self.rule_rows],
            "blocks": [],
        }

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


def source_kind_for_path(value: Any) -> str:
    if isinstance(value, dict):
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
        return str(value.get("path") or value.get("file") or value.get("url") or "")
    return str(value or "")


def safe_stem(value: str) -> str:
    stem = re.sub(r"[^a-z0-9]+", "-", value.lower()).strip("-")
    return stem or "company"


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
