from __future__ import annotations

import json
import os
import queue
import sys
import threading
import time
import types
import unittest
import urllib.request
from datetime import datetime, timedelta
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))
PHOTO_APP = ROOT / "photo_tool" / "app"
if str(PHOTO_APP) not in sys.path:
    sys.path.insert(0, str(PHOTO_APP))

import app
import assignment_config_ui
import assignment_engine
import bridge_server
import google_sheets_import
import lucas_diagnostics
from comp_engine.workbook_io import WorkbookRow
from intake_io import append_company_sheet_rows, ensure_company_weekly_sheets, mark_received_in_workbooks, read_company_profit_records, read_simple_spreadsheet, write_pipeline_output, write_working_sheet
from shared_state import atomic_write_json, local_identity, read_json, shared_lock


if "google" not in sys.modules:
    google_module = types.ModuleType("google")
    genai_module = types.ModuleType("google.genai")
    genai_types_module = types.ModuleType("google.genai.types")

    class _Client:
        pass

    class _Part:
        @staticmethod
        def from_bytes(data=None, mime_type=None):
            return {"data": data, "mime_type": mime_type}

    class _GenerateContentConfig:
        def __init__(self, **kwargs):
            self.kwargs = kwargs

    class _ThinkingConfig:
        def __init__(self, **kwargs):
            self.kwargs = kwargs

    genai_module.Client = _Client
    genai_types_module.Part = _Part
    genai_types_module.GenerateContentConfig = _GenerateContentConfig
    genai_types_module.ThinkingConfig = _ThinkingConfig
    genai_module.types = genai_types_module
    google_module.genai = genai_module
    sys.modules["google"] = google_module
    sys.modules["google.genai"] = genai_module
    sys.modules["google.genai.types"] = genai_types_module

import multi_card_extraction


class SharedStateTests(unittest.TestCase):
    def test_google_ssl_context_uses_certifi_when_no_cert_env_is_set(self) -> None:
        with TemporaryDirectory() as tmp:
            cafile = Path(tmp) / "cacert.pem"
            cafile.write_text("", encoding="utf-8")
            fake_certifi = types.SimpleNamespace(where=lambda: str(cafile))
            old_context = google_sheets_import._SSL_CONTEXT
            old_diagnostics = dict(google_sheets_import.LAST_OAUTH_DIAGNOSTICS)
            old_ssl_cert = os.environ.pop("SSL_CERT_FILE", None)
            old_requests_bundle = os.environ.pop("REQUESTS_CA_BUNDLE", None)
            try:
                google_sheets_import._SSL_CONTEXT = None
                sentinel_context = object()
                with patch.dict(sys.modules, {"certifi": fake_certifi}), patch(
                    "google_sheets_import.ssl.create_default_context",
                    return_value=sentinel_context,
                ) as create_context:
                    context = google_sheets_import.google_ssl_context()
                self.assertIs(context, sentinel_context)
                create_context.assert_called_once_with(cafile=str(cafile))
                self.assertEqual(os.environ.get("SSL_CERT_FILE"), str(cafile))
                self.assertEqual(os.environ.get("REQUESTS_CA_BUNDLE"), str(cafile))
                self.assertEqual(google_sheets_import.LAST_OAUTH_DIAGNOSTICS["ssl_cert_file"], str(cafile))
            finally:
                google_sheets_import._SSL_CONTEXT = old_context
                google_sheets_import.LAST_OAUTH_DIAGNOSTICS.clear()
                google_sheets_import.LAST_OAUTH_DIAGNOSTICS.update(old_diagnostics)
                if old_ssl_cert is not None:
                    os.environ["SSL_CERT_FILE"] = old_ssl_cert
                else:
                    os.environ.pop("SSL_CERT_FILE", None)
                if old_requests_bundle is not None:
                    os.environ["REQUESTS_CA_BUNDLE"] = old_requests_bundle
                else:
                    os.environ.pop("REQUESTS_CA_BUNDLE", None)

    def test_shared_lock_serializes_concurrent_writers(self) -> None:
        with TemporaryDirectory() as tmp:
            root = Path(tmp)
            events: list[tuple[str, str, float]] = []

            def worker(name: str, delay: float) -> None:
                with shared_lock(root, "same-file", {"display_name": name, "machine": name}):
                    events.append((name, "enter", time.time()))
                    time.sleep(delay)
                    events.append((name, "exit", time.time()))

            first = threading.Thread(target=worker, args=("A", 0.25))
            second = threading.Thread(target=worker, args=("B", 0.01))
            first.start()
            time.sleep(0.05)
            second.start()
            first.join()
            second.join()

            self.assertEqual([event[:2] for event in events], [("A", "enter"), ("A", "exit"), ("B", "enter"), ("B", "exit")])

    def test_atomic_json_write_and_local_identity(self) -> None:
        with TemporaryDirectory() as tmp:
            root = Path(tmp)
            settings_path = root / "lucas_settings.json"
            identity = local_identity(settings_path)
            self.assertTrue(identity["user_id"])
            self.assertTrue(identity["machine"])

            state_path = root / "state.json"
            atomic_write_json(state_path, {"a": 1})
            atomic_write_json(state_path, {"b": [1, 2, 3]})
            self.assertEqual(read_json(state_path, {}), {"b": [1, 2, 3]})
            self.assertFalse(list(root.glob("*.tmp")))


class WorkbookCompanyProfitTests(unittest.TestCase):
    def test_company_sheet_week_start_rolls_forward_sunday_midnight(self) -> None:
        self.assertEqual(app.company_sheet_week_start_for_time(datetime(2026, 6, 13, 23, 59)).isoformat(), "2026-06-08")
        self.assertEqual(app.company_sheet_week_start_for_time(datetime(2026, 6, 14, 0, 0)).isoformat(), "2026-06-15")
        self.assertEqual(app.company_sheet_week_start_for_time(datetime(2026, 6, 14, 23, 59)).isoformat(), "2026-06-15")
        self.assertEqual(app.company_sheet_week_start_for_time(datetime(2026, 6, 15, 8, 0)).isoformat(), "2026-06-15")

    def test_ensure_company_weekly_sheets_creates_blank_company_tabs(self) -> None:
        with TemporaryDirectory() as tmp:
            root = Path(tmp) / "COMPANY SHEETS"

            result = ensure_company_weekly_sheets(root, ["Arena Club", "Fanatics"], app.company_sheet_week_start_for_time(datetime(2026, 6, 14, 0, 0)))
            repeat = ensure_company_weekly_sheets(root, ["Arena Club", "Fanatics"], app.company_sheet_week_start_for_time(datetime(2026, 6, 14, 0, 0)))

            self.assertEqual(len(result["created"]), 2)
            self.assertEqual(len(repeat["existing"]), 2)
            arena_path = root / "Arena Club" / "Arena Club.xlsx"
            self.assertTrue(arena_path.exists())
            workbook = load_workbook(arena_path, read_only=True, data_only=True)
            self.assertIn("Week of 2026-06-15", workbook.sheetnames)
            workbook.close()
            rows = read_simple_spreadsheet(arena_path)
            self.assertEqual(rows, [])

    def test_cy_sheet_estimate_and_confidence_are_preserved_separately_from_purchase(self) -> None:
        with TemporaryDirectory() as tmp:
            path = Path(tmp) / "cy.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.append(["Grader", "Cert", "Description", "Grade", "Purchase", "Estimate", "Confidence"])
            sheet.append(["CGC", "1401045404276", "2022 Paradigm Trigger Unown VSTAR", "g10", 17.37, 19.30, 4])
            workbook.save(path)

            rows = read_simple_spreadsheet(path)

            self.assertEqual(len(rows), 1)
            self.assertEqual(rows[0]["purchase_price"], 17.37)
            self.assertEqual(rows[0]["cy_value"], 19.30)
            self.assertEqual(rows[0]["cy_confidence"], 4)

    def test_sheet_loader_uses_headers_when_date_column_shifts_values(self) -> None:
        with TemporaryDirectory() as tmp:
            path = Path(tmp) / "shifted.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.append(["Date", "Cert", "Grader", "Card", "Comp Value", "Purchase Price", "CL Value"])
            sheet.append(["2026-06-24", "137915162", "PSA", "Test Card PSA 10", 120, 45, 150])
            workbook.save(path)

            rows = read_simple_spreadsheet(path)

            self.assertEqual(len(rows), 1)
            self.assertEqual(rows[0]["cert_number"], "137915162")
            self.assertEqual(rows[0]["purchase_price"], 45)
            self.assertEqual(rows[0]["card_ladder_comps_average"], 120)
            self.assertEqual(rows[0]["card_ladder_value"], 150)
            self.assertEqual(rows[0]["date_added"], "2026-06-24")

    def test_sheet_loader_detects_header_row_below_title_rows(self) -> None:
        with TemporaryDirectory() as tmp:
            path = Path(tmp) / "title-row.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.append(["Network mode test upload"])
            sheet.append(["Date", "Cert", "Grader", "Card", "Comps", "Purchase"])
            sheet.append(["2026-06-24", "22222222", "PSA", "Another Test Card PSA 9", 80, 25])
            workbook.save(path)

            rows = read_simple_spreadsheet(path)

            self.assertEqual(len(rows), 1)
            self.assertEqual(rows[0]["purchase_price"], 25)
            self.assertEqual(rows[0]["card_ladder_comps_average"], 80)

    def test_courtyard_weekly_sheet_uses_cy_ingest_format(self) -> None:
        with TemporaryDirectory() as tmp:
            company_dir = Path(tmp) / "COMPANY SHEETS"
            rows = [
                WorkbookRow(
                    excel_row=2,
                    cert_number="1401045404276",
                    grader="CGC",
                    category="pokemon",
                    card_title="2022 Paradigm Trigger Unown VSTAR CGC 10",
                    existing_value=17.37,
                    card_ladder_value=22,
                    card_ladder_comps_average=21,
                    cy_value=19.30,
                    cy_confidence=4,
                    best_company="CourtYard",
                    estimated_payout=18.33,
                    status="Received",
                )
            ]

            result = append_company_sheet_rows(company_dir, rows, {2: "cy.xlsx:2"}, {2: "cy.xlsx"})

            self.assertEqual(result["rows_added"], 1)
            path = company_dir / "CourtYard" / "CourtYard.xlsx"
            workbook = load_workbook(path, data_only=True)
            sheet = workbook[workbook.sheetnames[0]]
            try:
                first_headers = [sheet.cell(1, column).value for column in range(1, 8)]
                self.assertEqual(first_headers, ["Grader", "Cert", "Description", "Grade", "Purchase", "Estimate", "Confidence"])
                self.assertTrue(sheet.column_dimensions["H"].hidden)
            finally:
                workbook.close()

            imported = read_simple_spreadsheet(path)
            self.assertEqual(imported[0]["grader"], "CGC")
            self.assertEqual(imported[0]["cert_number"], "1401045404276")
            self.assertEqual(imported[0]["purchase_price"], 17.37)
            self.assertEqual(imported[0]["cy_value"], 19.30)
            self.assertEqual(imported[0]["cy_confidence"], 4)
            profit_records = read_company_profit_records(company_dir)
            self.assertEqual(len(profit_records), 1)
            self.assertEqual(profit_records[0]["sale_price"], 18.33)
            self.assertEqual(profit_records[0]["source_sheet"], "cy.xlsx")

    def test_working_sheet_round_trips_manual_sport(self) -> None:
        with TemporaryDirectory() as tmp:
            path = Path(tmp) / "manual-sport.xlsx"
            rows = [
                WorkbookRow(
                    excel_row=2,
                    cert_number="0010355805",
                    grader="BGS",
                    category="baseball",
                    card_title="Generic Prospect Auto BGS 9.5",
                    existing_value=25,
                )
            ]

            write_working_sheet(path, rows)
            loaded = read_simple_spreadsheet(path)

            self.assertEqual(loaded[0]["sport"], "baseball")

    def test_receive_company_append_dedupes_and_profit_backfills(self) -> None:
        with TemporaryDirectory() as tmp:
            root = Path(tmp)
            working_dir = root / "WORKING SHEETS"
            company_dir = root / "COMPANY SHEETS"
            working_dir.mkdir()

            source_rows = [
                WorkbookRow(excel_row=2, cert_number="11111111", grader="PSA", card_title="Card One PSA 10", existing_value=50),
                WorkbookRow(excel_row=3, cert_number="22222222", grader="PSA", card_title="Card Two PSA 9", existing_value=75),
            ]
            sheet_path = working_dir / "test.xlsx"
            write_working_sheet(sheet_path, source_rows)

            receive_result = mark_received_in_workbooks([sheet_path], {"11111111"})
            self.assertEqual(receive_result["rows_marked"], 1)
            self.assertIn("11111111", receive_result["certs_marked"])
            self.assertEqual(len(read_simple_spreadsheet(sheet_path)), 2)

            sold_row = WorkbookRow(
                excel_row=2,
                cert_number="11111111",
                grader="PSA",
                card_title="Card One PSA 10",
                existing_value=50,
                card_ladder_value=100,
                card_ladder_comps_average=100,
                cy_value=88,
                cy_confidence=4,
                best_company="Arena Club",
                estimated_payout=90,
                company_pile=True,
                status="Received",
            )
            first_append = append_company_sheet_rows(company_dir, [sold_row], {2: "test.xlsx:2"}, {2: "test.xlsx"})
            second_append = append_company_sheet_rows(company_dir, [sold_row], {2: "test.xlsx:2"}, {2: "test.xlsx"})

            self.assertEqual(first_append["rows_added"], 1)
            self.assertEqual(second_append["rows_added"], 0)
            self.assertEqual(len(first_append["added_records"]), 1)

            profit_records = read_company_profit_records(company_dir)
            self.assertEqual(len(profit_records), 1)
            self.assertEqual(profit_records[0]["cy_value"], 88.0)
            self.assertEqual(profit_records[0]["cy_confidence"], 4)
            self.assertEqual(profit_records[0]["purchase_price"], 50.0)
            self.assertEqual(profit_records[0]["sale_price"], 90.0)

    def test_manual_company_sheet_rows_without_source_sheet_do_not_backfill_profit(self) -> None:
        with TemporaryDirectory() as tmp:
            company_dir = Path(tmp) / "COMPANY SHEETS"
            path = company_dir / "Arena Club" / "Arena Club.xlsx"
            path.parent.mkdir(parents=True)
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Week of 2026-06-22"
            sheet.append(["Date Added", "Source Sheet", "Source", "Certification Number", "Grader", "Card Description", "Purchase Price", "Card Ladder Value", "Comps", "CY Estimate", "CY Confidence", "Best Company", "Estimated Payout", "Status", "Notes"])
            sheet.append(["2026-06-22", "", "Manual", "999", "PSA", "Manual Card PSA 10", 10, "", 20, "", "", "Arena Club", 18, "Received", ""])
            sheet.append(["2026-06-22", "Lot A.xlsx", "LUCAS", "111", "PSA", "LUCAS Card PSA 10", 10, "", 20, "", "", "Arena Club", 18, "Received", ""])
            workbook.save(path)

            records = read_company_profit_records(company_dir)

            self.assertEqual(len(records), 1)
            self.assertEqual(records[0]["cert_number"], "111")

    def test_pipeline_output_round_trips_cy_value(self) -> None:
        with TemporaryDirectory() as tmp:
            path = Path(tmp) / "comps.xlsx"
            rows = [
                WorkbookRow(
                    excel_row=2,
                    cert_number="11111111",
                    grader="PSA",
                    card_title="Card One PSA 10",
                    existing_value=50,
                    card_ladder_value=100,
                    card_ladder_comps_average=95,
                    cy_value=87.5,
                    cy_confidence=3,
                    status="Card Ladder OK",
                )
            ]

            write_pipeline_output(path, rows, {2: "source"})
            reloaded = read_simple_spreadsheet(path)

            self.assertEqual(reloaded[0]["cy_value"], 87.5)
            self.assertEqual(reloaded[0]["cy_confidence"], 3)


class CYLookupTests(unittest.TestCase):
    def test_cardladder_extension_error_is_recorded_on_row(self) -> None:
        state = bridge_server.BridgeState()
        row = WorkbookRow(excel_row=2, cert_number="156327815", grader="PSA", card_title="")
        result = {
            "excelRow": 2,
            "certNumber": "156327815",
            "grader": "PSA",
            "status": "extension_error",
            "error": "Card Ladder lookup failed after the cert opened.",
            "ocr": {"debugImage": ""},
            "extensionVersion": app.EXPECTED_CARDLADDER_EXTENSION_VERSION,
        }

        state._apply_cardladder_result_to_row(row, result)

        self.assertEqual(row.status, "Card Ladder extension error")
        self.assertIsNone(row.card_ladder_value)
        self.assertEqual(row.card_ladder_comps, "")
        self.assertIn("lookup failed", row.notes)

    def test_cardladder_result_triggers_cy_lookup_on_mac(self) -> None:
        state = bridge_server.BridgeState()
        row = WorkbookRow(
            excel_row=2,
            cert_number="11111111",
            grader="PSA",
            card_title="Card One PSA 10",
        )
        state.set_rows([row])

        with patch.object(bridge_server, "cy_lookup_enabled", return_value=True), \
                patch.object(bridge_server, "lookup_cy_buy_price", return_value=(87.5, "ok")):
            state.post_cardladder_result(
                {
                    "excelRow": 2,
                    "certNumber": "11111111",
                    "value": "$100",
                    "status": "ok",
                    "ocr": {"comps": [{"price": "$90", "date_sold": "2026-06-01"}]},
                }
            )

            deadline = time.time() + 2
            while row.cy_value is None and time.time() < deadline:
                time.sleep(0.01)

        self.assertEqual(row.card_ladder_value, 100)
        self.assertEqual(row.cy_value, 87.5)
        self.assertEqual(row.status, "Card Ladder OK")
        self.assertIn("CY value: $87.50", row.notes)

    def test_cardladder_status_survives_cy_unavailable(self) -> None:
        state = bridge_server.BridgeState()
        row = WorkbookRow(
            excel_row=2,
            cert_number="11111111",
            grader="PSA",
            card_title="Card One PSA 10",
        )
        state.set_rows([row])

        with patch.object(bridge_server, "cy_lookup_enabled", return_value=True), \
                patch.object(bridge_server, "lookup_cy_buy_price", return_value=(None, "not found")):
            state.post_cardladder_result(
                {
                    "excelRow": 2,
                    "certNumber": "11111111",
                    "value": "$100",
                    "status": "ok",
                    "ocr": {"comps": [{"price": "$90", "date_sold": "2026-06-01"}]},
                }
            )

            deadline = time.time() + 2
            while "CY lookup:" not in row.notes and time.time() < deadline:
                time.sleep(0.01)

        self.assertEqual(row.card_ladder_value, 100)
        self.assertIsNone(row.cy_value)
        self.assertEqual(row.status, "Card Ladder OK")
        self.assertIn("CY lookup: not found", row.notes)

    def test_cy_lookup_waits_until_cardladder_finishes(self) -> None:
        state = bridge_server.BridgeState()
        row = WorkbookRow(
            excel_row=2,
            cert_number="11111111",
            grader="PSA",
            card_title="Card One PSA 10",
        )
        state.set_rows([row])
        state.cardladder_running = True

        with patch.object(bridge_server, "cy_lookup_enabled", return_value=True), \
                patch.object(bridge_server, "lookup_cy_buy_price", return_value=(87.5, "ok")) as lookup:
            state.post_cardladder_result(
                {
                    "excelRow": 2,
                    "certNumber": "11111111",
                    "value": "$100",
                    "status": "ok",
                    "ocr": {"comps": [{"price": "$90", "date_sold": "2026-06-01"}]},
                }
            )
            time.sleep(0.05)
            lookup.assert_not_called()
            self.assertIsNone(row.cy_value)

            state.finish_cardladder({})
            deadline = time.time() + 2
            while row.cy_value is None and time.time() < deadline:
                time.sleep(0.01)

        self.assertEqual(row.cy_value, 87.5)

    def test_cy_only_batch_closes_app_after_last_lookup(self) -> None:
        state = bridge_server.BridgeState()
        rows = [
            WorkbookRow(excel_row=2, cert_number="11111111", grader="PSA", card_title="Card One PSA 10"),
        ]
        state.set_rows(rows)

        with patch.object(bridge_server, "cy_lookup_enabled", return_value=True), \
                patch.object(bridge_server, "lookup_cy_buy_price", return_value=(87.5, "ok")), \
                patch.object(bridge_server, "close_cy_adapter") as close_cy:
            state.start_cy_lookups(rows)
            deadline = time.time() + 2
            while any(row.cy_value is None for row in rows) and time.time() < deadline:
                time.sleep(0.01)
            while close_cy.call_count == 0 and time.time() < deadline:
                time.sleep(0.01)

        self.assertEqual([row.cy_value for row in rows], [87.5])
        self.assertEqual(rows[0].status, "CY OK")
        self.assertGreaterEqual(close_cy.call_count, 1)

    def test_cy_only_unavailable_sets_cy_unavailable_status(self) -> None:
        state = bridge_server.BridgeState()
        rows = [
            WorkbookRow(excel_row=2, cert_number="11111111", grader="PSA", card_title="Card One PSA 10"),
        ]
        state.set_rows(rows)

        with patch.object(bridge_server, "cy_lookup_enabled", return_value=True), \
                patch.object(bridge_server, "lookup_cy_buy_price", return_value=(None, "not found")), \
                patch.object(bridge_server, "close_cy_adapter"):
            state.start_cy_lookups(rows)
            deadline = time.time() + 2
            while rows[0].status == "CY queued" and time.time() < deadline:
                time.sleep(0.01)

        self.assertIsNone(rows[0].cy_value)
        self.assertEqual(rows[0].status, "CY unavailable")
        self.assertIn("CY lookup: not found", rows[0].notes)

    def test_cy_gui_lookup_is_serialized(self) -> None:
        entered: list[str] = []
        active = 0
        max_active = 0
        lock = threading.Lock()

        class FakeAdapter:
            def submit_cert_lookup(self, cert_number: str, slab_type: str) -> dict:
                nonlocal active, max_active
                with lock:
                    entered.append(cert_number)
                    active += 1
                    max_active = max(max_active, active)
                time.sleep(0.05)
                with lock:
                    active -= 1
                return {"cy_buy_price": "87.50", "message": "ok"}

        old_lock = bridge_server._CY_LOOKUP_LOCK
        bridge_server._CY_LOOKUP_LOCK = threading.Lock()
        try:
            with patch.object(bridge_server, "get_cy_adapter", return_value=FakeAdapter()):
                threads = [
                    threading.Thread(target=bridge_server.lookup_cy_buy_price, args=(cert, "PSA"))
                    for cert in ("11111111", "22222222")
                ]
                for thread in threads:
                    thread.start()
                for thread in threads:
                    thread.join(timeout=2)
        finally:
            bridge_server._CY_LOOKUP_LOCK = old_lock

        self.assertEqual(set(entered), {"11111111", "22222222"})
        self.assertEqual(max_active, 1)

    def test_cy_lookup_is_disabled_off_mac(self) -> None:
        self.assertFalse(bridge_server.cy_lookup_enabled("win32"))
        self.assertFalse(bridge_server.cy_lookup_enabled("linux"))


class GoogleSheetCacheTests(unittest.TestCase):
    def test_authenticated_google_sheet_export_writes_xlsx_cache(self) -> None:
        def fake_tabs(_url: str, interactive: bool = False, sheet_name: str = ""):
            return [
                ("Rules/Main*?", [["Category", "Value"], ["Baseball", "100"]]),
                ("Payouts", [["CATEGORY", "YOUR PAYOUT %"], ["Baseball", "90%"]]),
            ]

        with TemporaryDirectory() as tmp, patch.object(google_sheets_import, "read_google_sheet_tabs", side_effect=fake_tabs):
            output_path = Path(tmp) / "cache.xlsx"
            google_sheets_import.export_google_sheet_to_xlsx("https://docs.google.com/spreadsheets/d/test/edit", output_path)

            from openpyxl import load_workbook

            workbook = load_workbook(output_path, read_only=True, data_only=True)
            try:
                self.assertEqual(workbook.sheetnames, ["Rules Main", "Payouts"])
                self.assertEqual(workbook["Payouts"].cell(2, 2).value, "90%")
            finally:
                workbook.close()

    def test_oauth_callback_ignores_unrelated_browser_requests(self) -> None:
        server = google_sheets_import.OAuthCallbackServer(("127.0.0.1", 0), google_sheets_import.OAuthCallbackHandler)
        server.timeout = 1
        server.expected_state = "state-123"
        base_url = f"http://127.0.0.1:{server.server_port}"

        def serve_once() -> None:
            server.handle_request()

        try:
            thread = threading.Thread(target=serve_once)
            thread.start()
            with urllib.request.urlopen(f"{base_url}/favicon.ico", timeout=5) as response:
                self.assertEqual(response.status, 204)
            thread.join(timeout=5)
            self.assertFalse(thread.is_alive())
            self.assertFalse(server.done)
            self.assertEqual(server.code, "")
            self.assertEqual(server.error, "")

            thread = threading.Thread(target=serve_once)
            thread.start()
            with urllib.request.urlopen(f"{base_url}/oauth2callback?state=state-123&code=abc123", timeout=5) as response:
                self.assertEqual(response.status, 200)
                body = response.read().decode("utf-8")
            thread.join(timeout=5)
            self.assertFalse(thread.is_alive())
            self.assertTrue(server.done)
            self.assertEqual(server.code, "abc123")
            self.assertEqual(server.error, "")
            self.assertIn("Google Sheets connected", body)
        finally:
            server.server_close()

    def test_startup_google_sheet_cache_refresh_discovers_and_exports_sources(self) -> None:
        class Dummy:
            _saved_google_sheet_sources = app.CardPipelineApp._saved_google_sheet_sources
            _google_sheet_cache_source = app.CardPipelineApp._google_sheet_cache_source
            _refresh_startup_google_sheet_caches = app.CardPipelineApp._refresh_startup_google_sheet_caches

        with TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            old_pipeline = app.CARD_PIPELINE_DIR
            old_config = app.ASSIGNMENT_CONFIG_PATH
            app.CARD_PIPELINE_DIR = tmp_path / "CARD_PIPELINE"
            app.ASSIGNMENT_CONFIG_PATH = tmp_path / "assignment_companies.json"
            cache_path = app.CARD_PIPELINE_DIR / "ASSIGNMENT RULES" / "SHEET EXPORTS" / "rules.xlsx"
            app.ASSIGNMENT_CONFIG_PATH.write_text(
                json.dumps(
                    {
                        "companies": [
                            {
                                "name": "Test",
                                "rules": {
                                    "kind": "google_sheet",
                                    "url": "https://docs.google.com/spreadsheets/d/abc/edit",
                                    "path": str(cache_path),
                                    "name": "Rules",
                                },
                            }
                        ]
                    }
                ),
                encoding="utf-8",
            )
            calls: list[tuple[str, Path, bool]] = []

            def fake_export(url: str, output_path: Path, interactive: bool = False) -> Path:
                calls.append((url, Path(output_path), interactive))
                Path(output_path).parent.mkdir(parents=True, exist_ok=True)
                Path(output_path).write_bytes(b"fake")
                return Path(output_path)

            dummy = Dummy()
            dummy.lucas_identity = {"display_name": "Tester", "machine": "Test"}
            try:
                with patch.object(app, "export_google_sheet_to_xlsx", side_effect=fake_export):
                    result = dummy._refresh_startup_google_sheet_caches()
                self.assertEqual(result["refreshed"], 1)
                self.assertEqual(result["errors"], [])
                self.assertEqual(calls[0][0], "https://docs.google.com/spreadsheets/d/abc/edit")
                self.assertFalse(calls[0][2])
                self.assertTrue(cache_path.exists())
            finally:
                app.CARD_PIPELINE_DIR = old_pipeline
                app.ASSIGNMENT_CONFIG_PATH = old_config

    def test_startup_google_sheet_cache_materializes_url_sources(self) -> None:
        class Dummy:
            _saved_google_sheet_sources = app.CardPipelineApp._saved_google_sheet_sources
            _google_sheet_cache_source = app.CardPipelineApp._google_sheet_cache_source
            _refresh_startup_google_sheet_caches = app.CardPipelineApp._refresh_startup_google_sheet_caches

        with TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            old_pipeline = app.CARD_PIPELINE_DIR
            old_config = app.ASSIGNMENT_CONFIG_PATH
            app.CARD_PIPELINE_DIR = tmp_path / "CARD_PIPELINE"
            app.ASSIGNMENT_CONFIG_PATH = tmp_path / "assignment_companies.json"
            app.ASSIGNMENT_CONFIG_PATH.write_text(
                json.dumps(
                    {
                        "companies": [
                            {
                                "name": "Rules URL",
                                "rules": "https://docs.google.com/spreadsheets/d/raw-url/edit",
                            },
                            {
                                "name": "Payout URL",
                                "payout": {
                                    "kind": "google_sheet",
                                    "url": "https://docs.google.com/spreadsheets/d/dict-url/edit",
                                    "name": "Payout URL",
                                },
                            },
                        ]
                    }
                ),
                encoding="utf-8",
            )
            calls: list[tuple[str, Path, bool]] = []

            def fake_export(url: str, output_path: Path, interactive: bool = False) -> Path:
                calls.append((url, Path(output_path), interactive))
                Path(output_path).parent.mkdir(parents=True, exist_ok=True)
                Path(output_path).write_bytes(b"fake")
                return Path(output_path)

            dummy = Dummy()
            dummy.lucas_identity = {"display_name": "Tester", "machine": "Test"}
            try:
                with patch.object(app, "export_google_sheet_to_xlsx", side_effect=fake_export):
                    result = dummy._refresh_startup_google_sheet_caches()
                self.assertEqual(result["refreshed"], 2)
                self.assertEqual({call[0] for call in calls}, {
                    "https://docs.google.com/spreadsheets/d/raw-url/edit",
                    "https://docs.google.com/spreadsheets/d/dict-url/edit",
                })
                self.assertTrue(all(call[1].parent.name == "SHEET EXPORTS" for call in calls))
            finally:
                app.CARD_PIPELINE_DIR = old_pipeline
                app.ASSIGNMENT_CONFIG_PATH = old_config

    def test_google_keep_source_reads_cached_note_text(self) -> None:
        with TemporaryDirectory() as tmp:
            root = Path(tmp)
            cache_path = root / "ASSIGNMENT RULES" / "KEEP EXPORTS" / "note.txt"
            cache_path.parent.mkdir(parents=True)
            cache_path.write_text("Basketball $10-$100 90%\n", encoding="utf-8")

            text = assignment_engine.read_source_text(
                {
                    "kind": "google_keep",
                    "url": "https://keep.google.com/u/0/#NOTE/abc123",
                    "path": str(cache_path),
                    "name": "Rules",
                },
                root,
            )

            self.assertIn("Basketball $10-$100 90%", text)

    def test_saved_google_keep_source_uses_pipeline_cache_folder(self) -> None:
        with TemporaryDirectory() as tmp:
            root = Path(tmp)
            output_dir = root / "ASSIGNMENT RULES" / "KEEP EXPORTS"
            source = app.CardPipelineApp._google_keep_cache_source(
                object(),
                "https://keep.google.com/u/0/#NOTE/abc123",
                output_dir,
            )

            self.assertIsNotNone(source)
            self.assertEqual(Path(source["path"]).parent, output_dir)

    def test_saved_google_keep_source_rehomes_repo_cache_path(self) -> None:
        with TemporaryDirectory() as tmp:
            root = Path(tmp)
            output_dir = root / "CARD_PIPELINE" / "ASSIGNMENT RULES" / "KEEP EXPORTS"
            old_config = app.ASSIGNMENT_CONFIG_PATH
            app.ASSIGNMENT_CONFIG_PATH = root / "assignment_companies.json"
            try:
                stale_path = root / "ASSIGNMENT RULES" / "KEEP EXPORTS" / "Google Keep note.txt"
                source = app.CardPipelineApp._google_keep_cache_source(
                    object(),
                    {
                        "kind": "google_keep",
                        "url": "https://keep.google.com/u/0/#NOTE/abc123",
                        "path": str(stale_path),
                        "name": "Google Keep note",
                    },
                    output_dir,
                )
            finally:
                app.ASSIGNMENT_CONFIG_PATH = old_config

            self.assertIsNotNone(source)
            self.assertEqual(Path(source["path"]).parent, output_dir)

    def test_sync_google_keep_notes_opens_saved_sources(self) -> None:
        class Status:
            value = ""

            def set(self, value: str) -> None:
                self.value = value

        class Dummy:
            def __init__(self) -> None:
                self.assignment_config_status = Status()
                self.refreshed = False
                self.opened: list[str] = []

            def _refresh_keep_source_registry(self) -> None:
                self.refreshed = True

            def _saved_google_keep_sources(self) -> list[dict[str, object]]:
                return [
                    {"url": "https://keep.google.com/u/0/#NOTE/abc123", "path": "rules.txt"},
                    {"url": "https://keep.google.com/u/0/#NOTE/def456", "path": "payouts.txt"},
                ]

            def _open_google_keep_source_url(self, url: str) -> bool:
                self.opened.append(url)
                return True

        dummy = Dummy()
        app.CardPipelineApp.sync_google_keep_notes(dummy)

        self.assertTrue(dummy.refreshed)
        self.assertEqual(
            dummy.opened,
            [
                "https://keep.google.com/u/0/#NOTE/abc123",
                "https://keep.google.com/u/0/#NOTE/def456",
            ],
        )
        self.assertIn("Opened 2 Google Keep notes", dummy.assignment_config_status.value)

    def test_bridge_replaces_matching_google_keep_cache(self) -> None:
        with TemporaryDirectory() as tmp:
            cache_path = Path(tmp) / "rules.txt"
            bridge = app.BridgeState()
            bridge.register_keep_note_sources(
                [
                    {
                        "url": "https://keep.google.com/u/0/#NOTE/abc123",
                        "path": str(cache_path),
                        "name": "Rules",
                    }
                ]
            )

            result = bridge.post_google_keep_note(
                {
                    "url": "https://keep.google.com/u/0/#NOTE/abc123",
                    "title": "Rules",
                    "text": "Football $20-$200 95%",
                }
            )

            self.assertTrue(result["ok"])
            self.assertEqual(cache_path.read_text(encoding="utf-8").strip(), "Football $20-$200 95%")

    def test_bridge_poll_advertises_google_keep_sources(self) -> None:
        bridge = app.BridgeState()
        bridge.register_keep_note_sources(
            [
                {
                    "url": "https://keep.google.com/u/0/#NOTE/abc123",
                    "path": "rules.txt",
                    "name": "Rules",
                }
            ]
        )

        result = bridge.extension_poll({"extensionVersion": app.EXPECTED_CARDLADDER_EXTENSION_VERSION})

        self.assertEqual(result["keepNoteSources"][0]["url"], "https://keep.google.com/u/0/#NOTE/abc123")
        self.assertEqual(result["keepNoteSources"][0]["path"], "rules.txt")

    def test_google_keep_hash_note_url_matches_notes_url(self) -> None:
        self.assertTrue(
            bridge_server.keep_urls_match(
                "https://keep.google.com/u/0/#NOTE/abc123",
                "https://keep.google.com/u/0/notes/abc123",
            )
        )

    def test_raw_google_sheet_url_uses_authenticated_reader(self) -> None:
        sheet_url = "https://docs.google.com/spreadsheets/d/private-sheet-id/edit#gid=0"
        with TemporaryDirectory() as tmp:
            with patch.object(assignment_engine, "read_google_sheet_text", return_value="Arena Club rules") as reader:
                text = assignment_engine.read_source_text(sheet_url, Path(tmp), interactive_google=True)

        self.assertEqual(text, "Arena Club rules")
        reader.assert_called_once_with(sheet_url, interactive=True)


class DiagnosticsTests(unittest.TestCase):
    def test_setup_doctor_reports_pipeline_and_google_paths(self) -> None:
        with TemporaryDirectory() as tmp:
            root = Path(tmp)
            for name in ("WORKING SHEETS", "INCOMING SHEETS", "RECEIVED SHEETS", "ASSIGNMENT RULES", "COMPANY SHEETS"):
                (root / name).mkdir(parents=True)

            rows = lucas_diagnostics.setup_doctor_results(
                root,
                {"extensionVersion": "0.1.5", "expectedExtensionVersion": "0.1.5"},
                "Mac",
            )
            by_name = {row["name"]: row for row in rows}

            self.assertEqual(by_name["Shared pipeline folders"]["status"], "OK")
            self.assertEqual(by_name["Card Ladder helper version"]["status"], "OK")
            self.assertIn("lucas_google_sheets_token.json", by_name["Sheets OAuth token"]["detail"])
            self.assertIn("LUCAS Mac", lucas_diagnostics.lucas_version_label("Mac"))


class AssignmentEngineTests(unittest.TestCase):
    def test_kemba_walker_title_infers_basketball_without_short_name_false_matches(self) -> None:
        title = "2019 Panini Contenders Optic Uniformity 21 Kemba Walker PSA 10"
        matches = assignment_engine.find_known_player_sports(title)

        self.assertTrue(any(match["key"] == "kemba walker" and match["sport"] == "basketball" for match in matches))
        self.assertFalse(any(match["sport"] in {"disney", "marvel", "star wars"} for match in matches))
        self.assertEqual(assignment_engine.infer_sport(title, "Kemba Walker"), "basketball")

    def test_shintaro_fujinami_title_infers_baseball_not_wwe(self) -> None:
        title = "2022 Bowman Npb 82 Shintaro Fujinami Chrome-Refractor PSA 10"
        matches = assignment_engine.find_known_player_sports(title)

        self.assertTrue(any(match["key"] == "shintaro fujinami" and match["sport"] == "baseball" for match in matches))
        self.assertFalse(any(match["key"] == "tatsumi fujinami" and match["sport"] == "wwe" for match in matches))
        self.assertEqual(assignment_engine.infer_sport(title, "Shintaro Fujinami"), "baseball")

    def test_bowman_chrome_prospect_title_infers_baseball_without_known_player(self) -> None:
        parsed = assignment_engine.parse_card_for_matching(
            "2017 Bowman Chrome Prospect Autographs Gold Shimmer Refractors #CPACF Clint Frazier BGS 9.5"
        )

        self.assertEqual(parsed["sport"], "baseball")

    def test_rule_category_aliases_remember_common_shorthand(self) -> None:
        self.assertEqual(assignment_engine.infer_sport("B-Ball $17-50 PSA 10"), "basketball")
        self.assertEqual(assignment_engine.infer_sport("bball $50-250 SGC 9"), "basketball")
        self.assertEqual(assignment_engine.infer_sport("1 Piece $15-50 PSA 10"), "one piece")
        self.assertEqual(assignment_engine.infer_sport("Poke $10-100 CGC 10"), "pokemon")

    def test_keep_rule_parser_handles_cy_shorthand_ranges_and_no_blocks(self) -> None:
        rules = assignment_engine.parse_rules(
            "\n".join(
                [
                    "1 Piece $1.75-2.5k None 5 4+",
                    "- NO 7-10k CGC",
                    "- DO NOT BUY ANY MARIO OR LUIGI",
                ]
            )
        )

        self.assertEqual((rules.ranges[0].min_price, rules.ranges[0].max_price), (1750.0, 2500.0))
        self.assertTrue(any(rule.block and rule.min_price == 7000.0 and rule.max_price == 10000.0 for rule in rules.blocks))
        self.assertTrue(any(rule.block and "MARIO OR LUIGI" in rule.matcher for rule in rules.blocks))

    def test_sports_titles_do_not_match_single_word_entertainment_characters(self) -> None:
        examples = {
            "2023 Topps Chrome Platinum Anniversary Autographs Yd Yusniel Diaz PSA 9": ("Yusniel Diaz", "baseball"),
            "2021 Bowman Chrome Futurist Jd Jasson Dominguez Aqua Refractor PSA 10": ("Jasson Dominguez", "baseball"),
            "2023 Bowman Chrome Prospect Autographs Cparo Ricardo Olivar Speckle Refractor PSA 10": ("Ricardo Olivar", "baseball"),
            "1995 Zenith 111 Chipper Jones PSA 9": ("Chipper Jones", "baseball"),
        }

        for title, expected in examples.items():
            with self.subTest(title=title):
                parsed = assignment_engine.parse_card_for_matching(title)
                self.assertEqual((parsed["playerName"], parsed["sport"]), expected)

    def test_world_series_title_infers_baseball_without_false_character_match(self) -> None:
        parsed = assignment_engine.parse_card_for_matching("1990 Score #702 World Series Game 3 PSA 9")

        self.assertEqual(parsed["sport"], "baseball")
        self.assertEqual(parsed["playerName"], "")
        self.assertNotEqual(parsed["sport"], "disney")

    def test_tim_tebow_title_infers_football(self) -> None:
        parsed = assignment_engine.parse_card_for_matching("2010 Donruss Rated Rookies 95 Tim Tebow PSA 10")

        self.assertEqual(parsed["sport"], "football")
        self.assertEqual(parsed["playerName"], "Tim Tebow")

    def test_player_sport_overrides_teach_unknown_players(self) -> None:
        with TemporaryDirectory() as tmp:
            path = Path(tmp) / "overrides.json"
            path.write_text(
                json.dumps({"players": {"Test Override Player": {"sport": "basketball", "displayName": "Test Override Player"}}}),
                encoding="utf-8",
            )

            loaded = assignment_engine.load_player_sport_overrides(path)

        self.assertEqual(loaded, 1)
        self.assertEqual(assignment_engine.infer_sport("2024 Panini Test Override Player PSA 10", "Test Override Player"), "basketball")

    def test_recommendation_chooses_highest_payout_among_accepted_companies(self) -> None:
        row = WorkbookRow(
            excel_row=2,
            cert_number="1",
            grader="PSA",
            card_title="2019 Panini Prizm Stephen Curry Silver PSA 10",
            card_ladder_comps_average=100,
            card_ladder_value=150,
        )
        engine = assignment_engine.AssignmentEngine(
            [
                assignment_engine.AssignmentCompany(
                    "Lower",
                    assignment_engine.CompanyRules(ranges=[assignment_engine.AssignmentRule("basketball", 10, 500)]),
                    [assignment_engine.PayoutTier(10, 500, 0.9, "NBA")],
                ),
                assignment_engine.AssignmentCompany(
                    "Higher",
                    assignment_engine.CompanyRules(ranges=[assignment_engine.AssignmentRule("basketball", 10, 500)]),
                    [assignment_engine.PayoutTier(10, 500, 0.95, "NBA")],
                ),
                assignment_engine.AssignmentCompany(
                    "Rejected",
                    assignment_engine.CompanyRules(ranges=[assignment_engine.AssignmentRule("baseball", 10, 500)]),
                    [assignment_engine.PayoutTier(10, 500, 1.0, "MLB")],
                ),
            ]
        )

        recommendation = engine.recommend(row)
        decisions = {decision.company: decision for decision in engine.evaluate(row)}

        self.assertEqual(recommendation.company, "Higher")
        self.assertEqual(recommendation.payout, 95)
        self.assertTrue(decisions["Lower"].accepted)
        self.assertFalse(decisions["Rejected"].accepted)

    def test_unlicensed_rule_matches_known_unlicensed_patterns(self) -> None:
        rules = assignment_engine.CompanyRules(ranges=[assignment_engine.AssignmentRule("unlicensed", 0, 1000)])
        examples = [
            "2021 Panini Prizm Mike Trout Red PSA 10",
            "2020 Panini Contenders Draft Picks Stephen Curry PSA 10",
            "2023 Leaf Metal Tom Brady Autograph PSA 10",
        ]

        for title in examples:
            with self.subTest(title=title):
                self.assertTrue(assignment_engine.company_accepts(rules, title, 100, "PSA"))

        self.assertFalse(assignment_engine.company_accepts(rules, "2019 Panini Prizm Stephen Curry Silver PSA 10", 100, "PSA"))

    def test_grade_bound_rules_reject_rows_with_missing_grade(self) -> None:
        rules = assignment_engine.CompanyRules(
            accept_all=True,
            grade_rules={"psa": assignment_engine.GradeRule(allowed=True, min_grade=10)},
        )

        self.assertTrue(assignment_engine.company_accepts(rules, "2019 Panini Prizm Stephen Curry Silver PSA 10", 100, "PSA"))
        self.assertFalse(assignment_engine.company_accepts(rules, "2019 Panini Prizm Stephen Curry Silver", 100, "PSA"))

    def test_person_payout_policy_locks_companies_and_overrides_rates(self) -> None:
        row = WorkbookRow(
            excel_row=2,
            cert_number="1",
            grader="PSA",
            card_title="2019 Panini Prizm Stephen Curry Silver PSA 10",
            card_ladder_comps_average=100,
        )
        policies = assignment_engine.parse_person_policies(
            "Person,Company,Min,Max,Rate\n"
            "Lucas,Arena Club,0,,95%\n"
            "Mikey,Arena Club,0,,90%\n"
        )
        engine = assignment_engine.AssignmentEngine(
            [
                assignment_engine.AssignmentCompany(
                    "Arena Club",
                    assignment_engine.CompanyRules(ranges=[assignment_engine.AssignmentRule("basketball", 10, 500)]),
                    [assignment_engine.PayoutTier(10, 500, 0.8, "NBA")],
                ),
                assignment_engine.AssignmentCompany(
                    "Other Buyer",
                    assignment_engine.CompanyRules(ranges=[assignment_engine.AssignmentRule("basketball", 10, 500)]),
                    [assignment_engine.PayoutTier(10, 500, 1.0, "NBA")],
                ),
            ],
            person_policies=policies,
        )

        lucas = engine.recommend(row, person="Lucas")
        mikey = engine.recommend(row, person="Mikey")
        unlisted = engine.recommend(row, person="Unlisted")
        lucas_decisions = {decision.company: decision for decision in engine.evaluate(row, person="Lucas")}

        self.assertEqual(lucas.company, "Arena Club")
        self.assertEqual(lucas.payout, 95)
        self.assertEqual(mikey.company, "Arena Club")
        self.assertEqual(mikey.payout, 90)
        self.assertEqual(unlisted.company, "Other Buyer")
        self.assertEqual(unlisted.payout, 100)
        self.assertFalse(lucas_decisions["Other Buyer"].accepted)
        self.assertIn("not allowed", lucas_decisions["Other Buyer"].reason)

    def test_assignment_engine_loads_person_payout_source_file(self) -> None:
        with TemporaryDirectory() as tmp:
            root = Path(tmp)
            policy_path = root / "person-payouts.csv"
            config_path = root / "assignment_companies.json"
            policy_path.write_text(
                "Person,Company,Rate\nLucas,Arena Club,95%\n",
                encoding="utf-8",
            )
            config_path.write_text(
                json.dumps(
                    {
                        "person_payouts_source": str(policy_path),
                        "companies": [
                            {
                                "name": "Arena Club",
                                "accept_all": True,
                                "rate": "80%",
                            }
                        ],
                    }
                ),
                encoding="utf-8",
            )

            engine = assignment_engine.AssignmentEngine.load(config_path)
            row = WorkbookRow(excel_row=2, cert_number="1", grader="PSA", card_title="Test Card", card_ladder_comps_average=100)

            self.assertEqual(engine.recommend(row, person="Lucas").payout, 95)
            self.assertEqual(engine.recommend(row, person="Other").payout, 80)

    def test_company_can_prefer_card_ladder_value_over_comps(self) -> None:
        row = WorkbookRow(
            excel_row=2,
            cert_number="4",
            grader="PSA",
            card_title="2020 Panini Prizm Patrick Mahomes PSA 10",
            card_ladder_comps_average=100,
            card_ladder_value=150,
        )
        engine = assignment_engine.AssignmentEngine(
            [
                assignment_engine.AssignmentCompany(
                    "Comps Buyer",
                    assignment_engine.CompanyRules(ranges=[assignment_engine.AssignmentRule("football", 10, 500)]),
                    [assignment_engine.PayoutTier(10, 500, 0.95, "NFL")],
                    value_source="comps",
                ),
                assignment_engine.AssignmentCompany(
                    "CL Buyer",
                    assignment_engine.CompanyRules(ranges=[assignment_engine.AssignmentRule("football", 10, 500)]),
                    [assignment_engine.PayoutTier(10, 500, 0.9, "NFL")],
                    value_source="card_ladder",
                ),
            ]
        )

        recommendation = engine.recommend(row)
        decisions = {decision.company: decision for decision in engine.evaluate(row)}

        self.assertEqual(decisions["Comps Buyer"].source_value, 100)
        self.assertEqual(decisions["Comps Buyer"].payout, 95)
        self.assertEqual(decisions["CL Buyer"].source_value, 150)
        self.assertEqual(decisions["CL Buyer"].payout, 135)
        self.assertEqual(recommendation.company, "CL Buyer")
        self.assertEqual(recommendation.payout, 135)

    def test_manual_row_category_drives_assignment_matching(self) -> None:
        row = WorkbookRow(
            excel_row=2,
            cert_number="1",
            grader="BGS",
            category="baseball",
            card_title="Generic Prospect Auto BGS 9.5",
            card_ladder_comps_average=100,
        )
        engine = assignment_engine.AssignmentEngine(
            [
                assignment_engine.AssignmentCompany(
                    "Baseball Buyer",
                    assignment_engine.CompanyRules(ranges=[assignment_engine.AssignmentRule("baseball", 10, 500)]),
                    [assignment_engine.PayoutTier(10, 500, 0.9, "baseball")],
                )
            ]
        )

        recommendation = engine.recommend(row)

        self.assertEqual(recommendation.company, "Baseball Buyer")
        self.assertEqual(recommendation.payout, 90)

    def test_card_ladder_value_source_rejects_company_when_cl_missing(self) -> None:
        row = WorkbookRow(
            excel_row=2,
            cert_number="5",
            grader="PSA",
            card_title="2020 Panini Prizm Patrick Mahomes PSA 10",
            card_ladder_comps_average=100,
            card_ladder_value=None,
        )
        engine = assignment_engine.AssignmentEngine(
            [
                assignment_engine.AssignmentCompany(
                    "Comps Buyer",
                    assignment_engine.CompanyRules(ranges=[assignment_engine.AssignmentRule("football", 10, 500)]),
                    [assignment_engine.PayoutTier(10, 500, 0.9, "NFL")],
                    value_source="comps",
                ),
                assignment_engine.AssignmentCompany(
                    "CL Required Buyer",
                    assignment_engine.CompanyRules(ranges=[assignment_engine.AssignmentRule("football", 10, 500)]),
                    [assignment_engine.PayoutTier(10, 500, 1.0, "NFL")],
                    value_source="card_ladder",
                ),
            ]
        )

        recommendation = engine.recommend(row)
        decisions = {decision.company: decision for decision in engine.evaluate(row)}

        self.assertEqual(recommendation.company, "Comps Buyer")
        self.assertEqual(recommendation.payout, 90)
        self.assertFalse(decisions["CL Required Buyer"].accepted)
        self.assertIsNone(decisions["CL Required Buyer"].source_value)
        self.assertIn("missing", decisions["CL Required Buyer"].reason)

    def test_company_can_require_cy_estimate_value(self) -> None:
        row = WorkbookRow(
            excel_row=2,
            cert_number="6",
            grader="CGC",
            card_title="2022 Paradigm Trigger Unown VSTAR CGC 10",
            card_ladder_comps_average=100,
            card_ladder_value=110,
            cy_value=150,
        )
        engine = assignment_engine.AssignmentEngine(
            [
                assignment_engine.AssignmentCompany(
                    "Comps Buyer",
                    assignment_engine.CompanyRules(accept_all=True),
                    [assignment_engine.PayoutTier(10, 500, 0.9)],
                    value_source="comps",
                ),
                assignment_engine.AssignmentCompany(
                    "CY Buyer",
                    assignment_engine.CompanyRules(accept_all=True),
                    [assignment_engine.PayoutTier(10, 500, 0.85)],
                    value_source="cy_estimate",
                ),
            ]
        )

        recommendation = engine.recommend(row)
        decisions = {decision.company: decision for decision in engine.evaluate(row)}

        self.assertEqual(decisions["Comps Buyer"].source_value, 100)
        self.assertEqual(decisions["CY Buyer"].source_value, 150)
        self.assertEqual(recommendation.company, "CY Buyer")
        self.assertEqual(recommendation.payout, 127.5)

    def test_default_company_value_source_falls_back_to_cy_after_comps_and_card_ladder(self) -> None:
        row = WorkbookRow(
            excel_row=2,
            cert_number="8",
            grader="CGC",
            card_title="2022 Paradigm Trigger Unown VSTAR CGC 10",
            card_ladder_comps_average=None,
            card_ladder_value=None,
            cy_value=150,
        )
        engine = assignment_engine.AssignmentEngine(
            [
                assignment_engine.AssignmentCompany(
                    "Default Buyer",
                    assignment_engine.CompanyRules(accept_all=True),
                    [assignment_engine.PayoutTier(10, 500, 0.8)],
                    value_source="comps",
                )
            ]
        )

        recommendation = engine.recommend(row)
        decisions = {decision.company: decision for decision in engine.evaluate(row)}

        self.assertEqual(decisions["Default Buyer"].source_value, 150)
        self.assertEqual(recommendation.company, "Default Buyer")
        self.assertEqual(recommendation.payout, 120)

    def test_courtyard_value_source_alias_uses_cy_estimate(self) -> None:
        company = assignment_engine.load_company(
            {"name": "CY Alias Buyer", "accept_all": True, "rate": "80%", "value_source": "CourtYard"},
            Path("."),
        )
        row = WorkbookRow(
            excel_row=2,
            cert_number="9",
            grader="CGC",
            card_title="2022 Paradigm Trigger Unown VSTAR CGC 10",
            card_ladder_comps_average=100,
            card_ladder_value=110,
            cy_value=150,
        )
        engine = assignment_engine.AssignmentEngine([company])

        recommendation = engine.recommend(row)

        self.assertEqual(company.value_source, "cy_estimate")
        self.assertEqual(recommendation.company, "CY Alias Buyer")
        self.assertEqual(recommendation.payout, 120)

    def test_cy_estimate_value_source_rejects_company_when_cy_missing(self) -> None:
        row = WorkbookRow(
            excel_row=2,
            cert_number="7",
            grader="CGC",
            card_title="2022 Paradigm Trigger Unown VSTAR CGC 10",
            card_ladder_comps_average=100,
            cy_value=None,
        )
        engine = assignment_engine.AssignmentEngine(
            [
                assignment_engine.AssignmentCompany(
                    "CY Required Buyer",
                    assignment_engine.CompanyRules(accept_all=True),
                    [assignment_engine.PayoutTier(10, 500, 1.0)],
                    value_source="cy_estimate",
                )
            ]
        )

        recommendation = engine.recommend(row)
        decisions = {decision.company: decision for decision in engine.evaluate(row)}

        self.assertEqual(recommendation.company, "")
        self.assertFalse(decisions["CY Required Buyer"].accepted)
        self.assertIsNone(decisions["CY Required Buyer"].source_value)
        self.assertIn("missing", decisions["CY Required Buyer"].reason)

    def test_goat_payout_category_uses_payout_range_not_rule_goat_range(self) -> None:
        row = WorkbookRow(
            excel_row=2,
            cert_number="2",
            grader="PSA",
            card_title="2019 Panini Mosaic Stephen Curry Green Mosaic PSA 10",
            card_ladder_comps_average=85.61,
        )
        rules = assignment_engine.CompanyRules(
            ranges=[assignment_engine.AssignmentRule("basketball", 10, 500)],
            goat_players={"stephen curry"},
            goat_ranges=[assignment_engine.AssignmentRule("Stephen Curry", 100, 7500)],
        )
        engine = assignment_engine.AssignmentEngine(
            [
                assignment_engine.AssignmentCompany(
                    "Goat Buyer",
                    rules,
                    [assignment_engine.PayoutTier(50, 99, 0.95, "GOATS")],
                )
            ]
        )

        recommendation = engine.recommend(row)

        self.assertEqual(recommendation.company, "Goat Buyer")
        self.assertEqual(recommendation.payout, 81.33)

    def test_date_weighted_uses_newest_comp_when_two_newest_are_more_than_seven_days_apart(self) -> None:
        newest = datetime.now() - timedelta(days=2)
        older = newest - timedelta(days=11)
        comps = [
            {"date_sold": (older - timedelta(days=30)).strftime("%b %d, %Y"), "price": "$8.00", "title": "Test Card PSA 10"},
            {"date_sold": newest.strftime("%b %d, %Y"), "price": "$25.00", "title": "Test Card PSA 10"},
            {"date_sold": older.strftime("%b %d, %Y"), "price": "$10.00", "title": "Test Card PSA 10"},
            {"date_sold": (older - timedelta(days=45)).strftime("%b %d, %Y"), "price": "$2.00", "title": "Test Card PSA 10"},
        ]

        self.assertEqual(app.comp_price(comps, app.COMP_STRATEGY_STALE_NEWEST), 25.0)

    def test_date_weighted_averages_when_two_newest_are_within_seven_days(self) -> None:
        newest = datetime.now() - timedelta(days=2)
        second = newest - timedelta(days=4)
        comps = [
            {"date_sold": (second - timedelta(days=30)).strftime("%b %d, %Y"), "price": "$8.00", "title": "Test Card PSA 10"},
            {"date_sold": newest.strftime("%b %d, %Y"), "price": "$25.00", "title": "Test Card PSA 10"},
            {"date_sold": second.strftime("%b %d, %Y"), "price": "$10.00", "title": "Test Card PSA 10"},
            {"date_sold": (second - timedelta(days=45)).strftime("%b %d, %Y"), "price": "$2.00", "title": "Test Card PSA 10"},
        ]

        self.assertEqual(app.comp_price(comps, app.COMP_STRATEGY_STALE_NEWEST), 11.25)

    def test_date_weighted_uses_newest_comp_when_newest_sale_is_stale(self) -> None:
        newest = datetime.now() - timedelta(days=8)
        second = newest - timedelta(days=4)
        comps = [
            {"date_sold": newest.strftime("%b %d, %Y"), "price": "$25.00", "title": "Test Card PSA 10"},
            {"date_sold": second.strftime("%b %d, %Y"), "price": "$10.00", "title": "Test Card PSA 10"},
            {"date_sold": (second - timedelta(days=30)).strftime("%b %d, %Y"), "price": "$8.00", "title": "Test Card PSA 10"},
        ]

        self.assertEqual(app.comp_price(comps, app.COMP_STRATEGY_STALE_NEWEST), 25.0)

    def test_date_weighted_dedupes_same_sale_before_comparing_newest_dates(self) -> None:
        comps = [
            {"date_sold": "Apr 5, 2026", "price": "$11.50", "source": "EBAY", "title": "ZANDGEMPORIUM Pokemon Magnezone Stormfront Holo Rare #5 PSA 7 $12.02"},
            {"date_sold": "Apr 5, 2026", "price": "$12.02", "source": "EBAY", "title": "ZANDGEMPORIUM Pokemon Magnezone Stormfront Holo Rare #5 PSA 7"},
            {"date_sold": "Aug 24, 2025", "price": "$10.50", "source": "EBAY", "title": "ALMAR ENTERPRISES 2008 POKEMON DIAMOND & PEARL STORMFRONT MAGNEZONE LV. 44 #5/100 RARE HOLO PSA 7"},
            {"date_sold": "Oct 2, 2022", "price": "$11.50", "source": "EBAY", "title": "Pokemon Magnezone D&P Stormfront Holo Rare #5 PSA 7 -454"},
        ]

        self.assertEqual(app.comp_price(comps, app.COMP_STRATEGY_STALE_NEWEST), 12.02)

    def test_date_weighted_prefers_clean_duplicate_when_title_contains_conflicting_price(self) -> None:
        comps = [
            {
                "date_sold": "Jun 1, 2025",
                "price": "$493.63",
                "sale_type": "Best Offer",
                "source": "EBAY",
                "title": "POKECARDS-2-TRUST 2022-23 Donruss Chet Holmgren RC #202 - Yellow Holo Laser /25 PSA 9 Rookie Pop 3- $130.50",
            },
            {
                "date_sold": "Jun 1, 2025",
                "price": "$130.50",
                "sale_type": "Best Offer",
                "source": "EBAY",
                "title": "POKECARDS-2-TRUST 2022-23 Donruss Chet Holmgren RC #202 - Yellow Holo Laser /25 PSA 9 Rookie Pop 3",
            },
            {
                "date_sold": "May 4, 2025",
                "price": "$141.25",
                "sale_type": "Auction",
                "source": "EBAY",
                "title": "POKECARDS-2-TRUST 2022-23 Donruss Chet Holmgren RC #202 - Yellow Holo Laser /25 PSA 9 Rookie Pop 3",
            },
        ]

        self.assertEqual(app.comp_price(comps, app.COMP_STRATEGY_STALE_NEWEST), 130.5)

    def test_date_weighted_prefers_best_same_day_stale_newest_comp(self) -> None:
        comps = [
            {
                "source": "BECKETT",
                "title": ", Autograph Grade: 10, Profile: 2014 Panini Flawless Greats Patches Autographs Gold #22 Joe Montana (Pop 2) $2,100.00",
                "date_sold": "Sep 4, 2025",
                "sale_type": "Auction",
                "price": "$1,424.00",
            },
            {
                "source": "ALT",
                "title": "(CONFIRMED PAID) 2014 Panini Flawless Greats Patch Autograph Gold Joe Montana #22 BGS 9 Auto 10 /10",
                "date_sold": "Sep 4, 2025",
                "sale_type": "Auction",
                "price": "$2,760.00",
            },
            {
                "source": "EBAY",
                "title": "225 BREAKERS 2014 Panini Flawless Joe Montana Greats Patch On Card AUTO Gold /10 BGS 9",
                "date_sold": "Jun 18, 2025",
                "sale_type": "Auction",
                "price": "$1,424.00",
            },
        ]

        self.assertEqual(app.comp_price(comps, app.COMP_STRATEGY_STALE_NEWEST), 2760.0)

    def test_cardladder_result_rejects_obvious_wrong_card_comp_rows(self) -> None:
        bridge = app.BridgeState()
        row = WorkbookRow(
            excel_row=2,
            cert_number="99505674",
            grader="PSA",
            card_title="2022 Panini Donruss 202 Chet Holmgren Yellow Holo Laser PSA 9",
        )
        result = {
            "excelRow": 2,
            "certNumber": "99505674",
            "grader": "PSA",
            "status": "ok",
            "value": 73,
            "ocr": {
                "profileTitle": "2022 Panini Donruss 202 Chet Holmgren Yellow Holo Laser",
                "profileGrader": "PSA",
                "profileGrade": "9",
                "resultCount": 2,
                "comps": [
                    {
                        "source": "EBAY",
                        "title": "2022 Panini Donruss Chet Holmgren Yellow Holo Laser #202 PSA 9",
                        "date_sold": "Jun 1, 2026",
                        "sale_type": "Auction",
                        "price": "$73.00",
                    },
                    {
                        "source": "EBAY",
                        "title": "2022 Panini Prizm Chet Holmgren Gold Rookie #266 PSA 10",
                        "date_sold": "Jun 2, 2026",
                        "sale_type": "Auction",
                        "price": "$314.75",
                    },
                ],
            },
        }

        bridge._apply_cardladder_result_to_row(row, result)

        self.assertEqual(row.card_ladder_comps_average, 73.0)
        self.assertIn("Yellow Holo Laser", row.card_ladder_comps)
        self.assertNotIn("Prizm", row.card_ladder_comps)
        self.assertIn("Rejected 1 likely wrong-card", row.notes)

    def test_date_weighted_normalizes_equivalent_date_formats_before_dedupe(self) -> None:
        comps = [
            {"date_sold": "Sept. 4, 2025", "price": "$90.00", "source": "EBAY", "sale_type": "Auction", "title": "Test Card PSA 10 $120.00"},
            {"date_sold": "Sep 4, 2025", "price": "$120.00", "source": "EBAY", "sale_type": "Auction", "title": "Test Card PSA 10"},
            {"date_sold": "8/1/25", "price": "$50.00", "source": "EBAY", "sale_type": "Auction", "title": "Older Test Card PSA 10"},
        ]

        self.assertEqual(app.comp_price(comps, app.COMP_STRATEGY_STALE_NEWEST), 120.0)

    def test_date_weighted_parses_date_embedded_in_ocr_label_text(self) -> None:
        comps = [
            {"date_sold": "Date Sold: Jun 1, 2025 Type Auction", "price": "$25.00", "title": "Test Card PSA 10"},
            {"date_sold": "5/20/25", "price": "$10.00", "title": "Test Card PSA 10"},
        ]

        self.assertEqual(app.comp_price(comps, app.COMP_STRATEGY_STALE_NEWEST), 25.0)

    def test_date_weighted_excludes_undated_comps_when_dated_comps_are_available(self) -> None:
        newest = datetime.now() - timedelta(days=2)
        second = newest - timedelta(days=4)
        comps = [
            {"date_sold": newest.strftime("%b %d, %Y"), "price": "$25.00", "title": "Test Card PSA 10"},
            {"date_sold": second.strftime("%b %d, %Y"), "price": "$15.00", "title": "Test Card PSA 10"},
            {"date_sold": "May 1", "price": "$500.00", "title": "Test Card PSA 10 missing year"},
        ]

        self.assertEqual(app.comp_price(comps, app.COMP_STRATEGY_STALE_NEWEST), 20.0)

    def test_accepted_company_without_matching_payout_cannot_win(self) -> None:
        row = WorkbookRow(
            excel_row=2,
            cert_number="3",
            grader="PSA",
            card_title="2022 Panini Prizm Patrick Mahomes PSA 10",
            card_ladder_comps_average=80,
        )
        engine = assignment_engine.AssignmentEngine(
            [
                assignment_engine.AssignmentCompany(
                    "No Payout Match",
                    assignment_engine.CompanyRules(ranges=[assignment_engine.AssignmentRule("football", 10, 500)]),
                    [assignment_engine.PayoutTier(100, 500, 1.0, "NFL")],
                ),
                assignment_engine.AssignmentCompany(
                    "Valid Payout",
                    assignment_engine.CompanyRules(ranges=[assignment_engine.AssignmentRule("football", 10, 500)]),
                    [assignment_engine.PayoutTier(10, 99, 0.9, "NFL")],
                ),
            ]
        )

        recommendation = engine.recommend(row)
        decisions = {decision.company: decision for decision in engine.evaluate(row)}

        self.assertEqual(recommendation.company, "Valid Payout")
        self.assertIsNone(decisions["No Payout Match"].payout)
        self.assertIn("no payout tier", decisions["No Payout Match"].reason)

    def test_dnb_over_50_section_only_blocks_cards_above_50(self) -> None:
        values = [
            ["", "", "", "Do Not Buy these Players", "", "", "", "Do Not Buy these Players over $50"],
            ["", "Do not buy", "", "Basketball", "Football", "Baseball", "", "Basketball", "Football", "Baseball"],
            ["", "", "", "", "", "", "", "", "Jayden Daniels", ""],
        ]
        rules_text = "\n".join(["football $10-$500", *assignment_engine.synthesize_do_not_buy_rules(values)])
        rules = assignment_engine.parse_rules(rules_text)
        engine = assignment_engine.AssignmentEngine([
            assignment_engine.AssignmentCompany("Arena", rules, [assignment_engine.PayoutTier(10, 500, 0.9)])
        ])
        low_value = WorkbookRow(
            excel_row=2,
            cert_number="",
            grader="PSA",
            category="football",
            card_title="2024 Panini Absolute Jayden Daniels PSA 9",
            card_ladder_comps_average=50,
        )
        high_value = WorkbookRow(
            excel_row=3,
            cert_number="",
            grader="PSA",
            category="football",
            card_title="2024 Panini Absolute Jayden Daniels PSA 9",
            card_ladder_comps_average=51,
        )

        self.assertIn("block: Jayden Daniels over 50", rules_text)
        self.assertTrue(engine.evaluate(low_value)[0].accepted)
        self.assertFalse(engine.evaluate(high_value)[0].accepted)


class AppSharedWorkflowLogicTests(unittest.TestCase):
    def test_bridge_only_hands_commands_to_expected_extension_version(self) -> None:
        bridge = app.BridgeState()
        row = WorkbookRow(excel_row=2, cert_number="123", grader="CGC", card_title="Test Card CGC 9")
        bridge.set_rows([row])
        bridge.start_all_comps(requery_all=True)

        stale = bridge.extension_poll({"extensionVersion": "older-helper"})
        self.assertIsNone(stale["command"])

        current = bridge.extension_poll({"extensionVersion": app.EXPECTED_CARDLADDER_EXTENSION_VERSION})
        self.assertIsNotNone(current["command"])
        self.assertEqual(current["command"]["type"], "RUN_ALL_COMPS")

    def test_unassigned_player_is_recorded_for_unmatched_valued_row(self) -> None:
        class Dummy:
            _load_unassigned_players = app.CardPipelineApp._load_unassigned_players
            _save_unassigned_players = app.CardPipelineApp._save_unassigned_players
            _record_unassigned_player = app.CardPipelineApp._record_unassigned_player
            _append_unique_sample = app.CardPipelineApp._append_unique_sample
            _guess_unassigned_player = app.CardPipelineApp._guess_unassigned_player
            _strip_card_variant_tail = app.CardPipelineApp._strip_card_variant_tail

        with TemporaryDirectory() as tmp:
            old_pipeline = app.CARD_PIPELINE_DIR
            old_unassigned = app.UNASSIGNED_PLAYERS_PATH
            app.CARD_PIPELINE_DIR = Path(tmp)
            app.UNASSIGNED_PLAYERS_PATH = Path(tmp) / "unassigned_players.json"
            dummy = Dummy()
            dummy.lucas_identity = {"display_name": "Tester", "machine": "Test"}
            row = WorkbookRow(
                excel_row=2,
                cert_number="",
                grader="PSA",
                card_title="2024 Panini Prizm 12 Jalen Madeup PSA 10",
                card_ladder_comps_average=25,
            )
            try:
                dummy._record_unassigned_player(row)
                entries = json.loads(app.UNASSIGNED_PLAYERS_PATH.read_text(encoding="utf-8"))["entries"]
                self.assertIn("jalen madeup", entries)
                self.assertEqual(entries["jalen madeup"]["player"], "Jalen Madeup")
            finally:
                app.CARD_PIPELINE_DIR = old_pipeline
                app.UNASSIGNED_PLAYERS_PATH = old_unassigned

    def test_unassigned_player_records_failed_assignment_even_with_wrong_partial_match(self) -> None:
        class Dummy:
            _load_unassigned_players = app.CardPipelineApp._load_unassigned_players
            _save_unassigned_players = app.CardPipelineApp._save_unassigned_players
            _record_unassigned_player = app.CardPipelineApp._record_unassigned_player
            _append_unique_sample = app.CardPipelineApp._append_unique_sample
            _guess_unassigned_player = app.CardPipelineApp._guess_unassigned_player
            _strip_card_variant_tail = app.CardPipelineApp._strip_card_variant_tail

        with TemporaryDirectory() as tmp:
            old_pipeline = app.CARD_PIPELINE_DIR
            old_unassigned = app.UNASSIGNED_PLAYERS_PATH
            app.CARD_PIPELINE_DIR = Path(tmp)
            app.UNASSIGNED_PLAYERS_PATH = Path(tmp) / "unassigned_players.json"
            dummy = Dummy()
            dummy.lucas_identity = {"display_name": "Tester", "machine": "Test"}
            row = WorkbookRow(
                excel_row=2,
                cert_number="",
                grader="PSA",
                card_title="2022 Bowman Npb 82 Shintaro Fujinami Chrome-Refractor PSA 10",
                card_ladder_comps_average=25,
            )
            try:
                dummy._record_unassigned_player(row)
                entries = json.loads(app.UNASSIGNED_PLAYERS_PATH.read_text(encoding="utf-8"))["entries"]
                self.assertIn("shintaro fujinami", entries)
                self.assertEqual(entries["shintaro fujinami"]["player"], "Shintaro Fujinami")
            finally:
                app.CARD_PIPELINE_DIR = old_pipeline
                app.UNASSIGNED_PLAYERS_PATH = old_unassigned

    def test_failed_assignment_is_labeled_nobody_takes_and_recorded(self) -> None:
        class Dummy:
            _apply_recommendations = app.CardPipelineApp._apply_recommendations
            _assignment_person_for_row = app.CardPipelineApp._assignment_person_for_row
            _home_sheet_key = app.CardPipelineApp._home_sheet_key
            _split_home_sheet_key = app.CardPipelineApp._split_home_sheet_key
            _load_unassigned_players = app.CardPipelineApp._load_unassigned_players
            _save_unassigned_players = app.CardPipelineApp._save_unassigned_players
            _record_unassigned_player = app.CardPipelineApp._record_unassigned_player
            _append_unique_sample = app.CardPipelineApp._append_unique_sample
            _guess_unassigned_player = app.CardPipelineApp._guess_unassigned_player
            _strip_card_variant_tail = app.CardPipelineApp._strip_card_variant_tail

        with TemporaryDirectory() as tmp:
            old_pipeline = app.CARD_PIPELINE_DIR
            old_unassigned = app.UNASSIGNED_PLAYERS_PATH
            app.CARD_PIPELINE_DIR = Path(tmp)
            app.UNASSIGNED_PLAYERS_PATH = Path(tmp) / "unassigned_players.json"
            row = WorkbookRow(
                excel_row=2,
                cert_number="",
                grader="PSA",
                card_title="2023 Topps Chrome Platinum Anniversary Autographs Yd Yusniel Diaz PSA 9",
                card_ladder_comps_average=20,
            )
            dummy = Dummy()
            dummy.lucas_identity = {"display_name": "Tester", "machine": "Test"}
            dummy.state = types.SimpleNamespace(rows=[row])
            dummy.review_rows = []
            dummy.assignment_engine = assignment_engine.AssignmentEngine(
                [
                    assignment_engine.AssignmentCompany(
                        "Basketball Buyer",
                        assignment_engine.CompanyRules(ranges=[assignment_engine.AssignmentRule("basketball", 10, 500)]),
                        [assignment_engine.PayoutTier(10, 500, 0.9, "NBA")],
                    )
                ]
            )
            try:
                dummy._apply_recommendations()
                self.assertEqual(row.best_company, app.NO_COMPANY_TAKES_LABEL)
                self.assertIsNone(row.estimated_payout)
                entries = json.loads(app.UNASSIGNED_PLAYERS_PATH.read_text(encoding="utf-8"))["entries"]
                self.assertIn("yusniel diaz", entries)
            finally:
                app.CARD_PIPELINE_DIR = old_pipeline
                app.UNASSIGNED_PLAYERS_PATH = old_unassigned

    def test_assignment_person_for_row_uses_sheet_marker(self) -> None:
        class Dummy:
            _assignment_person_for_row = app.CardPipelineApp._assignment_person_for_row
            _home_sheet_key = app.CardPipelineApp._home_sheet_key
            _split_home_sheet_key = app.CardPipelineApp._split_home_sheet_key

        dummy = Dummy()
        dummy.comp_sheet_sources = {}
        dummy.review_sheet_sources = {}
        dummy.intake_sheet_sources = {}
        dummy.row_sources = {2: "Lot A.xlsx"}
        dummy.review_sources = {}
        dummy.intake_sources = {}
        dummy.selected_working_sheet = types.SimpleNamespace(get=lambda: "")
        dummy.home_sheet_markers = {"Working|Lot A.xlsx": {"assigned_person": "Lucas"}}
        row = WorkbookRow(excel_row=2, cert_number="1", grader="PSA", card_title="Test", card_ladder_comps_average=100)

        self.assertEqual(dummy._assignment_person_for_row(row), "Lucas")

    def test_scoped_assignment_results_leave_unreturned_rows_unchanged(self) -> None:
        class FieldVar:
            def __init__(self):
                self.value = ""

            def set(self, value):
                self.value = value

        class Dummy:
            _apply_assignment_recommendation_results = app.CardPipelineApp._apply_assignment_recommendation_results

            def __init__(self):
                self.assignment_recommendation_job = 7
                self.assignment_recommendation_running = True
                self.assignment_progress_value = FieldVar()
                self.review_status = FieldVar()
                self.status_var = FieldVar()
                self.comp_output_saved = True
                self.review_rows = []
                self.state = types.SimpleNamespace(
                    rows=[
                        WorkbookRow(1, "111", "Changed Card", "PSA", best_company="Old A", estimated_payout=1),
                        WorkbookRow(2, "222", "Untouched Card", "PSA", best_company="Keep Me", estimated_payout=22),
                    ]
                )
                self.refreshed_comp = False

            def _record_unassigned_players(self, rows):
                self.unassigned_rows = rows

            def _refresh_comp_table(self, schedule_recommendations=False):
                self.refreshed_comp = True

            def _refresh_table(self, schedule_recommendations=False):
                raise AssertionError("scoped comp-only results should not refresh every table")

        dummy = Dummy()
        target = dummy.state.rows[0]

        dummy._apply_assignment_recommendation_results(
            {"job_id": 7, "total": 1, "results": [(id(target), "Arena Club", 95.0)]}
        )

        self.assertEqual(dummy.state.rows[0].best_company, "Arena Club")
        self.assertEqual(dummy.state.rows[0].estimated_payout, 95.0)
        self.assertEqual(dummy.state.rows[1].best_company, "Keep Me")
        self.assertEqual(dummy.state.rows[1].estimated_payout, 22)
        self.assertTrue(dummy.refreshed_comp)

    def test_nobody_takes_is_not_an_assignable_company(self) -> None:
        class Dummy:
            _row_has_assignable_company = app.CardPipelineApp._row_has_assignable_company

        dummy = Dummy()
        nobody_row = WorkbookRow(excel_row=2, cert_number="", grader="PSA", card_title="Test Card", best_company=app.NO_COMPANY_TAKES_LABEL)
        buyer_row = WorkbookRow(excel_row=3, cert_number="", grader="PSA", card_title="Test Card", best_company="Arena Club")

        self.assertFalse(dummy._row_has_assignable_company(nobody_row))
        self.assertTrue(dummy._row_has_assignable_company(buyer_row))

    def test_unassigned_auto_categorize_uses_web_text_and_saves_override(self) -> None:
        class Dummy:
            _load_unassigned_players = app.CardPipelineApp._load_unassigned_players
            _save_unassigned_players = app.CardPipelineApp._save_unassigned_players
            _auto_categorize_unassigned_players = app.CardPipelineApp._auto_categorize_unassigned_players
            _search_unassigned_player_category = app.CardPipelineApp._search_unassigned_player_category
            _category_research_text = app.CardPipelineApp._category_research_text
            _infer_category_from_web_text = app.CardPipelineApp._infer_category_from_web_text
            _write_player_category_override = app.CardPipelineApp._write_player_category_override
            def _wikipedia_search_text(self, _player: str) -> str:
                return "Example Player is an NBA basketball guard who appears on Panini basketball cards."
            def _wikidata_search_text(self, _player: str) -> str:
                return ""
            def _duckduckgo_search_text(self, _query: str) -> str:
                return ""

        with TemporaryDirectory() as tmp:
            old_pipeline = app.CARD_PIPELINE_DIR
            old_unassigned = app.UNASSIGNED_PLAYERS_PATH
            old_overrides = app.PLAYER_OVERRIDES_PATH
            app.CARD_PIPELINE_DIR = Path(tmp)
            app.UNASSIGNED_PLAYERS_PATH = Path(tmp) / "unassigned_players.json"
            app.PLAYER_OVERRIDES_PATH = Path(tmp) / "assignment_player_overrides.json"
            dummy = Dummy()
            dummy.lucas_identity = {"display_name": "Tester", "machine": "Test"}
            try:
                dummy._save_unassigned_players({
                    "example player": {
                        "player": "Example Player",
                        "last_title": "2024 Panini Prizm 12 Example Player PSA 10",
                    }
                })
                result = dummy._auto_categorize_unassigned_players()
                overrides = json.loads(app.PLAYER_OVERRIDES_PATH.read_text(encoding="utf-8"))
                remaining = json.loads(app.UNASSIGNED_PLAYERS_PATH.read_text(encoding="utf-8"))["entries"]
                self.assertEqual(result["resolved"], 1)
                self.assertEqual(overrides["players"]["Example Player"]["sport"], "basketball")
                self.assertEqual(remaining, {})
            finally:
                app.CARD_PIPELINE_DIR = old_pipeline
                app.UNASSIGNED_PLAYERS_PATH = old_unassigned
                app.PLAYER_OVERRIDES_PATH = old_overrides

    def test_sheet_marker_save_merges_latest_and_honors_tombstones(self) -> None:
        class MarkerDummy:
            _load_sheet_markers = app.CardPipelineApp._load_sheet_markers
            _save_sheet_markers = app.CardPipelineApp._save_sheet_markers
            _delete_sheet_marker = app.CardPipelineApp._delete_sheet_marker

        with TemporaryDirectory() as tmp:
            old_pipeline = app.CARD_PIPELINE_DIR
            old_markers = app.SHEET_MARKERS_PATH
            app.CARD_PIPELINE_DIR = Path(tmp)
            app.SHEET_MARKERS_PATH = Path(tmp) / "sheet_markers.json"
            app.SHEET_MARKERS_PATH.write_text(
                json.dumps(
                    {
                        "Incoming|A.xlsx": {"assigned_person": "A"},
                        "Incoming|B.xlsx": {"assigned_person": "B"},
                    }
                ),
                encoding="utf-8",
            )
            dummy = MarkerDummy()
            dummy.lucas_identity = {"display_name": "Tester", "machine": "Test"}
            dummy.home_sheet_markers = {"Incoming|C.xlsx": {"assigned_person": "C"}}
            dummy.deleted_sheet_marker_keys = {"Incoming|A.xlsx"}
            try:
                dummy._save_sheet_markers()
                saved = json.loads(app.SHEET_MARKERS_PATH.read_text(encoding="utf-8"))
                self.assertNotIn("Incoming|A.xlsx", saved)
                self.assertEqual(saved["Incoming|B.xlsx"]["assigned_person"], "B")
                self.assertEqual(saved["Incoming|C.xlsx"]["assigned_person"], "C")
            finally:
                app.CARD_PIPELINE_DIR = old_pipeline
                app.SHEET_MARKERS_PATH = old_markers

    def test_create_seller_assignment_pays_only_after_receive(self) -> None:
        class SellerSheetDummy:
            _home_sheet_key = app.CardPipelineApp._home_sheet_key
            _split_home_sheet_key = app.CardPipelineApp._split_home_sheet_key
            _load_sheet_markers = app.CardPipelineApp._load_sheet_markers
            _save_sheet_markers = app.CardPipelineApp._save_sheet_markers
            _delete_sheet_marker = app.CardPipelineApp._delete_sheet_marker
            _marker_for_stage = app.CardPipelineApp._marker_for_stage
            _sheet_path_for_stage = app.CardPipelineApp._sheet_path_for_stage
            _move_home_sheet_to_stage = app.CardPipelineApp._move_home_sheet_to_stage
            _assign_sheet_to_seller = app.CardPipelineApp._assign_sheet_to_seller
            _active_payout_balance = app.CardPipelineApp._active_payout_balance
            _payout_sheet_status = app.CardPipelineApp._payout_sheet_status
            _payout_sheet_items = app.CardPipelineApp._payout_sheet_items
            _sheet_marker_is_seller_payout = app.CardPipelineApp._sheet_marker_is_seller_payout
            _source_sheet_is_seller_payout = app.CardPipelineApp._source_sheet_is_seller_payout
            _sheet_marker_is_seller_payout = app.CardPipelineApp._sheet_marker_is_seller_payout
            _source_sheet_is_seller_payout = app.CardPipelineApp._source_sheet_is_seller_payout
            _realized_profit_totals_by_person_sheet = app.CardPipelineApp._realized_profit_totals_by_person_sheet
            _realized_profit_groups_by_person_sheet = app.CardPipelineApp._realized_profit_groups_by_person_sheet
            _normalize_profit_record = app.CardPipelineApp._normalize_profit_record
            _money_value = app.CardPipelineApp._money_value
            _person_for_profit_record = app.CardPipelineApp._person_for_profit_record
            _enrich_profit_records_with_people = app.CardPipelineApp._enrich_profit_records_with_people
            def _seller_terms_seller_names(self):
                return {"john seller"}
            def _load_profit_ledger(self):
                return []

        with TemporaryDirectory() as tmp:
            root = Path(tmp)
            incoming_dir = root / "INCOMING SHEETS"
            working_dir = root / "WORKING SHEETS"
            received_dir = root / "RECEIVED SHEETS"
            working_dir.mkdir(parents=True)
            incoming_dir.mkdir()
            received_dir.mkdir()
            (working_dir / "Lot A.xlsx").write_text("placeholder", encoding="utf-8")

            old_pipeline = app.CARD_PIPELINE_DIR
            old_incoming = app.INCOMING_SHEETS_DIR
            old_working = app.WORKING_SHEETS_DIR
            old_received = app.RECEIVED_SHEETS_DIR
            old_markers = app.SHEET_MARKERS_PATH
            app.CARD_PIPELINE_DIR = root
            app.INCOMING_SHEETS_DIR = incoming_dir
            app.WORKING_SHEETS_DIR = working_dir
            app.RECEIVED_SHEETS_DIR = received_dir
            app.SHEET_MARKERS_PATH = root / "sheet_markers.json"
            dummy = SellerSheetDummy()
            dummy.lucas_identity = {"display_name": "Tester", "machine": "Test"}
            dummy.home_sheet_markers = {}
            dummy.deleted_sheet_marker_keys = set()
            dummy.home_sheet_paths = {"Incoming": {}, "Working": {"Lot A.xlsx": working_dir / "Lot A.xlsx"}, "Received": {}}
            dummy.received_sheet_paths = {}
            dummy.home_sheet_summaries = {}
            try:
                self.assertTrue(dummy._assign_sheet_to_seller("Working", "Lot A.xlsx", "John Seller", "Arena Club", {"rate": 0.9}))
                self.assertEqual(dummy.home_sheet_markers["Working|Lot A.xlsx"]["assigned_person"], "John Seller")
                self.assertTrue(dummy.home_sheet_markers["Working|Lot A.xlsx"]["seller_terms_applied"])
                self.assertEqual(dummy.home_sheet_markers["Working|Lot A.xlsx"]["seller_sheet_type"], "Arena Club")
                self.assertEqual(dummy.home_sheet_markers["Working|Lot A.xlsx"]["seller_rate"], 0.9)

                moved_key, cleanup = dummy._move_home_sheet_to_stage("Working|Lot A.xlsx", "Incoming")
                self.assertEqual(moved_key, "Incoming|Lot A.xlsx")
                self.assertEqual(cleanup, {})
                self.assertEqual(dummy.home_sheet_markers[moved_key]["assigned_person"], "John Seller")
                dummy.home_sheet_paths = {"Incoming": {"Lot A.xlsx": incoming_dir / "Lot A.xlsx"}, "Working": {}, "Received": {}}
                dummy.home_sheet_summaries = {moved_key: {"row_count": 2, "received_count": 0, "purchase_total": 123.45, "estimated_payout_total": 200.0}}

                payout_items = dummy._payout_sheet_items()
                self.assertEqual(payout_items, [])

                received_key, cleanup = dummy._move_home_sheet_to_stage(moved_key, "Received")
                self.assertEqual(received_key, "Received|Lot A.xlsx")
                self.assertEqual(cleanup, {})
                self.assertEqual(dummy.home_sheet_markers[received_key]["assigned_person"], "John Seller")
                dummy.home_sheet_paths = {"Incoming": {}, "Working": {}, "Received": {"Lot A.xlsx": received_dir / "Lot A.xlsx"}}
                dummy.home_sheet_summaries = {received_key: {"row_count": 2, "received_count": 2, "purchase_total": 123.45, "estimated_payout_total": 200.0}}

                payout_items = dummy._payout_sheet_items()
                self.assertEqual(len(payout_items), 1)
                self.assertEqual(payout_items[0]["person"], "John Seller")
                self.assertEqual(payout_items[0]["stage"], "Received")
                self.assertEqual(payout_items[0]["purchase_total"], 123.45)
                self.assertEqual(payout_items[0]["payout_balance"], 123.45)
            finally:
                app.CARD_PIPELINE_DIR = old_pipeline
                app.INCOMING_SHEETS_DIR = old_incoming
                app.WORKING_SHEETS_DIR = old_working
                app.RECEIVED_SHEETS_DIR = old_received
                app.SHEET_MARKERS_PATH = old_markers

    def test_active_payout_balance_uses_seller_or_team_member_rule(self) -> None:
        class PayoutDummy:
            _active_payout_balance = app.CardPipelineApp._active_payout_balance

        dummy = PayoutDummy()
        sellers = {"john seller"}

        self.assertEqual(dummy._active_payout_balance("John Seller", 80.0, 150.0, sellers), (80.0, "Seller purchase total"))
        self.assertEqual(dummy._active_payout_balance("Kevin Hambone", 80.0, 150.0, sellers), (0.0, "Team half sold profit"))
        self.assertEqual(dummy._active_payout_balance("Kevin Hambone", 80.0, 150.0, sellers, realized_profit_total=70.0), (35.0, "Team half sold profit"))
        self.assertEqual(dummy._active_payout_balance("Kevin Hambone", 100.0, 80.0, sellers, realized_profit_total=-20.0), (0.0, "Team half sold profit"))

    def test_sheet_marker_controls_seller_payout_classification(self) -> None:
        class PayoutDummy:
            _home_sheet_key = app.CardPipelineApp._home_sheet_key
            _sheet_marker_is_seller_payout = app.CardPipelineApp._sheet_marker_is_seller_payout
            _source_sheet_is_seller_payout = app.CardPipelineApp._source_sheet_is_seller_payout

            def __init__(self):
                self.home_sheet_markers = {
                    "Received|Seller Lot.xlsx": {"assigned_person": "John Seller", "seller_terms_applied": True},
                    "Received|Team Lot.xlsx": {"assigned_person": "John Seller"},
                }

            def _seller_terms_seller_names(self):
                return {"john seller"}

        dummy = PayoutDummy()
        self.assertTrue(dummy._source_sheet_is_seller_payout("Seller Lot.xlsx", "John Seller"))
        self.assertFalse(dummy._source_sheet_is_seller_payout("Team Lot.xlsx", "John Seller"))
        self.assertTrue(dummy._source_sheet_is_seller_payout("Legacy Missing Marker.xlsx", "John Seller"))

    def test_known_people_includes_seller_terms_people(self) -> None:
        class PeopleDummy:
            _known_people = app.CardPipelineApp._known_people

            def _known_assigned_people(self):
                return ["Kevin Hambone"]

            def _load_seller_terms(self):
                return [{"seller": "John Seller", "sheet_type": "Arena Club", "rate": 0.9}]

            def _load_profit_ledger(self):
                return []

            def _load_inventory_ledger(self):
                return []

        self.assertEqual(PeopleDummy()._known_people(), ["John Seller", "Kevin Hambone"])

    def test_save_working_sheet_requires_valid_network_seller_terms(self) -> None:
        class Var:
            def __init__(self, value=""):
                self.value = value

            def get(self):
                return self.value

        class SaveDummy:
            save_working_sheet = app.CardPipelineApp.save_working_sheet
            _network_mode_enabled = app.CardPipelineApp._network_mode_enabled
            _money_value = app.CardPipelineApp._money_value

            def __init__(self):
                self.intake_rows = [WorkbookRow(excel_row=2, cert_number="1", grader="PSA", card_title="Test", existing_value=10)]
                self.working_sheet_title = Var("Network Lot")
                self.create_network_mode_var = Var(True)
                self.seller_terms_seller_var = Var("John Seller")
                self.seller_terms_sheet_type_var = Var("Arena Club")
                self.applied_terms = False

            def _seller_terms_match(self, seller, sheet_type):
                return None

            def _seller_terms_company_price(self, row, company_name, rate=None, deduction=None):
                return 90.0

            def apply_create_seller_terms(self, show_status=True):
                self.applied_terms = True
                return 1

        dummy = SaveDummy()
        with patch.object(app.messagebox, "showinfo") as showinfo:
            dummy.save_working_sheet()
        self.assertTrue(showinfo.called)
        self.assertFalse(dummy.applied_terms)

    def test_save_working_sheet_no_seller_rows_explains_assignment_reason(self) -> None:
        class Var:
            def __init__(self, value=""):
                self.value = value

            def get(self):
                return self.value

        class SaveDummy:
            save_working_sheet = app.CardPipelineApp.save_working_sheet
            _network_mode_enabled = app.CardPipelineApp._network_mode_enabled
            _money_value = app.CardPipelineApp._money_value
            _seller_terms_company_decision = app.CardPipelineApp._seller_terms_company_decision
            _seller_terms_company_price = app.CardPipelineApp._seller_terms_company_price
            _seller_terms_no_match_details = app.CardPipelineApp._seller_terms_no_match_details
            _seller_terms_no_match_message = app.CardPipelineApp._seller_terms_no_match_message

            def __init__(self):
                self.intake_rows = [WorkbookRow(excel_row=2, cert_number="137915162", grader="PSA", card_title="Test Card PSA 10", existing_value=10)]
                self.working_sheet_title = Var("Network Lot")
                self.create_network_mode_var = Var(True)
                self.seller_terms_seller_var = Var("John Seller")
                self.seller_terms_sheet_type_var = Var("Arena Club")
                self.assignment_engine = types.SimpleNamespace(
                    evaluate=lambda row: [
                        types.SimpleNamespace(
                            company="Arena Club",
                            accepted=False,
                            payout=None,
                            source_value=None,
                            reason="missing comp/card ladder value",
                        )
                    ]
                )
                self.applied_terms = False

            def _seller_terms_match(self, seller, sheet_type):
                return {"seller": seller, "sheet_type": sheet_type, "deduction": 0.1}

            def apply_create_seller_terms(self, show_status=True):
                self.applied_terms = True
                return 0

        dummy = SaveDummy()
        with patch.object(app.messagebox, "showinfo") as showinfo:
            dummy.save_working_sheet()
        self.assertTrue(showinfo.called)
        self.assertEqual(showinfo.call_args.args[0], "No seller payout rows")
        self.assertIn("seller_terms.csv row", showinfo.call_args.args[1])
        self.assertIn("137915162", showinfo.call_args.args[1])
        self.assertIn("missing comp/card ladder value", showinfo.call_args.args[1])
        self.assertFalse(dummy.applied_terms)

    def test_seller_terms_health_reports_duplicates_and_company_issues(self) -> None:
        with TemporaryDirectory() as tmp:
            path = Path(tmp) / "seller_terms.csv"
            path.write_text(
                "Seller,Sheet Type,Seller Rate,Deduction\n"
                "John,Arena Club,90%,\n"
                "John,Arena Club,85%,\n"
                "Mary,Fanatics,,5%\n"
                "No Company,Unknown,90%,\n"
                "Bad Rate,Arena Club,nope,\n",
                encoding="utf-8",
            )
            lines = assignment_config_ui.seller_terms_health_lines(
                path,
                [
                    {"name": "Arena Club", "active": True},
                    {"name": "Fanatics", "active": False},
                ],
            )
        text = "\n".join(lines)
        self.assertIn("3 valid row(s)", text)
        self.assertIn("duplicate Seller/Sheet Type", text)
        self.assertIn("inactive assignment company", text)
        self.assertIn("is not an assignment company", text)
        self.assertIn("invalid Seller Rate", text)

    def test_paid_received_sheets_archive_after_two_weeks_only_when_paid(self) -> None:
        class ArchiveDummy:
            _home_sheet_key = app.CardPipelineApp._home_sheet_key
            _split_home_sheet_key = app.CardPipelineApp._split_home_sheet_key
            _load_sheet_markers = app.CardPipelineApp._load_sheet_markers
            _save_sheet_markers = app.CardPipelineApp._save_sheet_markers
            _delete_sheet_marker = app.CardPipelineApp._delete_sheet_marker
            _archive_eligible_received_sheets = app.CardPipelineApp._archive_eligible_received_sheets
            _received_sheet_is_archive_age = app.CardPipelineApp._received_sheet_is_archive_age
            _received_at_for_archive = app.CardPipelineApp._received_at_for_archive
            _unique_archive_path = app.CardPipelineApp._unique_archive_path

        with TemporaryDirectory() as tmp:
            root = Path(tmp)
            received_dir = root / "RECEIVED SHEETS"
            archive_dir = root / "ARCHIVED SHEETS"
            received_dir.mkdir()
            paid_old = received_dir / "paid-old.xlsx"
            unpaid_old = received_dir / "unpaid-old.xlsx"
            paid_recent = received_dir / "paid-recent.xlsx"
            for path in (paid_old, unpaid_old, paid_recent):
                path.write_text("placeholder", encoding="utf-8")

            old_pipeline = app.CARD_PIPELINE_DIR
            old_received = app.RECEIVED_SHEETS_DIR
            old_archive = app.ARCHIVED_SHEETS_DIR
            old_markers = app.SHEET_MARKERS_PATH
            app.CARD_PIPELINE_DIR = root
            app.RECEIVED_SHEETS_DIR = received_dir
            app.ARCHIVED_SHEETS_DIR = archive_dir
            app.SHEET_MARKERS_PATH = root / "sheet_markers.json"
            old_date = "2026-05-01T12:00:00"
            recent_date = datetime.now().isoformat(timespec="seconds")
            app.SHEET_MARKERS_PATH.write_text(
                json.dumps(
                    {
                        "Received|paid-old.xlsx": {"paid": True, "all_received": True, "received_at": old_date},
                        "Received|unpaid-old.xlsx": {"paid": False, "all_received": True, "received_at": old_date},
                        "Received|paid-recent.xlsx": {"paid": True, "all_received": True, "received_at": recent_date},
                    }
                ),
                encoding="utf-8",
            )
            dummy = ArchiveDummy()
            dummy.lucas_identity = {"display_name": "Tester", "machine": "Test"}
            dummy.home_sheet_markers = dummy._load_sheet_markers()
            dummy.deleted_sheet_marker_keys = set()
            try:
                archived = dummy._archive_eligible_received_sheets()
                saved = json.loads(app.SHEET_MARKERS_PATH.read_text(encoding="utf-8"))

                self.assertEqual(archived, ["paid-old.xlsx"])
                self.assertFalse(paid_old.exists())
                self.assertTrue((archive_dir / "paid-old.xlsx").exists())
                self.assertTrue(unpaid_old.exists())
                self.assertTrue(paid_recent.exists())
                self.assertNotIn("Received|paid-old.xlsx", saved)
                self.assertIn("Archived|paid-old.xlsx", saved)
                self.assertIn("Received|unpaid-old.xlsx", saved)
                self.assertIn("Received|paid-recent.xlsx", saved)
            finally:
                app.CARD_PIPELINE_DIR = old_pipeline
                app.RECEIVED_SHEETS_DIR = old_received
                app.ARCHIVED_SHEETS_DIR = old_archive
                app.SHEET_MARKERS_PATH = old_markers

    def test_team_payout_uses_sold_profit_not_unsold_estimated_profit(self) -> None:
        class PayoutDummy:
            _home_sheet_key = app.CardPipelineApp._home_sheet_key
            _split_home_sheet_key = app.CardPipelineApp._split_home_sheet_key
            _money_value = app.CardPipelineApp._money_value
            _profit_record_key = app.CardPipelineApp._profit_record_key
            _normalize_profit_record = app.CardPipelineApp._normalize_profit_record
            _person_for_profit_record = app.CardPipelineApp._person_for_profit_record
            _enrich_profit_records_with_people = app.CardPipelineApp._enrich_profit_records_with_people
            _sold_payout_key = app.CardPipelineApp._sold_payout_key
            _realized_profit_totals_by_person_sheet = app.CardPipelineApp._realized_profit_totals_by_person_sheet
            _realized_profit_groups_by_person_sheet = app.CardPipelineApp._realized_profit_groups_by_person_sheet
            _active_payout_balance = app.CardPipelineApp._active_payout_balance
            _payout_sheet_status = app.CardPipelineApp._payout_sheet_status
            _payout_sheet_items = app.CardPipelineApp._payout_sheet_items
            _sheet_marker_is_seller_payout = app.CardPipelineApp._sheet_marker_is_seller_payout
            _source_sheet_is_seller_payout = app.CardPipelineApp._source_sheet_is_seller_payout

            def __init__(self):
                self.home_sheet_paths = {"Incoming": {"Lot A.xlsx": Path("Lot A.xlsx")}, "Received": {}}
                self.home_sheet_markers = {"Incoming|Lot A.xlsx": {"assigned_person": "Kevin Hambone"}}
                self.home_sheet_summaries = {
                    "Incoming|Lot A.xlsx": {
                        "row_count": 2,
                        "received_count": 0,
                        "purchase_total": 80.0,
                        "estimated_payout_total": 150.0,
                    }
                }
                self.ledger = []

            def _seller_terms_seller_names(self):
                return set()

            def _load_profit_ledger(self):
                return self.ledger

        dummy = PayoutDummy()
        payout_items = dummy._payout_sheet_items()
        self.assertEqual(payout_items, [])

        dummy.ledger = [
            {
                "assigned_person": "Kevin Hambone",
                "source_sheet": "Lot A.xlsx",
                "purchase_price": 80.0,
                "sale_price": 150.0,
                "company": "Arena Club",
                "cert_number": "123",
                "date_added": "2026-06-19",
            }
        ]
        payout_items = dummy._payout_sheet_items()
        self.assertEqual(len(payout_items), 1)
        self.assertEqual(payout_items[0]["stage"], "Sold")
        self.assertEqual(payout_items[0]["name"], "Lot A.xlsx")
        self.assertEqual(payout_items[0]["realized_profit_total"], 70.0)
        self.assertEqual(payout_items[0]["payout_balance"], 35.0)

    def test_moving_received_sheet_back_clears_received_profit_and_company_rows(self) -> None:
        class MoveDummy:
            _home_sheet_key = app.CardPipelineApp._home_sheet_key
            _split_home_sheet_key = app.CardPipelineApp._split_home_sheet_key
            _sheet_path_for_stage = app.CardPipelineApp._sheet_path_for_stage
            _delete_sheet_marker = app.CardPipelineApp._delete_sheet_marker
            _marker_for_stage = app.CardPipelineApp._marker_for_stage
            _money_value = app.CardPipelineApp._money_value
            _profit_record_key = app.CardPipelineApp._profit_record_key
            _normalize_profit_record = app.CardPipelineApp._normalize_profit_record
            _load_profit_ledger = app.CardPipelineApp._load_profit_ledger
            _save_profit_ledger = app.CardPipelineApp._save_profit_ledger
            _remove_profit_ledger_rows_for_source = app.CardPipelineApp._remove_profit_ledger_rows_for_source
            _cleanup_sheet_received_side_effects = app.CardPipelineApp._cleanup_sheet_received_side_effects
            _move_home_sheet_to_stage = app.CardPipelineApp._move_home_sheet_to_stage

        with TemporaryDirectory() as tmp:
            root = Path(tmp)
            incoming_dir = root / "INCOMING SHEETS"
            working_dir = root / "WORKING SHEETS"
            received_dir = root / "RECEIVED SHEETS"
            company_dir = root / "COMPANY SHEETS"
            received_dir.mkdir(parents=True)
            working_dir.mkdir()

            source_path = received_dir / "Lot A.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.append(["Certification Number", "Card Description", "RECEIVED"])
            sheet.append(["123", "Test Card", "X"])
            workbook.save(source_path)

            company_path = company_dir / "Arena Club" / "Arena Club.xlsx"
            company_path.parent.mkdir(parents=True)
            company_workbook = Workbook()
            company_sheet = company_workbook.active
            company_sheet.title = "Week of 2026-06-15"
            company_sheet.append(["Date Added", "Source Sheet", "Source", "Certification Number", "Grader", "Card Description", "Purchase Price", "Card Ladder Value", "Comps", "CY Estimate", "CY Confidence", "Best Company", "Estimated Payout", "Status", "Notes"])
            company_sheet.append(["2026-06-17", "Lot A.xlsx", "", "123", "PSA", "Test Card", 40, "", 100, "", "", "Arena Club", 90, "Received", ""])
            company_workbook.save(company_path)

            old_incoming = app.INCOMING_SHEETS_DIR
            old_working = app.WORKING_SHEETS_DIR
            old_received = app.RECEIVED_SHEETS_DIR
            old_company = app.COMPANY_SHEETS_DIR
            old_ledger = app.PROFIT_LEDGER_PATH
            app.INCOMING_SHEETS_DIR = incoming_dir
            app.WORKING_SHEETS_DIR = working_dir
            app.RECEIVED_SHEETS_DIR = received_dir
            app.COMPANY_SHEETS_DIR = company_dir
            app.PROFIT_LEDGER_PATH = root / "profit_ledger.json"
            app.PROFIT_LEDGER_PATH.write_text(
                json.dumps([
                    {
                        "date_added": "2026-06-17",
                        "company": "Arena Club",
                        "weekly_sheet_name": "Arena Club.xlsx:Week of 2026-06-15",
                        "source_sheet": "Lot A.xlsx",
                        "cert_number": "123",
                        "card_title": "Test Card",
                        "purchase_price": 40,
                        "sale_price": 90,
                    }
                ]),
                encoding="utf-8",
            )
            dummy = MoveDummy()
            dummy.home_sheet_paths = {"Incoming": {}, "Working": {}, "Received": {}}
            dummy.received_sheet_paths = {"Lot A.xlsx": source_path}
            dummy.home_sheet_markers = {
                "Received|Lot A.xlsx": {
                    "assigned_person": "Lucas",
                    "paid": True,
                    "all_received": True,
                    "received_at": "2026-06-17T10:00:00",
                }
            }
            dummy.deleted_sheet_marker_keys = set()
            try:
                moved_key, cleanup = dummy._move_home_sheet_to_stage("Received|Lot A.xlsx", "Working")

                self.assertEqual(moved_key, "Working|Lot A.xlsx")
                self.assertTrue((working_dir / "Lot A.xlsx").exists())
                moved_workbook = load_workbook(working_dir / "Lot A.xlsx", data_only=True)
                try:
                    self.assertEqual(moved_workbook.active.cell(2, 3).value, None)
                finally:
                    moved_workbook.close()
                company_rows = read_company_profit_records(company_dir)
                ledger = json.loads(app.PROFIT_LEDGER_PATH.read_text(encoding="utf-8"))
                marker = dummy.home_sheet_markers[moved_key]
                self.assertEqual(company_rows, [])
                self.assertEqual(ledger, [])
                self.assertFalse(marker["paid"])
                self.assertFalse(marker["all_received"])
                self.assertNotIn("received_at", marker)
                self.assertEqual(cleanup["received_rows_cleared"], 1)
                self.assertEqual(cleanup["company_rows_removed"], 1)
                self.assertEqual(cleanup["profit_rows_removed"], 1)
            finally:
                app.INCOMING_SHEETS_DIR = old_incoming
                app.WORKING_SHEETS_DIR = old_working
                app.RECEIVED_SHEETS_DIR = old_received
                app.COMPANY_SHEETS_DIR = old_company
                app.PROFIT_LEDGER_PATH = old_ledger

    def test_profit_sales_are_deduped_and_delta_is_recorded(self) -> None:
        class ProfitDummy:
            _load_profit_ledger = app.CardPipelineApp._load_profit_ledger
            _save_profit_ledger = app.CardPipelineApp._save_profit_ledger
            _profit_record_key = app.CardPipelineApp._profit_record_key
            _money_value = app.CardPipelineApp._money_value
            _normalize_profit_record = app.CardPipelineApp._normalize_profit_record
            record_profit_sales = app.CardPipelineApp.record_profit_sales
            refresh_profit_tab = lambda self: None

        with TemporaryDirectory() as tmp:
            old_pipeline = app.CARD_PIPELINE_DIR
            old_ledger = app.PROFIT_LEDGER_PATH
            app.CARD_PIPELINE_DIR = Path(tmp)
            app.PROFIT_LEDGER_PATH = Path(tmp) / "profit_ledger.json"
            dummy = ProfitDummy()
            dummy.lucas_identity = {"display_name": "Tester", "machine": "Test"}
            try:
                record = {
                    "date_added": "2026-06-11",
                    "company": "Arena Club",
                    "weekly_sheet_name": "Arena WEEK.xlsx",
                    "source_sheet": "source.xlsx",
                    "cert_number": "123",
                    "card_title": "Test Card",
                    "purchase_price": "$40.00",
                    "sale_price": "$90.00",
                }
                self.assertEqual(dummy.record_profit_sales([record]), 1)
                self.assertEqual(dummy.record_profit_sales([record]), 0)
                ledger = json.loads(app.PROFIT_LEDGER_PATH.read_text(encoding="utf-8"))
                self.assertEqual(len(ledger), 1)
                self.assertEqual(ledger[0]["profit"], 50.0)
                self.assertEqual(ledger[0]["recorded_by"], "Tester")
            finally:
                app.CARD_PIPELINE_DIR = old_pipeline
                app.PROFIT_LEDGER_PATH = old_ledger

    def test_profit_refresh_skips_company_scan_until_deep_sync(self) -> None:
        class ProfitRefreshDummy:
            _load_profit_ledger = app.CardPipelineApp._load_profit_ledger
            _save_profit_ledger = app.CardPipelineApp._save_profit_ledger
            _profit_record_key = app.CardPipelineApp._profit_record_key
            _is_manual_company_profit_backfill = app.CardPipelineApp._is_manual_company_profit_backfill
            _money_value = app.CardPipelineApp._money_value
            _normalize_profit_record = app.CardPipelineApp._normalize_profit_record
            _enrich_profit_records_with_people = app.CardPipelineApp._enrich_profit_records_with_people
            _filtered_profit_records = app.CardPipelineApp._filtered_profit_records
            _profit_record_date = app.CardPipelineApp._profit_record_date
            _profit_period_bounds = app.CardPipelineApp._profit_period_bounds
            _profit_today = lambda self: datetime(2026, 6, 20).date()
            refresh_profit_tab = app.CardPipelineApp.refresh_profit_tab

            def _person_for_profit_record(self, record):
                return str(record.get("assigned_person") or "")

        with TemporaryDirectory() as tmp:
            root = Path(tmp)
            old_pipeline = app.CARD_PIPELINE_DIR
            old_company = app.COMPANY_SHEETS_DIR
            old_ledger = app.PROFIT_LEDGER_PATH
            app.CARD_PIPELINE_DIR = root
            app.COMPANY_SHEETS_DIR = root / "COMPANY SHEETS"
            app.PROFIT_LEDGER_PATH = root / "profit_ledger.json"
            dummy = ProfitRefreshDummy()
            dummy.lucas_identity = {"display_name": "Tester", "machine": "Test"}
            dummy.profit_person_var = types.SimpleNamespace(get=lambda: "")
            dummy.profit_search_var = types.SimpleNamespace(get=lambda: "")
            dummy.profit_period_var = types.SimpleNamespace(get=lambda: "Total")
            company_record = {
                "assigned_person": "Kevin Hambone",
                "date_added": "2026-06-20",
                "company": "Arena Club",
                "weekly_sheet_name": "Arena Club.xlsx:Week of 2026-06-15",
                "source_sheet": "Lot A.xlsx",
                "cert_number": "123",
                "card_title": "Test Card",
                "purchase_price": 40,
                "sale_price": 90,
            }
            try:
                with patch("app.read_company_profit_records", side_effect=AssertionError("normal refresh scanned company sheets")):
                    dummy.refresh_profit_tab()
                self.assertEqual(dummy.profit_rows, [])
                self.assertFalse(app.PROFIT_LEDGER_PATH.exists())

                with patch("app.read_company_profit_records", return_value=[company_record]) as reader:
                    dummy.refresh_profit_tab(deep_sync=True)
                reader.assert_called_once_with(app.COMPANY_SHEETS_DIR)
                ledger = json.loads(app.PROFIT_LEDGER_PATH.read_text(encoding="utf-8"))
                self.assertEqual(len(ledger), 1)
                self.assertEqual(ledger[0]["cert_number"], "123")
                self.assertEqual(ledger[0]["profit"], 50.0)
            finally:
                app.CARD_PIPELINE_DIR = old_pipeline
                app.COMPANY_SHEETS_DIR = old_company
                app.PROFIT_LEDGER_PATH = old_ledger

    def test_home_workbook_summary_cache_reuses_unchanged_files(self) -> None:
        class HomeSummaryDummy:
            _home_summary_cache_key = app.CardPipelineApp._home_summary_cache_key
            _summarize_home_workbook_cached = app.CardPipelineApp._summarize_home_workbook_cached
            _prune_home_summary_cache = app.CardPipelineApp._prune_home_summary_cache

        with TemporaryDirectory() as tmp:
            path = Path(tmp) / "Lot A.xlsx"
            path.write_text("first", encoding="utf-8")
            dummy = HomeSummaryDummy()
            dummy.home_summary_cache = {}
            dummy.home_summary_cache_lock = threading.Lock()
            with patch("app.summarize_workbook", side_effect=[{"name": "first"}, {"name": "second"}]) as summarizer:
                self.assertEqual(dummy._summarize_home_workbook_cached(path), {"name": "first"})
                self.assertEqual(dummy._summarize_home_workbook_cached(path), {"name": "first"})
                self.assertEqual(summarizer.call_count, 1)

                path.write_text("second and larger", encoding="utf-8")
                self.assertEqual(dummy._summarize_home_workbook_cached(path), {"name": "second"})
                self.assertEqual(summarizer.call_count, 2)

                dummy._prune_home_summary_cache([])
                self.assertEqual(dummy.home_summary_cache, {})

    def test_expense_records_deduct_from_person_profit(self) -> None:
        class ExpenseDummy:
            _load_profit_ledger = app.CardPipelineApp._load_profit_ledger
            _save_profit_ledger = app.CardPipelineApp._save_profit_ledger
            _profit_record_key = app.CardPipelineApp._profit_record_key
            _money_value = app.CardPipelineApp._money_value
            _normalize_profit_record = app.CardPipelineApp._normalize_profit_record
            record_profit_sales = app.CardPipelineApp.record_profit_sales
            _profit_record_date = app.CardPipelineApp._profit_record_date
            _profit_today = lambda self: datetime(2026, 6, 19).date()
            _profit_period_bounds = app.CardPipelineApp._profit_period_bounds
            _profit_period_label = app.CardPipelineApp._profit_period_label
            _profit_graph_label = app.CardPipelineApp._profit_graph_label
            _profit_chart_title = app.CardPipelineApp._profit_chart_title
            _filtered_profit_records = app.CardPipelineApp._filtered_profit_records
            _profit_chart_series = app.CardPipelineApp._profit_chart_series
            _expense_related_label = app.CardPipelineApp._expense_related_label
            _expense_link_options = app.CardPipelineApp._expense_link_options
            _delete_profit_expense_records = app.CardPipelineApp._delete_profit_expense_records
            refresh_profit_tab = lambda self: None

        with TemporaryDirectory() as tmp:
            old_pipeline = app.CARD_PIPELINE_DIR
            old_ledger = app.PROFIT_LEDGER_PATH
            app.CARD_PIPELINE_DIR = Path(tmp)
            app.PROFIT_LEDGER_PATH = Path(tmp) / "profit_ledger.json"
            dummy = ExpenseDummy()
            dummy.lucas_identity = {"display_name": "Tester", "machine": "Test"}
            dummy.profit_person_var = types.SimpleNamespace(get=lambda: "Kevin")
            dummy.profit_period_var = types.SimpleNamespace(get=lambda: "YTD")
            dummy.profit_graph_var = types.SimpleNamespace(get=lambda: "Daily Trend")
            try:
                expense = {
                    "record_type": "expense",
                    "expense_id": "expense-1",
                    "assigned_person": "Kevin Hambone",
                    "date_added": "2026-06-18",
                    "expense_type": "Shipping",
                    "expense_amount": 25,
                    "related_type": "Card",
                    "source_sheet": "Lot A.xlsx",
                    "cert_number": "123",
                    "notes": "Shipment label",
                }
                sale = {
                    "assigned_person": "Kevin Hambone",
                    "cert_number": "123",
                    "card_title": "Test Card",
                    "source_sheet": "Lot A.xlsx",
                    "company": "Arena",
                    "purchase_price": 50,
                    "sale_price": 100,
                    "date_added": "2026-06-18",
                }
                self.assertEqual(dummy.record_profit_sales([expense, sale]), 2)
                ledger = [dummy._normalize_profit_record(record) for record in dummy._load_profit_ledger()]
                expense_row = next(record for record in ledger if record.get("record_type") == "expense")
                self.assertEqual(expense_row["profit"], -25)
                self.assertEqual(expense_row["company"], "Expense: Shipping")
                self.assertEqual(expense_row["source_sheet"], "Lot A.xlsx")
                self.assertEqual(expense_row["cert_number"], "123")
                self.assertEqual(dummy._expense_related_label(expense_row), "Lot A.xlsx | 123")
                raw_expense_row = dict(expense_row, source_sheet="Raw Inventory", cert_number="", item_id="RAW-20260624-0001")
                self.assertEqual(dummy._expense_related_label(raw_expense_row), "Raw Inventory | RAW-20260624-0001")
                sheets, cards, lookup = dummy._expense_link_options("Kevin")
                self.assertEqual(sheets, ["Lot A.xlsx"])
                self.assertEqual(len(cards), 1)
                self.assertIn("Lot A.xlsx | Test Card | $100.00", cards)
                self.assertEqual(lookup[cards[0]]["source_sheet"], "Lot A.xlsx")
                self.assertEqual(lookup[cards[0]]["cert_number"], "123")
                filtered = dummy._filtered_profit_records(ledger)
                _days, values = dummy._profit_chart_series(filtered)
                self.assertEqual(sum(values), 25)
                self.assertEqual(dummy._delete_profit_expense_records([expense_row]), 1)
                ledger_after_delete = [dummy._normalize_profit_record(record) for record in dummy._load_profit_ledger()]
                self.assertEqual(len(ledger_after_delete), 1)
                self.assertNotEqual(ledger_after_delete[0].get("record_type"), "expense")
                self.assertEqual(ledger_after_delete[0]["profit"], 50)
                self.assertEqual(dummy._delete_profit_expense_records([ledger_after_delete[0]]), 0)
            finally:
                app.CARD_PIPELINE_DIR = old_pipeline
                app.PROFIT_LEDGER_PATH = old_ledger

    def test_inventory_records_are_deduped_and_reactivated(self) -> None:
        class InventoryDummy:
            _money_value = app.CardPipelineApp._money_value
            _inventory_record_key = app.CardPipelineApp._inventory_record_key
            _normalize_inventory_record = app.CardPipelineApp._normalize_inventory_record
            _load_inventory_ledger = app.CardPipelineApp._load_inventory_ledger
            _save_inventory_ledger = app.CardPipelineApp._save_inventory_ledger
            add_inventory_records = app.CardPipelineApp.add_inventory_records
            _enrich_inventory_record_assignment = lambda self, record: record
            refresh_inventory_tab = lambda self: None

        with TemporaryDirectory() as tmp:
            old_inventory = app.INVENTORY_LEDGER_PATH
            app.INVENTORY_LEDGER_PATH = Path(tmp) / "inventory_ledger.json"
            dummy = InventoryDummy()
            try:
                record = {
                    "assigned_person": "Lucas",
                    "cert_number": "123",
                    "card_title": "Test Card",
                    "source_sheet": "Lot A.xlsx",
                    "purchase_price": 40,
                    "inventory_value": 100,
                    "status": "Refunded",
                }
                self.assertEqual(dummy.add_inventory_records([record]), 1)
                self.assertEqual(dummy.add_inventory_records([{**record, "status": "Active"}]), 0)
                ledger = json.loads(app.INVENTORY_LEDGER_PATH.read_text(encoding="utf-8"))["items"]
                self.assertEqual(len(ledger), 1)
                self.assertEqual(ledger[0]["status"], "Active")
            finally:
                app.INVENTORY_LEDGER_PATH = old_inventory

    def test_inventory_record_from_row_preserves_manual_sport_for_unknown_title(self) -> None:
        class InventoryDummy:
            _money_value = app.CardPipelineApp._money_value
            _inventory_record_key = app.CardPipelineApp._inventory_record_key
            _normalize_inventory_record = app.CardPipelineApp._normalize_inventory_record
            _inventory_sport_from_value = app.CardPipelineApp._inventory_sport_from_value
            _inventory_record_from_row = app.CardPipelineApp._inventory_record_from_row

        row = WorkbookRow(
            excel_row=2,
            cert_number="0010355805",
            grader="BGS",
            card_title="2017 Bowman Chrome Prospect Autographs Gold Shimmer Refractors #CPACF Clint Frazier BGS 9.5",
            category="baseball",
            card_ladder_comps_average=432,
        )

        record = InventoryDummy()._inventory_record_from_row(row, "James Copeland", "JAMES_NASHVILLE_INVENTORY_2.xlsx")

        self.assertEqual(record["sport"], "baseball")

    def test_inventory_record_normalizes_common_sport_aliases(self) -> None:
        class InventoryDummy:
            _money_value = app.CardPipelineApp._money_value
            _inventory_record_key = app.CardPipelineApp._inventory_record_key
            _normalize_inventory_record = app.CardPipelineApp._normalize_inventory_record

        dummy = InventoryDummy()

        self.assertEqual(dummy._normalize_inventory_record({"sport": "b-ball"})["sport"], "basketball")
        self.assertEqual(dummy._normalize_inventory_record({"category": "Poke"})["sport"], "pokemon")

    def test_inventory_filter_searches_cert_and_card_title(self) -> None:
        class FieldVar:
            def __init__(self, value=""):
                self.value = value

            def get(self):
                return self.value

        class InventoryDummy:
            _money_value = app.CardPipelineApp._money_value
            _filtered_inventory_records = app.CardPipelineApp._filtered_inventory_records

        dummy = InventoryDummy()
        dummy.inventory_person_var = FieldVar("")
        dummy.inventory_sport_var = FieldVar("")
        dummy.inventory_min_var = FieldVar("")
        dummy.inventory_max_var = FieldVar("")
        rows = [
            {"status": "Active", "cert_number": "151740304", "card_title": "2024 Prizm Victor Wembanyama Silver PSA 10", "inventory_value": 100},
            {"status": "Active", "cert_number": "222", "card_title": "2019 Panini Mosaic Stephen Curry Green PSA 10", "inventory_value": 90},
            {"status": "Sold", "cert_number": "333", "card_title": "Hidden Sold Card", "inventory_value": 80},
        ]

        dummy.inventory_search_var = FieldVar("174030")
        self.assertEqual([row["cert_number"] for row in dummy._filtered_inventory_records(rows)], ["151740304"])

        dummy.inventory_search_var = FieldVar("curry green")
        self.assertEqual([row["cert_number"] for row in dummy._filtered_inventory_records(rows)], ["222"])

        dummy.inventory_search_var = FieldVar("hidden sold")
        self.assertEqual(dummy._filtered_inventory_records(rows), [])

    def test_table_sorting_handles_money_and_blank_values(self) -> None:
        class SortDummy:
            _money_value = app.CardPipelineApp._money_value
            _expense_related_label = app.CardPipelineApp._expense_related_label
            _record_sort_value = app.CardPipelineApp._record_sort_value
            _sorted_records = app.CardPipelineApp._sorted_records

        dummy = SortDummy()
        inventory_rows = [
            {"cert_number": "1", "purchase_price": "$9.00"},
            {"cert_number": "2", "purchase_price": "$100.00"},
            {"cert_number": "3", "purchase_price": ""},
        ]

        self.assertEqual([row["cert_number"] for row in dummy._sorted_records(inventory_rows, "purchase", False, "inventory")], ["1", "2", "3"])
        self.assertEqual([row["cert_number"] for row in dummy._sorted_records(inventory_rows, "purchase", True, "inventory")], ["2", "1", "3"])

        profit_rows = [
            {"card_title": "A", "profit": "$5.00"},
            {"card_title": "B", "profit": "$15.00"},
            {"card_title": "C", "profit": ""},
        ]
        self.assertEqual([row["card_title"] for row in dummy._sorted_records(profit_rows, "profit", True, "profit")], ["B", "A", "C"])

    def test_inventory_refresh_purges_non_active_rows_from_ledger(self) -> None:
        class InventoryDummy:
            _money_value = app.CardPipelineApp._money_value
            _inventory_record_key = app.CardPipelineApp._inventory_record_key
            _normalize_inventory_record = app.CardPipelineApp._normalize_inventory_record
            _load_inventory_ledger = app.CardPipelineApp._load_inventory_ledger
            _save_inventory_ledger = app.CardPipelineApp._save_inventory_ledger
            _filtered_inventory_records = app.CardPipelineApp._filtered_inventory_records
            refresh_inventory_tab = app.CardPipelineApp.refresh_inventory_tab

        with TemporaryDirectory() as tmp:
            old_inventory = app.INVENTORY_LEDGER_PATH
            app.INVENTORY_LEDGER_PATH = Path(tmp) / "inventory_ledger.json"
            dummy = InventoryDummy()
            try:
                dummy._save_inventory_ledger([
                    dummy._normalize_inventory_record({"assigned_person": "Kevin Hambone", "cert_number": "111", "source_sheet": "A.xlsx", "status": "Active"}),
                    dummy._normalize_inventory_record({"assigned_person": "Kevin Hambone", "cert_number": "222", "source_sheet": "B.xlsx", "status": "Sold"}),
                    dummy._normalize_inventory_record({"assigned_person": "Kevin Hambone", "cert_number": "333", "source_sheet": "C.xlsx", "status": "Company Sheet"}),
                ])

                dummy.refresh_inventory_tab()

                ledger = json.loads(app.INVENTORY_LEDGER_PATH.read_text(encoding="utf-8"))["items"]
                self.assertEqual([row["cert_number"] for row in ledger], ["111"])
            finally:
                app.INVENTORY_LEDGER_PATH = old_inventory

    def test_inventory_table_values_can_be_copied_without_editing(self) -> None:
        class FakeTree:
            def item(self, _row_id, _option):
                return ("2026-06-18", "Kevin Hambone", "Basketball", "12345678", "PSA", "Test Card")

        class FakeStatus:
            def __init__(self):
                self.value = ""

            def set(self, value):
                self.value = value

        class InventoryDummy:
            _inventory_tree_cell_text = app.CardPipelineApp._inventory_tree_cell_text
            _inventory_tree_row_text = app.CardPipelineApp._inventory_tree_row_text
            _copy_inventory_text = app.CardPipelineApp._copy_inventory_text
            copy_inventory_cell_value = app.CardPipelineApp.copy_inventory_cell_value
            copy_inventory_row_values = app.CardPipelineApp.copy_inventory_row_values

            def __init__(self):
                self.inventory_tree = FakeTree()
                self.status_var = FakeStatus()
                self.clipboard = ""

            def clipboard_clear(self):
                self.clipboard = ""

            def clipboard_append(self, text):
                self.clipboard = text

        dummy = InventoryDummy()
        dummy.copy_inventory_cell_value("row-1", "#4")
        self.assertEqual(dummy.clipboard, "12345678")
        self.assertEqual(dummy.status_var.value, "Copied inventory cell.")

        dummy.copy_inventory_row_values("row-1")
        self.assertEqual(dummy.clipboard, "2026-06-18\tKevin Hambone\tBasketball\t12345678\tPSA\tTest Card")
        self.assertEqual(dummy.status_var.value, "Copied inventory row.")

    def test_inventory_table_shows_source_values_separately(self) -> None:
        class FieldVar:
            def __init__(self, value=""):
                self.value = value

            def get(self):
                return self.value

            def set(self, value):
                self.value = value

        class FakeTree:
            def __init__(self):
                self.rows = []

            def get_children(self):
                return []

            def delete(self, *_args):
                return None

            def insert(self, *_args, **kwargs):
                self.rows.append(kwargs["values"])
                return "row-1"

        class InventoryDummy:
            _money_value = app.CardPipelineApp._money_value
            _inventory_record_key = app.CardPipelineApp._inventory_record_key
            _normalize_inventory_record = app.CardPipelineApp._normalize_inventory_record
            _filtered_inventory_records = app.CardPipelineApp._filtered_inventory_records
            refresh_inventory_tab = app.CardPipelineApp.refresh_inventory_tab

            def __init__(self):
                self.inventory_person_var = FieldVar("")
                self.inventory_sport_var = FieldVar("")
                self.inventory_search_var = FieldVar("")
                self.inventory_min_var = FieldVar("")
                self.inventory_max_var = FieldVar("")
                self.inventory_tree = FakeTree()
                self.inventory_metric_var = FieldVar("")
                self.inventory_status_var = FieldVar("")

            def _load_inventory_ledger(self):
                return [
                    self._normalize_inventory_record(
                        {
                            "date_added": "2026-06-18",
                            "assigned_person": "Kevin Hambone",
                            "sport": "football",
                            "cert_number": "141119049",
                            "grader": "PSA",
                            "card_title": "2025 Panini Mosaic Elevate Travis Hunter PSA 10",
                            "purchase_price": 40,
                            "inventory_value": 41,
                            "card_ladder_value": 41,
                            "card_ladder_comps_average": 38.87,
                            "cy_value": "",
                            "cy_confidence": "",
                            "best_company": "FANATICS",
                            "estimated_payout": 38.95,
                            "source_sheet": "SCOTSBORO_HAMBONE_6_16_26.xlsx",
                            "status": "Active",
                        }
                    )
                ]

            def _save_inventory_ledger(self, _rows):
                raise AssertionError("active-only refresh should not rewrite this fixture")

            def _refresh_person_combo_values(self):
                return None

        self.assertNotIn("value", app.INVENTORY_TABLE_COLUMNS)
        self.assertIn("card_ladder", app.INVENTORY_TABLE_COLUMNS)
        self.assertIn("comps", app.INVENTORY_TABLE_COLUMNS)
        self.assertIn("cy_estimate", app.INVENTORY_TABLE_COLUMNS)
        self.assertIn("cy_confidence", app.INVENTORY_TABLE_COLUMNS)

        dummy = InventoryDummy()
        dummy.refresh_inventory_tab()

        row = dummy.inventory_tree.rows[0]
        columns = app.INVENTORY_TABLE_COLUMNS
        self.assertEqual(row[columns.index("card_ladder")], "$41.00")
        self.assertEqual(row[columns.index("comps")], "$38.87")
        self.assertEqual(row[columns.index("cy_estimate")], "")
        self.assertEqual(row[columns.index("company")], "FANATICS")
        self.assertEqual(row[columns.index("payout")], "$38.95")
        self.assertIn("Source Value: $41.00", dummy.inventory_metric_var.value)

    def test_filtered_inventory_refresh_only_enriches_visible_rows(self) -> None:
        class FieldVar:
            def __init__(self, value=""):
                self.value = value

            def get(self):
                return self.value

            def set(self, value):
                self.value = value

        class FakeTree:
            def get_children(self):
                return []

            def delete(self, *_args):
                return None

            def insert(self, *_args, **_kwargs):
                return "row"

        class FakeAssignment:
            def recommend(self, row, person=""):
                return types.SimpleNamespace(company=f"{person} Club", payout=88)

        class InventoryDummy:
            _money_value = app.CardPipelineApp._money_value
            _inventory_record_key = app.CardPipelineApp._inventory_record_key
            _normalize_inventory_record = app.CardPipelineApp._normalize_inventory_record
            _load_inventory_ledger = app.CardPipelineApp._load_inventory_ledger
            _save_inventory_ledger = app.CardPipelineApp._save_inventory_ledger
            _inventory_workbook_row = app.CardPipelineApp._inventory_workbook_row
            _enrich_inventory_record_assignment = app.CardPipelineApp._enrich_inventory_record_assignment
            _filtered_inventory_records = app.CardPipelineApp._filtered_inventory_records
            refresh_inventory_tab = app.CardPipelineApp.refresh_inventory_tab
            update_inventory_payouts = app.CardPipelineApp.update_inventory_payouts
            _refresh_person_combo_values = lambda self: None

        with TemporaryDirectory() as tmp:
            old_inventory = app.INVENTORY_LEDGER_PATH
            app.INVENTORY_LEDGER_PATH = Path(tmp) / "inventory_ledger.json"
            dummy = InventoryDummy()
            dummy.assignment_engine = FakeAssignment()
            dummy.inventory_person_var = FieldVar("Kevin")
            dummy.inventory_sport_var = FieldVar("")
            dummy.inventory_search_var = FieldVar("")
            dummy.inventory_min_var = FieldVar("")
            dummy.inventory_max_var = FieldVar("")
            dummy.inventory_tree = FakeTree()
            dummy.inventory_metric_var = FieldVar("")
            dummy.inventory_status_var = FieldVar("")
            try:
                dummy._save_inventory_ledger([
                    dummy._normalize_inventory_record({"assigned_person": "Kevin Hambone", "cert_number": "1", "source_sheet": "A.xlsx", "card_title": "Kevin Card", "best_company": "Old", "estimated_payout": 1}),
                    dummy._normalize_inventory_record({"assigned_person": "Lucas", "cert_number": "2", "source_sheet": "B.xlsx", "card_title": "Lucas Card", "best_company": "Old", "estimated_payout": 1}),
                ])

                dummy.update_inventory_payouts()

                ledger = json.loads(app.INVENTORY_LEDGER_PATH.read_text(encoding="utf-8"))["items"]
                by_person = {record["assigned_person"]: record for record in ledger}
                self.assertEqual(by_person["Kevin Hambone"]["best_company"], "Kevin Hambone Club")
                self.assertEqual(by_person["Kevin Hambone"]["estimated_payout"], 88)
                self.assertEqual(by_person["Lucas"]["best_company"], "Old")
                self.assertEqual(by_person["Lucas"]["estimated_payout"], 1)
                self.assertEqual(dummy.inventory_status_var.value, "Updated payouts for 1 visible inventory card(s); changed 1.")
            finally:
                app.INVENTORY_LEDGER_PATH = old_inventory

    def test_inventory_record_assignment_enrichment_adds_company_and_payout(self) -> None:
        class FakeAssignment:
            def recommend(self, row, person=""):
                self.last_row = row
                return types.SimpleNamespace(company="Arena Club", payout=95, source_value=150)

        class InventoryDummy:
            _money_value = app.CardPipelineApp._money_value
            _inventory_record_key = app.CardPipelineApp._inventory_record_key
            _normalize_inventory_record = app.CardPipelineApp._normalize_inventory_record
            _inventory_workbook_row = app.CardPipelineApp._inventory_workbook_row
            _enrich_inventory_record_assignment = app.CardPipelineApp._enrich_inventory_record_assignment

        dummy = InventoryDummy()
        dummy.assignment_engine = FakeAssignment()
        record = dummy._enrich_inventory_record_assignment(
            {
                "assigned_person": "Hambone",
                "cert_number": "123",
                "card_title": "Test Card",
                "source_sheet": "Lot.xlsx",
                "purchase_price": 40,
                "inventory_value": 100,
                "card_ladder_value": 150,
                "card_ladder_comps_average": 100,
                "cy_value": 90,
            }
        )
        self.assertEqual(record["best_company"], "Arena Club")
        self.assertEqual(record["estimated_payout"], 95)
        self.assertEqual(record["inventory_value"], 150)
        self.assertEqual(record["card_ladder_value"], 150)
        self.assertEqual(record["card_ladder_comps_average"], 100)
        self.assertEqual(dummy.assignment_engine.last_row.card_ladder_value, 150)
        self.assertEqual(dummy.assignment_engine.last_row.card_ladder_comps_average, 100)

    def test_inventory_assignment_uses_saved_sport_for_company_rules(self) -> None:
        class InventoryDummy:
            _money_value = app.CardPipelineApp._money_value
            _inventory_record_key = app.CardPipelineApp._inventory_record_key
            _normalize_inventory_record = app.CardPipelineApp._normalize_inventory_record
            _inventory_workbook_row = app.CardPipelineApp._inventory_workbook_row
            _enrich_inventory_record_assignment = app.CardPipelineApp._enrich_inventory_record_assignment

        dummy = InventoryDummy()
        dummy.assignment_engine = assignment_engine.AssignmentEngine([
            assignment_engine.AssignmentCompany(
                "Arena Club",
                assignment_engine.CompanyRules(
                    accept_all=True,
                    blocks=[assignment_engine.AssignmentRule("Football Anthony Richardson", block=True)],
                ),
                [assignment_engine.PayoutTier(0, 500, 1.05)],
            ),
            assignment_engine.AssignmentCompany(
                "Fanatics",
                assignment_engine.CompanyRules(accept_all=True),
                [assignment_engine.PayoutTier(0, 500, 0.95)],
            ),
        ])

        record = dummy._enrich_inventory_record_assignment(
            {
                "assigned_person": "Tyler Hamlin",
                "cert_number": "98514082",
                "sport": "football",
                "grader": "PSA",
                "card_title": "2023 Panini Select 117 Anthony Richardson PSA 10",
                "source_sheet": "TYLER_NASHVILLE_INVENTORY.xlsx",
                "purchase_price": 20,
                "card_ladder_value": 31,
                "card_ladder_comps_average": 31,
            },
            force=True,
        )

        self.assertEqual(record["best_company"], "Fanatics")
        self.assertEqual(record["estimated_payout"], 29.45)

    def test_inventory_refresh_hydrates_source_values_for_card_ladder_company(self) -> None:
        class InventoryDummy:
            _money_value = app.CardPipelineApp._money_value
            _inventory_record_key = app.CardPipelineApp._inventory_record_key
            _normalize_inventory_record = app.CardPipelineApp._normalize_inventory_record
            _inventory_workbook_row = app.CardPipelineApp._inventory_workbook_row
            _inventory_source_sheet_path = app.CardPipelineApp._inventory_source_sheet_path
            _hydrate_inventory_record_source_values = app.CardPipelineApp._hydrate_inventory_record_source_values
            _enrich_inventory_record_assignment = app.CardPipelineApp._enrich_inventory_record_assignment

        with TemporaryDirectory() as tmp:
            root = Path(tmp)
            received_dir = root / "RECEIVED SHEETS"
            received_dir.mkdir()
            source = received_dir / "Lot.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.append(["Certification Number", "Grader", "Card Description", "Purchase Price", "Card Ladder Value", "Comps"])
            sheet.append(["87378266", "PSA", "2023 Panini Phoenix Playing With Fire Bryce Young PSA 10", 20, 44, 32])
            workbook.save(source)

            old_received = app.RECEIVED_SHEETS_DIR
            old_working = app.WORKING_SHEETS_DIR
            old_incoming = app.INCOMING_SHEETS_DIR
            app.RECEIVED_SHEETS_DIR = received_dir
            app.WORKING_SHEETS_DIR = root / "WORKING SHEETS"
            app.INCOMING_SHEETS_DIR = root / "INCOMING SHEETS"
            dummy = InventoryDummy()
            dummy.assignment_engine = assignment_engine.AssignmentEngine([
                assignment_engine.AssignmentCompany(
                    "Arena Club",
                    assignment_engine.CompanyRules(accept_all=True),
                    [assignment_engine.PayoutTier(0, 500, 1.0)],
                    value_source="comps",
                ),
                assignment_engine.AssignmentCompany(
                    "Fanatics",
                    assignment_engine.CompanyRules(accept_all=True),
                    [assignment_engine.PayoutTier(0, 500, 0.95)],
                    value_source="card_ladder",
                ),
            ])
            try:
                record = dummy._enrich_inventory_record_assignment(
                    {
                        "assigned_person": "Kevin Hambone",
                        "cert_number": "87378266",
                        "grader": "PSA",
                        "card_title": "2023 Panini Phoenix Playing With Fire Bryce Young PSA 10",
                        "purchase_price": 20,
                        "inventory_value": 32,
                        "source_sheet": "Lot.xlsx",
                        "status": "Active",
                    },
                    force=True,
                )

                self.assertEqual(record["best_company"], "Fanatics")
                self.assertEqual(record["estimated_payout"], 41.8)
                self.assertEqual(record["inventory_value"], 44)
                self.assertEqual(record["card_ladder_value"], 44)
                self.assertEqual(record["card_ladder_comps_average"], 32)
            finally:
                app.RECEIVED_SHEETS_DIR = old_received
                app.WORKING_SHEETS_DIR = old_working
                app.INCOMING_SHEETS_DIR = old_incoming

    def test_add_inventory_records_recalculates_stale_nobody_takes_assignment(self) -> None:
        class FakeAssignment:
            def recommend(self, row, person=""):
                self.last_person = person
                return types.SimpleNamespace(company="Arena Club", payout=95)

        class InventoryDummy:
            _money_value = app.CardPipelineApp._money_value
            _inventory_record_key = app.CardPipelineApp._inventory_record_key
            _normalize_inventory_record = app.CardPipelineApp._normalize_inventory_record
            _load_inventory_ledger = app.CardPipelineApp._load_inventory_ledger
            _save_inventory_ledger = app.CardPipelineApp._save_inventory_ledger
            _inventory_workbook_row = app.CardPipelineApp._inventory_workbook_row
            _enrich_inventory_record_assignment = app.CardPipelineApp._enrich_inventory_record_assignment
            add_inventory_records = app.CardPipelineApp.add_inventory_records
            refresh_inventory_tab = lambda self: None

        with TemporaryDirectory() as tmp:
            old_inventory = app.INVENTORY_LEDGER_PATH
            app.INVENTORY_LEDGER_PATH = Path(tmp) / "inventory_ledger.json"
            dummy = InventoryDummy()
            dummy.assignment_engine = FakeAssignment()
            try:
                dummy.add_inventory_records([
                    {
                        "assigned_person": "Kevin Hambone",
                        "cert_number": "123",
                        "card_title": "Test Card",
                        "source_sheet": "Hambone.xlsx",
                        "purchase_price": 40,
                        "inventory_value": 100,
                        "best_company": app.NO_COMPANY_TAKES_LABEL,
                    }
                ])

                ledger = json.loads(app.INVENTORY_LEDGER_PATH.read_text(encoding="utf-8"))["items"]
                self.assertEqual(ledger[0]["best_company"], "Arena Club")
                self.assertEqual(ledger[0]["estimated_payout"], 95)
                self.assertEqual(dummy.assignment_engine.last_person, "Kevin Hambone")
            finally:
                app.INVENTORY_LEDGER_PATH = old_inventory

    def test_retarget_inventory_rows_for_source_changes_owner_without_duplicates(self) -> None:
        class FakeAssignment:
            def recommend(self, row, person=""):
                self.last_person = person
                return types.SimpleNamespace(company=f"{person} Club", payout=88)

        class InventoryDummy:
            _money_value = app.CardPipelineApp._money_value
            _inventory_record_key = app.CardPipelineApp._inventory_record_key
            _normalize_inventory_record = app.CardPipelineApp._normalize_inventory_record
            _load_inventory_ledger = app.CardPipelineApp._load_inventory_ledger
            _save_inventory_ledger = app.CardPipelineApp._save_inventory_ledger
            _inventory_workbook_row = app.CardPipelineApp._inventory_workbook_row
            _enrich_inventory_record_assignment = app.CardPipelineApp._enrich_inventory_record_assignment
            _retarget_inventory_rows_for_source = app.CardPipelineApp._retarget_inventory_rows_for_source

        with TemporaryDirectory() as tmp:
            old_inventory = app.INVENTORY_LEDGER_PATH
            app.INVENTORY_LEDGER_PATH = Path(tmp) / "inventory_ledger.json"
            dummy = InventoryDummy()
            dummy.assignment_engine = FakeAssignment()
            try:
                dummy._save_inventory_ledger([
                    dummy._normalize_inventory_record({"assigned_person": "Lucas", "cert_number": "123", "source_sheet": "Lot A.xlsx", "card_title": "Test Card", "inventory_value": 100, "best_company": "Arena Club", "estimated_payout": 95}),
                    dummy._normalize_inventory_record({"assigned_person": "Mikey", "cert_number": "123", "source_sheet": "Lot A.xlsx", "card_title": "Test Card", "inventory_value": 100, "best_company": "Arena Club", "estimated_payout": 90}),
                ])

                self.assertEqual(dummy._retarget_inventory_rows_for_source("Lot A.xlsx", "Mikey"), 1)

                ledger = json.loads(app.INVENTORY_LEDGER_PATH.read_text(encoding="utf-8"))["items"]
                self.assertEqual(len(ledger), 1)
                self.assertEqual(ledger[0]["assigned_person"], "Mikey")
                self.assertEqual(ledger[0]["inventory_key"], "123|lot a.xlsx|mikey")
                self.assertEqual(ledger[0]["best_company"], "Mikey Club")
                self.assertEqual(ledger[0]["estimated_payout"], 88)
            finally:
                app.INVENTORY_LEDGER_PATH = old_inventory

    def test_received_inventory_reconcile_skips_company_sheet_rows(self) -> None:
        class ReconcileDummy:
            _money_value = app.CardPipelineApp._money_value
            _inventory_record_key = app.CardPipelineApp._inventory_record_key
            _normalize_inventory_record = app.CardPipelineApp._normalize_inventory_record
            _company_sheet_source_cert_keys = app.CardPipelineApp._company_sheet_source_cert_keys
            _received_certs_in_workbook = app.CardPipelineApp._received_certs_in_workbook
            _received_inventory_candidate_records_for_sheet = app.CardPipelineApp._received_inventory_candidate_records_for_sheet
            _received_inventory_candidate_records = app.CardPipelineApp._received_inventory_candidate_records
            _home_sheet_key = app.CardPipelineApp._home_sheet_key

        with TemporaryDirectory() as tmp:
            root = Path(tmp)
            received_dir = root / "RECEIVED SHEETS"
            received_dir.mkdir()
            received_path = received_dir / "Hambone Lot.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.append(["Certification Number", "Grader", "Card Description", "Purchase Price", "Comps"])
            sheet.append(["111", "PSA", "Hambone Inventory Card", 40, 100])
            sheet.append(["222", "PSA", "Company Card", 50, 120])
            workbook.save(received_path)

            company_dir = root / "COMPANY SHEETS"
            company_path = company_dir / "Arena Club" / "Arena Club.xlsx"
            company_path.parent.mkdir(parents=True)
            company_workbook = Workbook()
            company_sheet = company_workbook.active
            company_sheet.title = "Week of 2026-06-15"
            company_sheet.append(["Date Added", "Source Sheet", "Source", "Certification Number", "Grader", "Card Description", "Purchase Price", "Card Ladder Value", "Comps", "CY Estimate", "CY Confidence", "Best Company", "Estimated Payout", "Status", "Notes"])
            company_sheet.append(["2026-06-17", "Hambone Lot.xlsx", "", "222", "PSA", "Company Card", 50, "", 120, "", "", "Arena Club", 90, "Received", ""])
            company_workbook.save(company_path)

            old_received = app.RECEIVED_SHEETS_DIR
            old_incoming = app.INCOMING_SHEETS_DIR
            old_working = app.WORKING_SHEETS_DIR
            old_company = app.COMPANY_SHEETS_DIR
            app.RECEIVED_SHEETS_DIR = received_dir
            app.INCOMING_SHEETS_DIR = root / "INCOMING SHEETS"
            app.WORKING_SHEETS_DIR = root / "WORKING SHEETS"
            app.COMPANY_SHEETS_DIR = company_dir
            dummy = ReconcileDummy()
            dummy.home_sheet_markers = {"Received|Hambone Lot.xlsx": {"assigned_person": "Hambone"}}
            try:
                records = dummy._received_inventory_candidate_records()
                self.assertEqual(len(records), 1)
                self.assertEqual(records[0]["cert_number"], "111")
                self.assertEqual(records[0]["assigned_person"], "Hambone")
            finally:
                app.RECEIVED_SHEETS_DIR = old_received
                app.INCOMING_SHEETS_DIR = old_incoming
                app.WORKING_SHEETS_DIR = old_working
                app.COMPANY_SHEETS_DIR = old_company

    def test_received_inventory_reconcile_skips_unassigned_sheet_markers(self) -> None:
        class ReconcileDummy:
            _money_value = app.CardPipelineApp._money_value
            _inventory_record_key = app.CardPipelineApp._inventory_record_key
            _normalize_inventory_record = app.CardPipelineApp._normalize_inventory_record
            _company_sheet_source_cert_keys = app.CardPipelineApp._company_sheet_source_cert_keys
            _received_certs_in_workbook = app.CardPipelineApp._received_certs_in_workbook
            _received_inventory_candidate_records_for_sheet = app.CardPipelineApp._received_inventory_candidate_records_for_sheet
            _received_inventory_candidate_records = app.CardPipelineApp._received_inventory_candidate_records
            _home_sheet_key = app.CardPipelineApp._home_sheet_key

        with TemporaryDirectory() as tmp:
            root = Path(tmp)
            received_dir = root / "RECEIVED SHEETS"
            received_dir.mkdir()
            workbook = Workbook()
            sheet = workbook.active
            sheet.append(["Certification Number", "Grader", "Card Description", "Purchase Price", "Comps"])
            sheet.append(["151740304", "PSA", "Deleted Test Card", 40, 100])
            workbook.save(received_dir / "HAMBONE_MIKEY_6_16_TEST.xlsx")

            old_received = app.RECEIVED_SHEETS_DIR
            old_incoming = app.INCOMING_SHEETS_DIR
            old_working = app.WORKING_SHEETS_DIR
            old_company = app.COMPANY_SHEETS_DIR
            app.RECEIVED_SHEETS_DIR = received_dir
            app.INCOMING_SHEETS_DIR = root / "INCOMING SHEETS"
            app.WORKING_SHEETS_DIR = root / "WORKING SHEETS"
            app.COMPANY_SHEETS_DIR = root / "COMPANY SHEETS"
            dummy = ReconcileDummy()
            dummy.home_sheet_markers = {}
            try:
                self.assertEqual(dummy._received_inventory_candidate_records(), [])
            finally:
                app.RECEIVED_SHEETS_DIR = old_received
                app.INCOMING_SHEETS_DIR = old_incoming
                app.WORKING_SHEETS_DIR = old_working
                app.COMPANY_SHEETS_DIR = old_company

    def test_received_inventory_candidates_preserve_source_sheet_sport(self) -> None:
        class InventoryDummy:
            _money_value = app.CardPipelineApp._money_value
            _inventory_record_key = app.CardPipelineApp._inventory_record_key
            _normalize_inventory_record = app.CardPipelineApp._normalize_inventory_record
            _inventory_sport_from_value = app.CardPipelineApp._inventory_sport_from_value
            _received_inventory_candidate_records_for_sheet = app.CardPipelineApp._received_inventory_candidate_records_for_sheet

        with TemporaryDirectory() as tmp:
            path = Path(tmp) / "manual-baseball.xlsx"
            write_working_sheet(
                path,
                [
                    WorkbookRow(
                        excel_row=2,
                        cert_number="0010355805",
                        grader="BGS",
                        card_title="2017 Bowman Chrome Prospect Autographs Gold Shimmer Refractors #CPACF Clint Frazier BGS 9.5",
                        category="baseball",
                        card_ladder_comps_average=432,
                    )
                ],
            )

            records = InventoryDummy()._received_inventory_candidate_records_for_sheet("Received", path, "James Copeland", set())

        self.assertEqual(len(records), 1)
        self.assertEqual(records[0]["sport"], "baseball")

    def test_home_all_received_marker_adds_received_sheet_inventory(self) -> None:
        class Status:
            def __init__(self):
                self.value = ""

            def set(self, value):
                self.value = value

        class MarkerDummy:
            _money_value = app.CardPipelineApp._money_value
            _inventory_record_key = app.CardPipelineApp._inventory_record_key
            _normalize_inventory_record = app.CardPipelineApp._normalize_inventory_record
            _load_inventory_ledger = app.CardPipelineApp._load_inventory_ledger
            _save_inventory_ledger = app.CardPipelineApp._save_inventory_ledger
            _received_certs_in_workbook = app.CardPipelineApp._received_certs_in_workbook
            _company_sheet_source_cert_keys = lambda self: set()
            _received_inventory_candidate_records_for_sheet = app.CardPipelineApp._received_inventory_candidate_records_for_sheet
            _sync_received_sheet_inventory_to_ledger = app.CardPipelineApp._sync_received_sheet_inventory_to_ledger
            _home_sheet_key = app.CardPipelineApp._home_sheet_key
            _split_home_sheet_key = app.CardPipelineApp._split_home_sheet_key
            _sheet_path_for_stage = app.CardPipelineApp._sheet_path_for_stage
            _delete_sheet_marker = app.CardPipelineApp._delete_sheet_marker
            _move_home_sheet_to_stage = app.CardPipelineApp._move_home_sheet_to_stage
            _move_sheet_to_received = app.CardPipelineApp._move_sheet_to_received
            _move_received_sheet_to_incoming = app.CardPipelineApp._move_received_sheet_to_incoming
            _move_working_sheet_to_incoming = app.CardPipelineApp._move_working_sheet_to_incoming
            _marker_for_stage = app.CardPipelineApp._marker_for_stage
            _retarget_inventory_rows_for_source = lambda self, source, person: 0
            _enrich_inventory_record_assignment = lambda self, record, force=False: record
            add_inventory_records = app.CardPipelineApp.add_inventory_records
            save_home_sheet_markers = app.CardPipelineApp.save_home_sheet_markers
            _load_sheet_markers = app.CardPipelineApp._load_sheet_markers
            _save_sheet_markers = app.CardPipelineApp._save_sheet_markers
            refresh_working_sheets = lambda self: None
            refresh_received_sheets = lambda self: None
            refresh_incoming_index = lambda self: None
            refresh_home = lambda self: None
            refresh_inventory_tab = lambda self, enrich=False: None

        with TemporaryDirectory() as tmp:
            root = Path(tmp)
            incoming_dir = root / "INCOMING SHEETS"
            received_dir = root / "RECEIVED SHEETS"
            incoming_dir.mkdir()
            received_dir.mkdir()
            source_path = incoming_dir / "Hambone Lot.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.append(["Certification Number", "Grader", "Card Description", "Purchase Price", "Comps"])
            sheet.append(["111", "PSA", "Hambone Inventory Card", 40, 100])
            workbook.save(source_path)

            old_pipeline = app.CARD_PIPELINE_DIR
            old_incoming = app.INCOMING_SHEETS_DIR
            old_received = app.RECEIVED_SHEETS_DIR
            old_working = app.WORKING_SHEETS_DIR
            old_company = app.COMPANY_SHEETS_DIR
            old_inventory = app.INVENTORY_LEDGER_PATH
            old_markers = app.SHEET_MARKERS_PATH
            app.CARD_PIPELINE_DIR = root
            app.INCOMING_SHEETS_DIR = incoming_dir
            app.RECEIVED_SHEETS_DIR = received_dir
            app.WORKING_SHEETS_DIR = root / "WORKING SHEETS"
            app.COMPANY_SHEETS_DIR = root / "COMPANY SHEETS"
            app.INVENTORY_LEDGER_PATH = root / "inventory_ledger.json"
            app.SHEET_MARKERS_PATH = root / "sheet_markers.json"
            dummy = MarkerDummy()
            dummy.lucas_identity = {"display_name": "Tester", "machine": "Test"}
            dummy.home_selected_sheet_key = "Incoming|Hambone Lot.xlsx"
            dummy.home_sheet_markers = {"Incoming|Hambone Lot.xlsx": {"assigned_person": "Kevin Hambone"}}
            dummy.home_sheet_paths = {"Incoming": {"Hambone Lot.xlsx": source_path}, "Working": {}, "Received": {}}
            dummy.received_sheet_paths = {}
            dummy.deleted_sheet_marker_keys = set()
            dummy.status_var = Status()
            try:
                dummy.save_home_sheet_markers(
                    {
                        "incoming_proper": False,
                        "tracking_number": "",
                        "all_received": True,
                        "assigned_person": "Kevin Hambone",
                    }
                )

                inventory = json.loads(app.INVENTORY_LEDGER_PATH.read_text(encoding="utf-8"))["items"]
                self.assertTrue((received_dir / "Hambone Lot.xlsx").exists())
                self.assertEqual(len(inventory), 1)
                self.assertEqual(inventory[0]["cert_number"], "111")
                self.assertEqual(inventory[0]["assigned_person"], "Kevin Hambone")
                self.assertIn("Added 1 inventory row", dummy.status_var.value)
            finally:
                app.CARD_PIPELINE_DIR = old_pipeline
                app.INCOMING_SHEETS_DIR = old_incoming
                app.RECEIVED_SHEETS_DIR = old_received
                app.WORKING_SHEETS_DIR = old_working
                app.COMPANY_SHEETS_DIR = old_company
                app.INVENTORY_LEDGER_PATH = old_inventory
                app.SHEET_MARKERS_PATH = old_markers

    def test_inventory_rows_can_be_removed_for_deleted_source_sheet(self) -> None:
        class InventoryDummy:
            _money_value = app.CardPipelineApp._money_value
            _inventory_record_key = app.CardPipelineApp._inventory_record_key
            _normalize_inventory_record = app.CardPipelineApp._normalize_inventory_record
            _load_inventory_ledger = app.CardPipelineApp._load_inventory_ledger
            _save_inventory_ledger = app.CardPipelineApp._save_inventory_ledger
            _remove_inventory_rows_for_source = app.CardPipelineApp._remove_inventory_rows_for_source

        with TemporaryDirectory() as tmp:
            old_inventory = app.INVENTORY_LEDGER_PATH
            app.INVENTORY_LEDGER_PATH = Path(tmp) / "inventory_ledger.json"
            dummy = InventoryDummy()
            try:
                dummy._save_inventory_ledger([
                    dummy._normalize_inventory_record({"assigned_person": "Unassigned", "cert_number": "151740304", "source_sheet": "HAMBONE_MIKEY_6_16_TEST.xlsx"}),
                    dummy._normalize_inventory_record({"assigned_person": "Hambone", "cert_number": "222", "source_sheet": "SCOTSBORO_HAMBONE_6_16_26.xlsx"}),
                ])

                self.assertEqual(dummy._remove_inventory_rows_for_source("HAMBONE_MIKEY_6_16_TEST.xlsx"), 1)

                ledger = json.loads(app.INVENTORY_LEDGER_PATH.read_text(encoding="utf-8"))["items"]
                self.assertEqual(len(ledger), 1)
                self.assertEqual(ledger[0]["source_sheet"], "SCOTSBORO_HAMBONE_6_16_26.xlsx")
            finally:
                app.INVENTORY_LEDGER_PATH = old_inventory

    def test_inventory_rows_can_be_deleted_by_key(self) -> None:
        class InventoryDummy:
            _money_value = app.CardPipelineApp._money_value
            _inventory_record_key = app.CardPipelineApp._inventory_record_key
            _normalize_inventory_record = app.CardPipelineApp._normalize_inventory_record
            _load_inventory_ledger = app.CardPipelineApp._load_inventory_ledger
            _save_inventory_ledger = app.CardPipelineApp._save_inventory_ledger
            _delete_inventory_records_by_keys = app.CardPipelineApp._delete_inventory_records_by_keys

        with TemporaryDirectory() as tmp:
            old_inventory = app.INVENTORY_LEDGER_PATH
            old_pipeline = app.CARD_PIPELINE_DIR
            app.INVENTORY_LEDGER_PATH = Path(tmp) / "inventory_ledger.json"
            app.CARD_PIPELINE_DIR = Path(tmp)
            dummy = InventoryDummy()
            dummy.lucas_identity = {"display_name": "Tester", "machine": "Test"}
            try:
                first = dummy._normalize_inventory_record({"assigned_person": "Kevin Hambone", "cert_number": "111", "source_sheet": "A.xlsx"})
                second = dummy._normalize_inventory_record({"assigned_person": "James Copeland", "cert_number": "222", "source_sheet": "B.xlsx"})
                third = dummy._normalize_inventory_record({"assigned_person": "Kevin Hambone", "cert_number": "333", "source_sheet": "C.xlsx"})
                dummy._save_inventory_ledger([first, second, third])

                deleted = dummy._delete_inventory_records_by_keys({first["inventory_key"], third["inventory_key"]})

                self.assertEqual(deleted, 2)
                ledger = json.loads(app.INVENTORY_LEDGER_PATH.read_text(encoding="utf-8"))["items"]
                self.assertEqual(len(ledger), 1)
                self.assertEqual(ledger[0]["inventory_key"], second["inventory_key"])
                self.assertEqual(ledger[0]["cert_number"], "222")
            finally:
                app.INVENTORY_LEDGER_PATH = old_inventory
                app.CARD_PIPELINE_DIR = old_pipeline

    def test_delete_person_records_unassigns_markers_inventory_and_profit(self) -> None:
        class DeletePersonDummy:
            _money_value = app.CardPipelineApp._money_value
            _inventory_record_key = app.CardPipelineApp._inventory_record_key
            _normalize_inventory_record = app.CardPipelineApp._normalize_inventory_record
            _load_inventory_ledger = app.CardPipelineApp._load_inventory_ledger
            _save_inventory_ledger = app.CardPipelineApp._save_inventory_ledger
            _profit_record_key = app.CardPipelineApp._profit_record_key
            _normalize_profit_record = app.CardPipelineApp._normalize_profit_record
            _load_profit_ledger = app.CardPipelineApp._load_profit_ledger
            _save_profit_ledger = app.CardPipelineApp._save_profit_ledger
            delete_person_records = app.CardPipelineApp.delete_person_records

            def _save_sheet_markers(self):
                self.saved_markers = True

        with TemporaryDirectory() as tmp:
            old_inventory = app.INVENTORY_LEDGER_PATH
            old_profit = app.PROFIT_LEDGER_PATH
            app.INVENTORY_LEDGER_PATH = Path(tmp) / "inventory_ledger.json"
            app.PROFIT_LEDGER_PATH = Path(tmp) / "profit_ledger.json"
            dummy = DeletePersonDummy()
            dummy.home_sheet_markers = {
                "Incoming|A.xlsx": {"assigned_person": "Lucas"},
                "Incoming|B.xlsx": {"assigned_person": "Mikey"},
            }
            dummy.saved_markers = False
            try:
                dummy._save_inventory_ledger([
                    dummy._normalize_inventory_record({"assigned_person": "Lucas", "cert_number": "1", "source_sheet": "A.xlsx"}),
                    dummy._normalize_inventory_record({"assigned_person": "Mikey", "cert_number": "2", "source_sheet": "B.xlsx"}),
                ])
                dummy._save_profit_ledger([
                    dummy._normalize_profit_record({"assigned_person": "Lucas", "cert_number": "1", "source_sheet": "A.xlsx", "company": "Arena", "date_added": "2026-06-17"}),
                    dummy._normalize_profit_record({"assigned_person": "Mikey", "cert_number": "2", "source_sheet": "B.xlsx", "company": "Arena", "date_added": "2026-06-17"}),
                ])

                counts = dummy.delete_person_records("Lucas")

                self.assertEqual(counts, {"markers": 1, "inventory": 1, "profit": 1})
                self.assertTrue(dummy.saved_markers)
                self.assertEqual(dummy.home_sheet_markers["Incoming|A.xlsx"]["assigned_person"], "")
                inventory = json.loads(app.INVENTORY_LEDGER_PATH.read_text(encoding="utf-8"))["items"]
                profit = json.loads(app.PROFIT_LEDGER_PATH.read_text(encoding="utf-8"))
                self.assertEqual(inventory[0]["assigned_person"], "Unassigned")
                self.assertEqual(inventory[1]["assigned_person"], "Mikey")
                self.assertEqual(profit[0]["assigned_person"], "")
                self.assertEqual(profit[1]["assigned_person"], "Mikey")
            finally:
                app.INVENTORY_LEDGER_PATH = old_inventory
                app.PROFIT_LEDGER_PATH = old_profit

    def test_home_person_filter_limits_sheet_names(self) -> None:
        class Var:
            def __init__(self, value: str) -> None:
                self.value = value

            def get(self) -> str:
                return self.value

        class HomeDummy:
            _home_sheet_key = app.CardPipelineApp._home_sheet_key
            _home_person_filter = app.CardPipelineApp._home_person_filter
            _home_sheet_matches_person_filter = app.CardPipelineApp._home_sheet_matches_person_filter
            _filtered_home_sheet_names = app.CardPipelineApp._filtered_home_sheet_names

        dummy = HomeDummy()
        dummy.home_person_var = Var("Kevin")
        dummy.home_sheet_paths = {
            "Incoming": {
                "kevin.xlsx": Path("kevin.xlsx"),
                "james.xlsx": Path("james.xlsx"),
                "blank.xlsx": Path("blank.xlsx"),
            }
        }
        dummy.home_sheet_markers = {
            "Incoming|kevin.xlsx": {"assigned_person": "Kevin Hambone"},
            "Incoming|james.xlsx": {"assigned_person": "James Copeland"},
            "Incoming|blank.xlsx": {},
        }

        self.assertEqual(dummy._filtered_home_sheet_names("Incoming"), ["kevin.xlsx"])

    def test_create_seller_terms_apply_and_restore_purchase_prices(self) -> None:
        class Var:
            def __init__(self, value=""):
                self.value = value

            def get(self):
                return self.value

            def set(self, value):
                self.value = value

        class SellerTermsDummy:
            _money_value = app.CardPipelineApp._money_value
            _seller_terms_rate = app.CardPipelineApp._seller_terms_rate
            _load_seller_terms = app.CardPipelineApp._load_seller_terms
            _seller_terms_match = app.CardPipelineApp._seller_terms_match
            _seller_terms_company_decision = app.CardPipelineApp._seller_terms_company_decision
            _seller_terms_company_price = app.CardPipelineApp._seller_terms_company_price
            _restore_create_seller_term_prices = app.CardPipelineApp._restore_create_seller_term_prices
            _network_mode_enabled = app.CardPipelineApp._network_mode_enabled
            apply_create_seller_terms = app.CardPipelineApp.apply_create_seller_terms

            def __init__(self):
                self.create_network_mode_var = Var(True)
                self.seller_terms_seller_var = Var("John")
                self.seller_terms_sheet_type_var = Var("Arena Club")
                self.status_var = Var("")
                self.refreshed = 0
                self.assignment_engine = types.SimpleNamespace(
                    evaluate=lambda row: (
                        [types.SimpleNamespace(company="Arena Club", payout=95.0, source_value=100.0)]
                        if row.cert_number == "1"
                        else []
                    )
                )
                self.intake_rows = [
                    WorkbookRow(excel_row=2, cert_number="1", grader="PSA", card_title="Test", existing_value=10, cy_value=100),
                    WorkbookRow(excel_row=3, cert_number="2", grader="PSA", card_title="No Arena Match", existing_value=20),
                ]

            def _refresh_table(self):
                self.refreshed += 1

        with TemporaryDirectory() as tmp:
            old_terms = app.SELLER_TERMS_PATH
            app.SELLER_TERMS_PATH = Path(tmp) / "seller_terms.csv"
            app.SELLER_TERMS_PATH.write_text(
                "Seller,Sheet Type,Seller Rate,Deduction\n"
                "John,Arena Club,80%,\n"
                "Mary,Arena Club,,5%\n",
                encoding="utf-8",
            )
            dummy = SellerTermsDummy()
            try:
                self.assertEqual(dummy.apply_create_seller_terms(), 1)
                self.assertEqual(dummy.intake_rows[0].existing_value, 80)
                self.assertEqual(dummy.intake_rows[1].existing_value, 20)
                self.assertIn("Arena Club rule value", dummy.status_var.value)

                dummy.seller_terms_seller_var.set("Mary")
                dummy.seller_terms_sheet_type_var.set("Arena Club")
                self.assertEqual(dummy.apply_create_seller_terms(), 1)
                self.assertEqual(dummy.intake_rows[0].existing_value, 90)
                self.assertEqual(dummy.intake_rows[1].existing_value, 20)
                self.assertIn("payout minus 5%", dummy.status_var.value)

                dummy.seller_terms_sheet_type_var.set("")
                self.assertEqual(dummy.apply_create_seller_terms(), 0)
                self.assertEqual(dummy.intake_rows[0].existing_value, 10)
                self.assertEqual(dummy.intake_rows[1].existing_value, 20)

                dummy.seller_terms_sheet_type_var.set("Arena Club")
                dummy.create_network_mode_var.set(False)
                self.assertEqual(dummy.apply_create_seller_terms(), 0)
                self.assertEqual(dummy.intake_rows[0].existing_value, 10)
                self.assertEqual(dummy.intake_rows[1].existing_value, 20)
            finally:
                app.SELLER_TERMS_PATH = old_terms

    def test_inventory_recomp_empty_scope_only_queues_missing_selected_fields(self) -> None:
        class RecompScopeDummy:
            _money_value = app.CardPipelineApp._money_value
            _inventory_recomp_record_matches_scope = app.CardPipelineApp._inventory_recomp_record_matches_scope

        dummy = RecompScopeDummy()
        full_record = {"card_ladder_value": 50, "card_ladder_comps_average": 40, "cy_value": 30, "cy_confidence": 4}
        missing_comps = {"card_ladder_value": 50, "card_ladder_comps_average": None, "cy_value": 30, "cy_confidence": 4}
        missing_cy = {"card_ladder_value": 50, "card_ladder_comps_average": 40, "cy_value": None, "cy_confidence": ""}
        missing_cy_confidence = {"card_ladder_value": 50, "card_ladder_comps_average": 40, "cy_value": 30, "cy_confidence": ""}

        self.assertTrue(dummy._inventory_recomp_record_matches_scope(full_record, {"scope": app.COMP_SCOPE_ALL, "card_ladder_comps": True}))
        self.assertFalse(dummy._inventory_recomp_record_matches_scope(full_record, {"scope": app.COMP_SCOPE_EMPTY, "card_ladder_value": True, "card_ladder_comps": True}))
        self.assertTrue(dummy._inventory_recomp_record_matches_scope(missing_comps, {"scope": app.COMP_SCOPE_EMPTY, "card_ladder_comps": True}))
        self.assertFalse(dummy._inventory_recomp_record_matches_scope(missing_cy, {"scope": app.COMP_SCOPE_EMPTY, "card_ladder_comps": True}))
        self.assertTrue(dummy._inventory_recomp_record_matches_scope(missing_cy, {"scope": app.COMP_SCOPE_EMPTY, "cy": True}))
        self.assertTrue(dummy._inventory_recomp_record_matches_scope(missing_cy_confidence, {"scope": app.COMP_SCOPE_EMPTY, "cy": True}))

    def test_inventory_recomp_sync_preserves_cy_confidence(self) -> None:
        class RecompSyncDummy:
            _money_value = app.CardPipelineApp._money_value
            _inventory_record_key = app.CardPipelineApp._inventory_record_key
            _normalize_inventory_record = app.CardPipelineApp._normalize_inventory_record
            _load_inventory_ledger = app.CardPipelineApp._load_inventory_ledger
            _save_inventory_ledger = app.CardPipelineApp._save_inventory_ledger
            _sync_inventory_recomp_results = app.CardPipelineApp._sync_inventory_recomp_results

            def _inventory_sport_from_value(self, sport, card_title):
                return sport

            def _enrich_inventory_record_assignment(self, record, force=False):
                return self._normalize_inventory_record(record)

        with TemporaryDirectory() as tmp:
            old_inventory = app.INVENTORY_LEDGER_PATH
            app.INVENTORY_LEDGER_PATH = Path(tmp) / "inventory_ledger.json"
            dummy = RecompSyncDummy()
            original = dummy._normalize_inventory_record(
                {
                    "assigned_person": "Kevin Hambone",
                    "cert_number": "123",
                    "card_title": "Test Card",
                    "source_sheet": "INVENTORY.xlsx",
                    "status": "Active",
                    "cy_value": None,
                    "cy_confidence": "",
                }
            )
            dummy._save_inventory_ledger([original])
            dummy.state = types.SimpleNamespace(
                lock=threading.Lock(),
                rows=[
                    WorkbookRow(
                        excel_row=2,
                        cert_number="123",
                        card_title="Test Card",
                        grader="PSA",
                        cy_value=88.0,
                        cy_confidence=4,
                        status="Ready",
                    )
                ],
            )
            dummy.inventory_recomp_context = {
                "keys_by_excel_row": {2: original["inventory_key"]},
                "features": {"cy": True},
            }
            try:
                self.assertEqual(dummy._sync_inventory_recomp_results(), 1)
                rows = json.loads(app.INVENTORY_LEDGER_PATH.read_text(encoding="utf-8"))["items"]
                self.assertEqual(rows[0]["cy_value"], 88.0)
                self.assertEqual(rows[0]["cy_confidence"], 4)
            finally:
                app.INVENTORY_LEDGER_PATH = old_inventory

    def test_edit_inventory_row_updates_visible_fields_and_rebuilds_key(self) -> None:
        class FakeTree:
            def selection(self):
                return ["row-1"]

        class FakeStatus:
            def __init__(self):
                self.value = ""

            def set(self, value):
                self.value = value

        class InventoryEditDummy:
            _money_value = app.CardPipelineApp._money_value
            _inventory_record_key = app.CardPipelineApp._inventory_record_key
            _normalize_inventory_record = app.CardPipelineApp._normalize_inventory_record
            _load_inventory_ledger = app.CardPipelineApp._load_inventory_ledger
            _save_inventory_ledger = app.CardPipelineApp._save_inventory_ledger
            _update_inventory_record_by_key = app.CardPipelineApp._update_inventory_record_by_key
            edit_selected_inventory_row = app.CardPipelineApp.edit_selected_inventory_row
            refresh_inventory_tab = lambda self: None

            def _inventory_edit_row_dialog(self, _record):
                return {
                    "date_added": "2026-06-20",
                    "assigned_person": "James Copeland",
                    "sport": "football",
                    "cert_number": "83861755",
                    "grader": "PSA",
                    "card_title": "2010 Donruss Rated Rookies 95 Tim Tebow PSA 10",
                    "purchase_price": 100,
                    "card_ladder_value": 116,
                    "card_ladder_comps_average": 99.97,
                    "cy_value": None,
                    "cy_confidence": "",
                    "best_company": "FANATICS",
                    "estimated_payout": 107.88,
                    "source_sheet": "JAMES_NASHVILLE_PICKUPS.xlsx",
                }

        with TemporaryDirectory() as tmp:
            old_pipeline = app.CARD_PIPELINE_DIR
            old_inventory = app.INVENTORY_LEDGER_PATH
            app.CARD_PIPELINE_DIR = Path(tmp)
            app.INVENTORY_LEDGER_PATH = Path(tmp) / "inventory_ledger.json"
            dummy = InventoryEditDummy()
            dummy.inventory_tree = FakeTree()
            dummy.lucas_identity = {"display_name": "Tester", "machine": "Test"}
            dummy.status_var = FakeStatus()
            record = dummy._normalize_inventory_record({"assigned_person": "Lucas", "cert_number": "1", "source_sheet": "Old.xlsx", "status": "Active"})
            dummy._save_inventory_ledger([record])
            dummy.inventory_tree_records = {"row-1": record}
            try:
                dummy.edit_selected_inventory_row()

                rows = json.loads(app.INVENTORY_LEDGER_PATH.read_text(encoding="utf-8"))["items"]
                self.assertEqual(len(rows), 1)
                edited = rows[0]
                self.assertEqual(edited["assigned_person"], "James Copeland")
                self.assertEqual(edited["cert_number"], "83861755")
                self.assertEqual(edited["source_sheet"], "JAMES_NASHVILLE_PICKUPS.xlsx")
                self.assertEqual(edited["card_ladder_value"], 116)
                self.assertEqual(edited["card_ladder_comps_average"], 99.97)
                self.assertEqual(edited["best_company"], "FANATICS")
                self.assertEqual(edited["estimated_payout"], 107.88)
                self.assertEqual(edited["inventory_key"], "83861755|james_nashville_pickups.xlsx|james copeland")
                self.assertEqual(dummy.status_var.value, "Edited 1 inventory row(s).")
            finally:
                app.CARD_PIPELINE_DIR = old_pipeline
                app.INVENTORY_LEDGER_PATH = old_inventory

    def test_inventory_records_can_move_to_company_sheets(self) -> None:
        class FakeTree:
            def selection(self):
                return ["row-1"]

        class FakeAssignment:
            def recommend(self, row, person=""):
                return types.SimpleNamespace(company="Arena Club", payout=90)

        class FakeStatus:
            def __init__(self):
                self.value = ""

            def set(self, value):
                self.value = value

        class InventoryMoveDummy:
            _money_value = app.CardPipelineApp._money_value
            _inventory_record_key = app.CardPipelineApp._inventory_record_key
            _normalize_inventory_record = app.CardPipelineApp._normalize_inventory_record
            _load_inventory_ledger = app.CardPipelineApp._load_inventory_ledger
            _save_inventory_ledger = app.CardPipelineApp._save_inventory_ledger
            _inventory_workbook_row = app.CardPipelineApp._inventory_workbook_row
            _mark_inventory_records_moved_to_company = app.CardPipelineApp._mark_inventory_records_moved_to_company
            _inventory_record_can_move_to_company_sheet = app.CardPipelineApp._inventory_record_can_move_to_company_sheet
            move_selected_inventory_to_company_sheets = app.CardPipelineApp.move_selected_inventory_to_company_sheets
            _move_inventory_records_to_company_sheets = app.CardPipelineApp._move_inventory_records_to_company_sheets
            _profit_record_key = app.CardPipelineApp._profit_record_key
            _normalize_profit_record = app.CardPipelineApp._normalize_profit_record
            _load_profit_ledger = app.CardPipelineApp._load_profit_ledger
            _save_profit_ledger = app.CardPipelineApp._save_profit_ledger
            record_profit_sales = app.CardPipelineApp.record_profit_sales
            refresh_inventory_tab = lambda self: None
            refresh_profit_tab = lambda self: None
            _append_activity = lambda self, action, summary, details=None: None

        with TemporaryDirectory() as tmp:
            root = Path(tmp)
            old_pipeline = app.CARD_PIPELINE_DIR
            old_company = app.COMPANY_SHEETS_DIR
            old_profit = app.PROFIT_LEDGER_PATH
            old_inventory = app.INVENTORY_LEDGER_PATH
            app.CARD_PIPELINE_DIR = root
            app.COMPANY_SHEETS_DIR = root / "COMPANY SHEETS"
            app.PROFIT_LEDGER_PATH = root / "profit_ledger.json"
            app.INVENTORY_LEDGER_PATH = root / "inventory_ledger.json"
            dummy = InventoryMoveDummy()
            dummy.assignment_engine = FakeAssignment()
            dummy.inventory_tree = FakeTree()
            dummy.lucas_identity = {"display_name": "Tester", "machine": "Test"}
            dummy.status_var = FakeStatus()
            record = dummy._normalize_inventory_record(
                {
                    "assigned_person": "Lucas",
                    "cert_number": "123",
                    "grader": "PSA",
                    "card_title": "Test Card",
                    "source_sheet": "Quick Load.xlsx",
                    "purchase_price": 40,
                    "inventory_value": 100,
                    "best_company": "Arena Club",
                    "estimated_payout": 90,
                    "status": "Active",
                }
            )
            dummy._save_inventory_ledger([record])
            dummy.inventory_tree_records = {"row-1": record}
            try:
                with patch("app.messagebox.askyesno", return_value=True):
                    dummy.move_selected_inventory_to_company_sheets()

                company_rows = read_company_profit_records(app.COMPANY_SHEETS_DIR)
                profit_rows = json.loads(app.PROFIT_LEDGER_PATH.read_text(encoding="utf-8"))
                inventory = json.loads(app.INVENTORY_LEDGER_PATH.read_text(encoding="utf-8"))["items"]
                self.assertEqual(len(company_rows), 1)
                self.assertEqual(company_rows[0]["company"], "Arena Club")
                self.assertEqual(len(profit_rows), 1)
                self.assertEqual(profit_rows[0]["assigned_person"], "Lucas")
                self.assertEqual(inventory, [])
            finally:
                app.CARD_PIPELINE_DIR = old_pipeline
                app.COMPANY_SHEETS_DIR = old_company
                app.PROFIT_LEDGER_PATH = old_profit
                app.INVENTORY_LEDGER_PATH = old_inventory

    def test_nobody_takes_inventory_record_cannot_move_to_company_sheet(self) -> None:
        class InventoryDummy:
            _inventory_record_can_move_to_company_sheet = app.CardPipelineApp._inventory_record_can_move_to_company_sheet

        dummy = InventoryDummy()
        self.assertFalse(
            dummy._inventory_record_can_move_to_company_sheet(
                {"status": "Active", "best_company": app.NO_COMPANY_TAKES_LABEL, "estimated_payout": None}
            )
        )
        self.assertTrue(
            dummy._inventory_record_can_move_to_company_sheet(
                {"status": "Active", "best_company": "Arena Club", "estimated_payout": 95}
            )
        )
        self.assertFalse(
            dummy._inventory_record_can_move_to_company_sheet(
                {"item_type": "Raw", "status": "Active", "best_company": "Arena Club", "estimated_payout": 95}
            )
        )

    def test_mark_inventory_record_sold_writes_profit_and_removes_inventory(self) -> None:
        class SoldDummy:
            _money_value = app.CardPipelineApp._money_value
            _inventory_record_key = app.CardPipelineApp._inventory_record_key
            _normalize_inventory_record = app.CardPipelineApp._normalize_inventory_record
            _load_inventory_ledger = app.CardPipelineApp._load_inventory_ledger
            _save_inventory_ledger = app.CardPipelineApp._save_inventory_ledger
            _mark_inventory_record_sold = app.CardPipelineApp._mark_inventory_record_sold
            _profit_record_key = app.CardPipelineApp._profit_record_key
            _normalize_profit_record = app.CardPipelineApp._normalize_profit_record
            _load_profit_ledger = app.CardPipelineApp._load_profit_ledger
            _save_profit_ledger = app.CardPipelineApp._save_profit_ledger
            _inventory_sale_profit_record = app.CardPipelineApp._inventory_sale_profit_record
            _inventory_sale_expense_record = app.CardPipelineApp._inventory_sale_expense_record
            _general_sold_sheet_name = app.CardPipelineApp._general_sold_sheet_name
            mark_inventory_record_sold = app.CardPipelineApp.mark_inventory_record_sold
            record_profit_sales = app.CardPipelineApp.record_profit_sales
            refresh_profit_tab = lambda self: None
            _append_activity = lambda self, action, summary, details=None: None

        with TemporaryDirectory() as tmp:
            old_pipeline = app.CARD_PIPELINE_DIR
            old_profit = app.PROFIT_LEDGER_PATH
            old_inventory = app.INVENTORY_LEDGER_PATH
            app.CARD_PIPELINE_DIR = Path(tmp)
            app.PROFIT_LEDGER_PATH = Path(tmp) / "profit_ledger.json"
            app.INVENTORY_LEDGER_PATH = Path(tmp) / "inventory_ledger.json"
            dummy = SoldDummy()
            dummy.lucas_identity = {"display_name": "Tester", "machine": "Test"}
            record = dummy._normalize_inventory_record(
                {
                    "assigned_person": "Hambone",
                    "cert_number": "123",
                    "grader": "PSA",
                    "card_title": "Test Card",
                    "source_sheet": "Lot.xlsx",
                    "purchase_price": 40,
                    "inventory_value": 100,
                    "best_company": "Arena Club",
                    "estimated_payout": 90,
                    "status": "Active",
                }
            )
            dummy._save_inventory_ledger([record])
            try:
                self.assertTrue(dummy.mark_inventory_record_sold(record, "Arena Club", 95, expense_type="Shipping", expense_amount=12.5, expense_notes="label"))
                inventory = json.loads(app.INVENTORY_LEDGER_PATH.read_text(encoding="utf-8"))["items"]
                profit = json.loads(app.PROFIT_LEDGER_PATH.read_text(encoding="utf-8"))
                self.assertEqual(inventory, [])
                self.assertEqual(len(profit), 2)
                sale_row = next(record for record in profit if record.get("record_type") != "expense")
                expense_row = next(record for record in profit if record.get("record_type") == "expense")
                self.assertEqual(sale_row["company"], "Arena Club")
                self.assertEqual(sale_row["assigned_person"], "Hambone")
                self.assertEqual(sale_row["sale_price"], 95.0)
                self.assertEqual(sale_row["profit"], 55.0)
                self.assertEqual(expense_row["company"], "Expense: Shipping")
                self.assertEqual(expense_row["assigned_person"], "Hambone")
                self.assertEqual(expense_row["related_type"], "Card")
                self.assertEqual(expense_row["source_sheet"], "Lot.xlsx")
                self.assertEqual(expense_row["cert_number"], "123")
                self.assertEqual(expense_row["expense_amount"], 12.5)
                self.assertEqual(expense_row["profit"], -12.5)
                self.assertEqual(expense_row["notes"], "label")
            finally:
                app.CARD_PIPELINE_DIR = old_pipeline
                app.PROFIT_LEDGER_PATH = old_profit
                app.INVENTORY_LEDGER_PATH = old_inventory

    def test_manual_inventory_add_accepts_cert_without_grader(self) -> None:
        class ManualAddDummy:
            _money_value = app.CardPipelineApp._money_value
            _inventory_record_key = app.CardPipelineApp._inventory_record_key
            _normalize_inventory_record = app.CardPipelineApp._normalize_inventory_record
            _load_inventory_ledger = app.CardPipelineApp._load_inventory_ledger
            _save_inventory_ledger = app.CardPipelineApp._save_inventory_ledger
            _inventory_sport_from_value = app.CardPipelineApp._inventory_sport_from_value
            _next_raw_item_id = app.CardPipelineApp._next_raw_item_id
            add_raw_inventory_card = app.CardPipelineApp.add_raw_inventory_card
            refresh_inventory_tab = lambda self: None
            _append_activity = lambda self, action, summary, details=None: None

            def _raw_inventory_card_dialog(self):
                return {
                    "assigned_person": "Kevin Hambone",
                    "cert_number": "12345678",
                    "grader": "",
                    "card_title": "2024 Panini Prizm Test Card",
                    "purchase_price": 25,
                }

            def _enrich_inventory_record_assignment(self, record, force=False):
                return self._normalize_inventory_record(record)

        with TemporaryDirectory() as tmp:
            old_pipeline = app.CARD_PIPELINE_DIR
            old_inventory = app.INVENTORY_LEDGER_PATH
            app.CARD_PIPELINE_DIR = Path(tmp)
            app.INVENTORY_LEDGER_PATH = Path(tmp) / "inventory_ledger.json"
            dummy = ManualAddDummy()
            dummy.lucas_identity = {"display_name": "Tester", "machine": "Test"}
            class Status:
                def __init__(self):
                    self.value = ""
                def set(self, value):
                    self.value = value
            dummy.status_var = Status()
            try:
                dummy.add_raw_inventory_card()
                inventory = json.loads(app.INVENTORY_LEDGER_PATH.read_text(encoding="utf-8"))["items"]
                self.assertEqual(len(inventory), 1)
                self.assertEqual(inventory[0]["cert_number"], "12345678")
                self.assertEqual(inventory[0]["grader"], "")
                self.assertEqual(inventory[0]["item_type"], "Graded")
                self.assertEqual(inventory[0]["item_id"], "")
                self.assertEqual(inventory[0]["source_sheet"], "Manual Inventory")
                self.assertEqual(dummy.status_var.value, "Added inventory card 12345678.")
            finally:
                app.CARD_PIPELINE_DIR = old_pipeline
                app.INVENTORY_LEDGER_PATH = old_inventory

    def test_raw_inventory_record_uses_item_id_for_sold_profit_and_expense(self) -> None:
        class SoldDummy:
            _money_value = app.CardPipelineApp._money_value
            _inventory_record_key = app.CardPipelineApp._inventory_record_key
            _normalize_inventory_record = app.CardPipelineApp._normalize_inventory_record
            _load_inventory_ledger = app.CardPipelineApp._load_inventory_ledger
            _save_inventory_ledger = app.CardPipelineApp._save_inventory_ledger
            _mark_inventory_record_sold = app.CardPipelineApp._mark_inventory_record_sold
            _profit_record_key = app.CardPipelineApp._profit_record_key
            _normalize_profit_record = app.CardPipelineApp._normalize_profit_record
            _load_profit_ledger = app.CardPipelineApp._load_profit_ledger
            _save_profit_ledger = app.CardPipelineApp._save_profit_ledger
            _inventory_sale_profit_record = app.CardPipelineApp._inventory_sale_profit_record
            _inventory_sale_expense_record = app.CardPipelineApp._inventory_sale_expense_record
            _general_sold_sheet_name = app.CardPipelineApp._general_sold_sheet_name
            _next_raw_item_id = app.CardPipelineApp._next_raw_item_id
            mark_inventory_record_sold = app.CardPipelineApp.mark_inventory_record_sold
            record_profit_sales = app.CardPipelineApp.record_profit_sales
            refresh_profit_tab = lambda self: None
            _append_activity = lambda self, action, summary, details=None: None

        with TemporaryDirectory() as tmp:
            old_pipeline = app.CARD_PIPELINE_DIR
            old_profit = app.PROFIT_LEDGER_PATH
            old_inventory = app.INVENTORY_LEDGER_PATH
            app.CARD_PIPELINE_DIR = Path(tmp)
            app.PROFIT_LEDGER_PATH = Path(tmp) / "profit_ledger.json"
            app.INVENTORY_LEDGER_PATH = Path(tmp) / "inventory_ledger.json"
            dummy = SoldDummy()
            dummy.lucas_identity = {"display_name": "Tester", "machine": "Test"}
            try:
                item_id = dummy._next_raw_item_id([])
                record = dummy._normalize_inventory_record(
                    {
                        "item_type": "Raw",
                        "item_id": item_id,
                        "assigned_person": "Hambone",
                        "card_title": "2024 Panini Prizm Test Raw Card",
                        "purchase_price": 40,
                        "card_ladder_value": 100,
                        "best_company": "Arena Club",
                        "estimated_payout": 90,
                        "source_sheet": "Raw Inventory",
                        "status": "Active",
                    }
                )
                self.assertEqual(record["inventory_key"], item_id.lower())
                dummy._save_inventory_ledger([record])

                self.assertTrue(dummy.mark_inventory_record_sold(record, "Arena Club", 95, expense_type="Shipping", expense_amount=5))

                inventory = json.loads(app.INVENTORY_LEDGER_PATH.read_text(encoding="utf-8"))["items"]
                profit = json.loads(app.PROFIT_LEDGER_PATH.read_text(encoding="utf-8"))
                self.assertEqual(inventory, [])
                sale_row = next(record for record in profit if record.get("record_type") != "expense")
                expense_row = next(record for record in profit if record.get("record_type") == "expense")
                self.assertEqual(sale_row["item_type"], "Raw")
                self.assertEqual(sale_row["item_id"], item_id)
                self.assertEqual(sale_row["cert_number"], "")
                self.assertEqual(expense_row["item_id"], item_id)
                self.assertEqual(expense_row["cert_number"], "")
            finally:
                app.CARD_PIPELINE_DIR = old_pipeline
                app.PROFIT_LEDGER_PATH = old_profit
                app.INVENTORY_LEDGER_PATH = old_inventory

    def test_blank_inventory_sale_company_uses_person_general_sold_sheet(self) -> None:
        class SoldDummy:
            _money_value = app.CardPipelineApp._money_value
            _inventory_record_key = app.CardPipelineApp._inventory_record_key
            _normalize_inventory_record = app.CardPipelineApp._normalize_inventory_record
            _profit_record_key = app.CardPipelineApp._profit_record_key
            _normalize_profit_record = app.CardPipelineApp._normalize_profit_record
            _general_sold_sheet_name = app.CardPipelineApp._general_sold_sheet_name
            _inventory_sale_profit_record = app.CardPipelineApp._inventory_sale_profit_record

        dummy = SoldDummy()
        record = dummy._inventory_sale_profit_record(
            {
                "assigned_person": "Kevin Hambone",
                "cert_number": "123",
                "card_title": "Test Card",
                "source_sheet": "Original.xlsx",
                "purchase_price": 40,
            },
            "",
            95,
        )
        self.assertEqual(record["company"], "General Sold")
        self.assertEqual(record["source_sheet"], "Kevin Hambone General Sold")
        self.assertEqual(record["assigned_person"], "Kevin Hambone")

    def test_refund_profit_record_returns_card_to_inventory_and_removes_company_row(self) -> None:
        class RefundDummy:
            _money_value = app.CardPipelineApp._money_value
            _profit_record_key = app.CardPipelineApp._profit_record_key
            _normalize_profit_record = app.CardPipelineApp._normalize_profit_record
            _load_profit_ledger = app.CardPipelineApp._load_profit_ledger
            _save_profit_ledger = app.CardPipelineApp._save_profit_ledger
            _inventory_record_key = app.CardPipelineApp._inventory_record_key
            _normalize_inventory_record = app.CardPipelineApp._normalize_inventory_record
            _load_inventory_ledger = app.CardPipelineApp._load_inventory_ledger
            _save_inventory_ledger = app.CardPipelineApp._save_inventory_ledger
            add_inventory_records = app.CardPipelineApp.add_inventory_records
            _enrich_inventory_record_assignment = lambda self, record: record
            refresh_inventory_tab = lambda self: None

        with TemporaryDirectory() as tmp:
            root = Path(tmp)
            company_dir = root / "COMPANY SHEETS"
            company_path = company_dir / "Arena Club" / "Arena Club.xlsx"
            company_path.parent.mkdir(parents=True)
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Week of 2026-06-15"
            sheet.append(["Date Added", "Source Sheet", "Source", "Certification Number", "Grader", "Card Description", "Purchase Price", "Card Ladder Value", "Comps", "CY Estimate", "CY Confidence", "Best Company", "Estimated Payout", "Status", "Notes"])
            sheet.append(["2026-06-17", "Lot A.xlsx", "", "123", "PSA", "Test Card", 40, "", 100, "", "", "Arena Club", 90, "Received", ""])
            workbook.save(company_path)

            old_company = app.COMPANY_SHEETS_DIR
            old_profit = app.PROFIT_LEDGER_PATH
            old_inventory = app.INVENTORY_LEDGER_PATH
            app.COMPANY_SHEETS_DIR = company_dir
            app.PROFIT_LEDGER_PATH = root / "profit_ledger.json"
            app.INVENTORY_LEDGER_PATH = root / "inventory_ledger.json"
            record = {
                "date_added": "2026-06-17",
                "company": "Arena Club",
                "weekly_sheet_name": "Arena Club.xlsx:Week of 2026-06-15",
                "source_sheet": "Lot A.xlsx",
                "cert_number": "123",
                "grader": "PSA",
                "card_title": "Test Card",
                "purchase_price": 40,
                "sale_price": 90,
                "assigned_person": "Lucas",
            }
            dummy = RefundDummy()
            normalized = dummy._normalize_profit_record(record)
            app.PROFIT_LEDGER_PATH.write_text(json.dumps([normalized]), encoding="utf-8")
            try:
                ledger = [dummy._normalize_profit_record(item) for item in dummy._load_profit_ledger()]
                kept = [item for item in ledger if item["ledger_key"] != normalized["ledger_key"]]
                dummy._save_profit_ledger(kept)
                remove_company_sheet_rows_for_source = app.remove_company_sheet_rows_for_source
                remove_company_sheet_rows_for_source(app.COMPANY_SHEETS_DIR, "Lot A.xlsx", {"123"})
                dummy.add_inventory_records([
                    dummy._normalize_inventory_record(
                        {
                            "assigned_person": "Lucas",
                            "cert_number": "123",
                            "grader": "PSA",
                            "card_title": "Test Card",
                            "purchase_price": 40,
                            "inventory_value": 90,
                            "source_sheet": "Lot A.xlsx",
                            "status": "Active",
                            "notes": "Refunded from sold cards",
                        }
                    )
                ])

                self.assertEqual(json.loads(app.PROFIT_LEDGER_PATH.read_text(encoding="utf-8")), [])
                self.assertEqual(read_company_profit_records(company_dir), [])
                inventory = json.loads(app.INVENTORY_LEDGER_PATH.read_text(encoding="utf-8"))["items"]
                self.assertEqual(len(inventory), 1)
                self.assertEqual(inventory[0]["status"], "Active")
                self.assertEqual(inventory[0]["assigned_person"], "Lucas")
            finally:
                app.COMPANY_SHEETS_DIR = old_company
                app.PROFIT_LEDGER_PATH = old_profit
                app.INVENTORY_LEDGER_PATH = old_inventory

    def test_profit_records_are_enriched_with_assigned_person_from_sheet_marker(self) -> None:
        class ProfitDummy:
            _profit_record_key = app.CardPipelineApp._profit_record_key
            _money_value = app.CardPipelineApp._money_value
            _normalize_profit_record = app.CardPipelineApp._normalize_profit_record
            _home_sheet_key = app.CardPipelineApp._home_sheet_key
            _split_home_sheet_key = app.CardPipelineApp._split_home_sheet_key
            _person_for_profit_record = app.CardPipelineApp._person_for_profit_record
            _enrich_profit_records_with_people = app.CardPipelineApp._enrich_profit_records_with_people
            _filtered_profit_records = app.CardPipelineApp._filtered_profit_records
            _profit_record_date = app.CardPipelineApp._profit_record_date
            _profit_today = lambda self: datetime(2026, 6, 17).date()
            _profit_period_bounds = app.CardPipelineApp._profit_period_bounds

        dummy = ProfitDummy()
        dummy.home_sheet_markers = {"Received|Lot A.xlsx": {"assigned_person": "Lucas"}}
        dummy.profit_person_var = types.SimpleNamespace(get=lambda: "luc")
        dummy.profit_period_var = types.SimpleNamespace(get=lambda: "Total")
        dummy.profit_search_var = types.SimpleNamespace(get=lambda: "")

        rows = dummy._enrich_profit_records_with_people([
            {
                "date_added": "2026-06-11",
                "company": "Arena Club",
                "source_sheet": "Lot A.xlsx",
                "cert_number": "123",
                "card_title": "Test Card",
                "purchase_price": 40,
                "sale_price": 90,
            }
        ])

        self.assertEqual(rows[0]["assigned_person"], "Lucas")
        self.assertEqual(rows[0]["profit"], 50)
        self.assertEqual(dummy._filtered_profit_records(rows), rows)

        dummy.profit_search_var = types.SimpleNamespace(get=lambda: "123 test")
        self.assertEqual(dummy._filtered_profit_records(rows), rows)

        dummy.profit_search_var = types.SimpleNamespace(get=lambda: "fanatics")
        self.assertEqual(dummy._filtered_profit_records(rows), [])

    def test_profit_period_filter_and_chart_series_support_daily_and_overall_views(self) -> None:
        class ProfitDummy:
            _money_value = app.CardPipelineApp._money_value
            _profit_record_date = app.CardPipelineApp._profit_record_date
            _profit_today = lambda self: datetime(2026, 6, 17).date()
            _profit_period_bounds = app.CardPipelineApp._profit_period_bounds
            _profit_period_label = app.CardPipelineApp._profit_period_label
            _profit_graph_label = app.CardPipelineApp._profit_graph_label
            _profit_chart_title = app.CardPipelineApp._profit_chart_title
            _filtered_profit_records = app.CardPipelineApp._filtered_profit_records
            _profit_chart_series = app.CardPipelineApp._profit_chart_series

        rows = [
            {"assigned_person": "Lucas", "date_added": "2026-06-17", "profit": 30},
            {"assigned_person": "Lucas", "date_added": "2026-06-13", "profit": 20},
            {"assigned_person": "Lucas", "date_added": "2026-06-10", "profit": 100},
            {"assigned_person": "Mikey", "date_added": "2026-06-17", "profit": 7},
            {"assigned_person": "Lucas", "date_added": "not-a-date", "profit": 50},
        ]

        dummy = ProfitDummy()
        self.assertEqual(dummy._profit_period_label(), "Year")
        self.assertEqual(dummy._profit_graph_label(), "Overall Profit")
        self.assertEqual(dummy._profit_chart_title(), "Overall Profit (Year)")

        dummy.profit_person_var = types.SimpleNamespace(get=lambda: "luc")
        dummy.profit_period_var = types.SimpleNamespace(get=lambda: "5 Days")
        dummy.profit_graph_var = types.SimpleNamespace(get=lambda: "Daily Trend")
        dummy.profit_search_var = types.SimpleNamespace(get=lambda: "")
        self.assertEqual(dummy._profit_chart_title(), "Daily Trend (5 Days)")

        filtered = dummy._filtered_profit_records(rows)
        days, daily_values = dummy._profit_chart_series(filtered)
        dummy.profit_graph_var = types.SimpleNamespace(get=lambda: "Overall Profit")
        self.assertEqual(dummy._profit_chart_title(), "Overall Profit (5 Days)")
        overall_days, overall_values = dummy._profit_chart_series(filtered)

        self.assertEqual([record["date_added"] for record in filtered], ["2026-06-17", "2026-06-13"])
        self.assertEqual(days, ["2026-06-13", "2026-06-14", "2026-06-15", "2026-06-16", "2026-06-17"])
        self.assertEqual(daily_values, [20, 0.0, 0.0, 0.0, 30])
        self.assertEqual(overall_days, days)
        self.assertEqual(overall_values, [20, 20.0, 20.0, 20.0, 50.0])

        ytd_rows = [
            {"assigned_person": "Lucas", "date_added": "2026-01-05", "profit": 40},
            {"assigned_person": "Lucas", "date_added": "2026-06-17", "profit": 30},
        ]
        dummy.profit_period_var = types.SimpleNamespace(get=lambda: "YTD")
        dummy.profit_graph_var = types.SimpleNamespace(get=lambda: "Daily Trend")
        ytd_filtered = dummy._filtered_profit_records(ytd_rows)
        ytd_days, ytd_values = dummy._profit_chart_series(ytd_filtered)

        self.assertEqual(ytd_days[0], "2026-01-01")
        self.assertEqual(ytd_days[-1], "2026-06-17")
        self.assertEqual(len(ytd_days), 168)
        self.assertEqual(ytd_values[ytd_days.index("2026-01-05")], 40)
        self.assertEqual(ytd_values[ytd_days.index("2026-06-17")], 30)

    def test_profit_sheet_rows_group_by_person_and_source_sheet(self) -> None:
        class ProfitDummy:
            _money_value = app.CardPipelineApp._money_value
            _profit_sheet_rows = app.CardPipelineApp._profit_sheet_rows

        dummy = ProfitDummy()
        rows = [
            {"assigned_person": "Lucas", "source_sheet": "Lot A.xlsx", "company": "Arena Club", "purchase_price": 40, "sale_price": 90, "profit": 50, "date_added": "2026-06-11"},
            {"assigned_person": "Lucas", "source_sheet": "Lot A.xlsx", "company": "Fanatics", "purchase_price": 20, "sale_price": 30, "profit": 10, "date_added": "2026-06-12"},
            {"assigned_person": "Mikey", "source_sheet": "Lot B.xlsx", "company": "Arena Club", "purchase_price": 100, "sale_price": 80, "profit": -20, "date_added": "2026-06-12"},
        ]

        grouped = dummy._profit_sheet_rows(rows)
        lot_a = next(item for item in grouped if item["sheet"] == "Lot A.xlsx")

        self.assertEqual(lot_a["person"], "Lucas")
        self.assertEqual(lot_a["cards"], 2)
        self.assertEqual(lot_a["purchase"], 60)
        self.assertEqual(lot_a["sale"], 120)
        self.assertEqual(lot_a["profit"], 60)
        self.assertEqual(lot_a["companies"], "Arena Club, Fanatics")


class PhotoOcrSpeedTests(unittest.TestCase):
    def test_detect_regions_skips_extra_label_sweeps_when_dense_target_is_met(self) -> None:
        calls: list[str] = []

        def fake_detect(_client, _bytes, _mime, prompt):
            calls.append(prompt)
            if prompt == multi_card_extraction.DETECTION_PROMPT:
                return [
                    {"card_index": index + 1, "position": f"slot {index + 1}", "bbox": [index * 50, 0, index * 50 + 40, 400], "detection_confidence": "high"}
                    for index in range(multi_card_extraction.PHOTO_OCR_REGION_TARGET)
                ]
            return []

        with patch.object(multi_card_extraction, "_detect_regions_for_prompt", side_effect=fake_detect), \
                patch.object(multi_card_extraction, "_detect_best_row_regions", return_value=[]), \
                patch.object(multi_card_extraction, "_add_uncovered_edge_regions", side_effect=lambda regions: regions):
            regions = multi_card_extraction._detect_regions_sync(object(), b"image", "image/jpeg")

        self.assertEqual(len(regions), multi_card_extraction.PHOTO_OCR_REGION_TARGET)
        self.assertIn(multi_card_extraction.DETECTION_PROMPT, calls)
        self.assertNotIn(multi_card_extraction.LABEL_DETECTION_PROMPT, calls)
        self.assertNotIn(multi_card_extraction.LABEL_SWEEP_PROMPT, calls)

    def test_detect_regions_uses_label_sweeps_below_dense_target(self) -> None:
        calls: list[str] = []

        def fake_detect(_client, _bytes, _mime, prompt):
            calls.append(prompt)
            if prompt == multi_card_extraction.DETECTION_PROMPT:
                return [
                    {"card_index": 1, "position": "left", "bbox": [0, 0, 200, 400], "detection_confidence": "high"},
                    {"card_index": 2, "position": "middle", "bbox": [220, 0, 420, 400], "detection_confidence": "high"},
                    {"card_index": 3, "position": "right", "bbox": [440, 0, 640, 400], "detection_confidence": "high"},
                ]
            if prompt == multi_card_extraction.LABEL_DETECTION_PROMPT:
                return [
                    {"card_index": index + 1, "position": f"label {index + 1}", "bbox": [index * 50, 0, index * 50 + 40, 100], "detection_confidence": "high"}
                    for index in range(6)
                ]
            return []

        with patch.object(multi_card_extraction, "_detect_regions_for_prompt", side_effect=fake_detect), \
                patch.object(multi_card_extraction, "_detect_best_row_regions", return_value=[]), \
                patch.object(multi_card_extraction, "_detect_best_prompt_regions", return_value=[]), \
                patch.object(multi_card_extraction, "_add_uncovered_edge_regions", side_effect=lambda regions: regions):
            regions = multi_card_extraction._detect_regions_sync(object(), b"image", "image/jpeg")

        self.assertTrue(regions)
        self.assertIn(multi_card_extraction.LABEL_DETECTION_PROMPT, calls)

    def test_identify_cards_reports_crop_progress_and_preserves_order(self) -> None:
        callbacks: list[str] = []
        regions = [
            {"card_index": 1, "position": "left", "bbox": [0, 0, 200, 400], "detection_confidence": "high"},
            {"card_index": 2, "position": "right", "bbox": [220, 0, 420, 400], "detection_confidence": "medium"},
        ]

        def fake_identify(_client, crop_b64):
            return {
                "is_graded_slab": True,
                "grading_company": "PSA",
                "cert_number": "111" if crop_b64 == "crop-1" else "222",
                "player": "Player",
                "year": "2020",
                "set": "Test",
                "card_number": "",
                "parallel": "",
                "subset": "",
                "grade": "10",
                "category": "baseball",
                "confidence": "high",
                "label_text": "label",
            }

        with patch.object(multi_card_extraction, "_prepare_image", return_value=(b"image", "image/jpeg")), \
                patch.object(multi_card_extraction, "_detect_regions_sync", return_value=regions), \
                patch.object(multi_card_extraction, "_decode_image", return_value=object()), \
                patch.object(multi_card_extraction, "_crop_region_to_base64", side_effect=["crop-1", "crop-2"]), \
                patch.object(multi_card_extraction, "_identify_crop_sync", side_effect=fake_identify):
            cards = multi_card_extraction.identify_cards_sync(object(), "fake-b64", progress_callback=callbacks.append)

        self.assertEqual([card["cert_number"] for card in cards], ["111", "222"])
        self.assertTrue(any("Detected 2 card(s)" in message for message in callbacks))
        self.assertTrue(any("Read 2/2" in message for message in callbacks))

    def test_identify_cards_keeps_detected_slab_when_crop_ocr_fails(self) -> None:
        regions = [
            {"card_index": 1, "position": "left", "bbox": [0, 0, 200, 400], "detection_confidence": "high"},
            {"card_index": 2, "position": "right", "bbox": [220, 0, 420, 400], "detection_confidence": "medium"},
        ]

        def fake_identify(_client, crop_b64):
            if crop_b64 == "crop-2":
                raise RuntimeError("label unreadable")
            return {
                "is_graded_slab": True,
                "grading_company": "PSA",
                "cert_number": "111",
                "player": "Player",
                "year": "2020",
                "set": "Test",
                "card_number": "",
                "parallel": "",
                "subset": "",
                "grade": "10",
                "category": "baseball",
                "confidence": "high",
                "label_text": "label",
            }

        with patch.object(multi_card_extraction, "_prepare_image", return_value=(b"image", "image/jpeg")), \
                patch.object(multi_card_extraction, "_detect_regions_sync", return_value=regions), \
                patch.object(multi_card_extraction, "_decode_image", return_value=object()), \
                patch.object(multi_card_extraction, "_crop_region_to_base64", side_effect=["crop-1", "crop-2"]), \
                patch.object(multi_card_extraction, "_identify_crop_sync", side_effect=fake_identify):
            cards = multi_card_extraction.identify_cards_sync(object(), "fake-b64")

        self.assertEqual(len(cards), 2)
        self.assertEqual(cards[0]["cert_number"], "111")
        self.assertEqual(cards[1]["card_index"], 2)
        self.assertTrue(cards[1]["is_graded_slab"])
        self.assertIn("label unreadable", cards[1]["error"])

    def test_photo_table_accepts_detected_slab_without_readable_inventory(self) -> None:
        card = {
            "card_index": 2,
            "position": "right",
            "is_graded_slab": True,
            "detection_confidence": "medium",
            "error": "label unreadable",
        }

        self.assertTrue(app.CardPipelineApp._photo_card_has_inventory(object(), card))
        row = app.CardPipelineApp._photo_card_to_row(object(), Path("dense.jpg"), card)
        self.assertEqual(row["source"], "Photo: dense.jpg")
        self.assertIn("right", row["notes"])
        self.assertIn("OCR review needed", row["notes"])

    def test_mobile_inventory_search_and_duplicate_add_flow(self) -> None:
        class MobileDummy:
            _money_value = app.CardPipelineApp._money_value
            _inventory_record_key = app.CardPipelineApp._inventory_record_key
            _normalize_inventory_record = app.CardPipelineApp._normalize_inventory_record
            _load_inventory_ledger = app.CardPipelineApp._load_inventory_ledger
            _save_inventory_ledger = app.CardPipelineApp._save_inventory_ledger
            _inventory_workbook_row = app.CardPipelineApp._inventory_workbook_row
            _inventory_source_sheet_path = app.CardPipelineApp._inventory_source_sheet_path
            _hydrate_inventory_record_source_values = app.CardPipelineApp._hydrate_inventory_record_source_values
            _enrich_inventory_record_assignment = app.CardPipelineApp._enrich_inventory_record_assignment
            _mobile_inventory_payload_record = app.CardPipelineApp._mobile_inventory_payload_record
            _mobile_inventory_json_record = app.CardPipelineApp._mobile_inventory_json_record
            _next_raw_item_id = app.CardPipelineApp._next_raw_item_id
            mobile_inventory_search = app.CardPipelineApp.mobile_inventory_search
            mobile_inventory_add = app.CardPipelineApp.mobile_inventory_add
            _append_activity = lambda self, action, summary, details=None: None

            def __init__(self, root: Path) -> None:
                self.events = queue.Queue()
                self.lucas_identity = {"display_name": "test", "machine": "test"}

            def _known_people(self):
                return sorted(
                    {
                        str(record.get("assigned_person") or "").strip()
                        for record in self._load_inventory_ledger()
                        if str(record.get("assigned_person") or "").strip()
                    },
                    key=str.lower,
                )

        old_root = app.CARD_PIPELINE_DIR
        old_inventory = app.INVENTORY_LEDGER_PATH
        with TemporaryDirectory() as tmp:
            root = Path(tmp)
            try:
                app.CARD_PIPELINE_DIR = root
                app.INVENTORY_LEDGER_PATH = root / "inventory_ledger.json"
                dummy = MobileDummy(root)
                dummy._save_inventory_ledger([
                    dummy._normalize_inventory_record(
                        {
                            "assigned_person": "Kevin Hambone",
                            "cert_number": "123456",
                            "grader": "PSA",
                            "card_title": "2024 Test Player Silver PSA 10",
                            "purchase_price": 42,
                            "inventory_value": 88,
                            "source_sheet": "Existing.xlsx",
                            "source": "Barcode",
                        }
                    )
                ])

                search = dummy.mobile_inventory_search({"query": "123456"})
                self.assertTrue(search["ok"])
                self.assertEqual(search["count"], 1)
                self.assertEqual(search["items"][0]["purchase_price_display"], "$42.00")
                self.assertEqual(search["people"], ["Kevin Hambone"])
                self.assertEqual(dummy.mobile_inventory_search({"person": "Kevin"})["count"], 1)
                self.assertEqual(dummy.mobile_inventory_search({"person": "Mike"})["count"], 0)

                duplicate = dummy.mobile_inventory_add({"cert_number": "123456", "purchase_price": "50"})
                self.assertFalse(duplicate["ok"])
                self.assertTrue(duplicate["duplicate"])

                update = dummy.mobile_inventory_add({"cert_number": "123456", "purchase_price": "50", "update_existing": True})
                self.assertTrue(update["ok"])
                self.assertEqual(update["action"], "updated")
                self.assertEqual(update["record"]["purchase_price"], 50.0)

                added = dummy.mobile_inventory_add({"cert_number": "777888", "grader": "SGC", "card_title": "Mobile Added Card", "purchase_price": "12.50", "source": "Show"})
                self.assertTrue(added["ok"])
                self.assertEqual(added["action"], "added")
                self.assertEqual(added["record"]["source"], "Show")

                raw_added = dummy.mobile_inventory_add({"card_title": "2024 Panini Prizm Test Raw Card", "purchase_price": "8"})
                self.assertTrue(raw_added["ok"])
                self.assertEqual(raw_added["record"]["item_type"], "Raw")
                self.assertTrue(raw_added["record"]["item_id"].startswith("RAW-"))
                self.assertEqual(raw_added["record"]["cert_number"], "")
                raw_search = dummy.mobile_inventory_search({"query": raw_added["record"]["item_id"]})
                self.assertEqual(raw_search["count"], 1)
                self.assertEqual(raw_search["items"][0]["item_id"], raw_added["record"]["item_id"])
            finally:
                app.CARD_PIPELINE_DIR = old_root
                app.INVENTORY_LEDGER_PATH = old_inventory

    def test_mobile_expenses_profit_summary_and_payouts(self) -> None:
        class MobileFinanceDummy:
            _money_value = app.CardPipelineApp._money_value
            _profit_record_key = app.CardPipelineApp._profit_record_key
            _load_profit_ledger = app.CardPipelineApp._load_profit_ledger
            _save_profit_ledger = app.CardPipelineApp._save_profit_ledger
            _normalize_profit_record = app.CardPipelineApp._normalize_profit_record
            _person_for_profit_record = app.CardPipelineApp._person_for_profit_record
            _enrich_profit_records_with_people = app.CardPipelineApp._enrich_profit_records_with_people
            _append_profit_records = app.CardPipelineApp._append_profit_records
            _profit_record_date = app.CardPipelineApp._profit_record_date
            _profit_today = lambda self: datetime(2026, 6, 19).date()
            _profit_period_bounds = app.CardPipelineApp._profit_period_bounds
            _mobile_profit_rows = app.CardPipelineApp._mobile_profit_rows
            _mobile_profit_chart_series = app.CardPipelineApp._mobile_profit_chart_series
            mobile_profit_summary = app.CardPipelineApp.mobile_profit_summary
            mobile_expense_add = app.CardPipelineApp.mobile_expense_add
            mobile_payouts = app.CardPipelineApp.mobile_payouts

            def __init__(self):
                self.events = queue.Queue()
                self.lucas_identity = {"display_name": "Tester", "machine": "Test"}
                self.home_sheet_markers = {}
                self.home_sheet_paths = {"Incoming": {}, "Received": {}}
                self.home_sheet_summaries = {}

            def _known_people(self):
                return ["Kevin Hambone", "Mike Seller"]

            def _payout_sheet_items(self):
                return [
                    {
                        "name": "Lot A.xlsx",
                        "stage": "Received",
                        "person": "Mike Seller",
                        "paid": False,
                        "row_count": 3,
                        "received_count": 3,
                        "payout_balance": 120.0,
                        "status": "Ready",
                    },
                    {
                        "name": "Lot B.xlsx",
                        "stage": "Received",
                        "person": "Kevin Hambone",
                        "paid": True,
                        "row_count": 2,
                        "received_count": 2,
                        "payout_balance": 0.0,
                        "status": "Paid",
                    },
                ]

        with TemporaryDirectory() as tmp:
            old_pipeline = app.CARD_PIPELINE_DIR
            old_ledger = app.PROFIT_LEDGER_PATH
            app.CARD_PIPELINE_DIR = Path(tmp)
            app.PROFIT_LEDGER_PATH = Path(tmp) / "profit_ledger.json"
            try:
                dummy = MobileFinanceDummy()
                dummy._save_profit_ledger(
                    [
                        {
                            "assigned_person": "Kevin Hambone",
                            "cert_number": "123",
                            "source_sheet": "Lot A.xlsx",
                            "company": "Arena",
                            "purchase_price": 40,
                            "sale_price": 100,
                            "date_added": "2026-06-18",
                        }
                    ]
                )

                added = dummy.mobile_expense_add(
                    {
                        "person": "Kevin Hambone",
                        "date": "2026-06-18",
                        "expense_type": "Travel Meal",
                        "amount": "15",
                        "related_type": "Card",
                        "cert_number": "123",
                        "notes": "Dinner",
                    }
                )
                self.assertTrue(added["ok"])
                summary = dummy.mobile_profit_summary({"person": "Kevin", "period": "YTD", "graph": "Overall Profit"})
                self.assertTrue(summary["ok"])
                self.assertEqual(summary["totals"]["gross_profit"], 60.0)
                self.assertEqual(summary["totals"]["expenses"], 15.0)
                self.assertEqual(summary["totals"]["net_profit"], 45.0)
                self.assertEqual(summary["chart"]["values"][-1], 45.0)

                payouts = dummy.mobile_payouts({"person": "Mike"})
                self.assertTrue(payouts["ok"])
                self.assertEqual(payouts["totals"]["balance"], 120.0)
                self.assertEqual(payouts["summary"][0]["person"], "Mike Seller")
                self.assertEqual(payouts["details"][0]["status"], "Ready")
            finally:
                app.CARD_PIPELINE_DIR = old_pipeline
                app.PROFIT_LEDGER_PATH = old_ledger

    def test_mobile_inventory_mark_sold_removes_inventory_and_records_method(self) -> None:
        class MobileSaleDummy:
            _money_value = app.CardPipelineApp._money_value
            _inventory_record_key = app.CardPipelineApp._inventory_record_key
            _normalize_inventory_record = app.CardPipelineApp._normalize_inventory_record
            _load_inventory_ledger = app.CardPipelineApp._load_inventory_ledger
            _save_inventory_ledger = app.CardPipelineApp._save_inventory_ledger
            _mobile_inventory_json_record = app.CardPipelineApp._mobile_inventory_json_record
            _mark_inventory_record_sold = app.CardPipelineApp._mark_inventory_record_sold
            _general_sold_sheet_name = app.CardPipelineApp._general_sold_sheet_name
            _profit_record_key = app.CardPipelineApp._profit_record_key
            _load_profit_ledger = app.CardPipelineApp._load_profit_ledger
            _save_profit_ledger = app.CardPipelineApp._save_profit_ledger
            _normalize_profit_record = app.CardPipelineApp._normalize_profit_record
            _person_for_profit_record = app.CardPipelineApp._person_for_profit_record
            _append_profit_records = app.CardPipelineApp._append_profit_records
            _profit_record_date = app.CardPipelineApp._profit_record_date
            _inventory_sale_profit_record = app.CardPipelineApp._inventory_sale_profit_record
            mobile_inventory_mark_sold = app.CardPipelineApp.mobile_inventory_mark_sold
            _append_activity = lambda self, action, summary, details=None: None

            def __init__(self):
                self.events = queue.Queue()
                self.lucas_identity = {"display_name": "Tester", "machine": "Test"}
                self.home_sheet_markers = {}

            def _known_people(self):
                return ["Kevin Hambone"]

        with TemporaryDirectory() as tmp:
            old_pipeline = app.CARD_PIPELINE_DIR
            old_inventory = app.INVENTORY_LEDGER_PATH
            old_profit = app.PROFIT_LEDGER_PATH
            app.CARD_PIPELINE_DIR = Path(tmp)
            app.INVENTORY_LEDGER_PATH = Path(tmp) / "inventory_ledger.json"
            app.PROFIT_LEDGER_PATH = Path(tmp) / "profit_ledger.json"
            try:
                dummy = MobileSaleDummy()
                record = dummy._normalize_inventory_record(
                    {
                        "assigned_person": "Kevin Hambone",
                        "cert_number": "999111",
                        "grader": "PSA",
                        "card_title": "Mobile Sold Card",
                        "purchase_price": 40,
                        "inventory_value": 90,
                        "source_sheet": "Lot A.xlsx",
                    }
                )
                dummy._save_inventory_ledger([record])

                result = dummy.mobile_inventory_mark_sold(
                    {
                        "inventory_key": record["inventory_key"],
                        "sale_price": "125",
                        "sale_date": "2026-06-20",
                        "sale_method": "Venmo",
                        "company": "Cash Buyer",
                    }
                )
                self.assertTrue(result["ok"])
                self.assertEqual(dummy._load_inventory_ledger(), [])
                profit = [dummy._normalize_profit_record(item) for item in dummy._load_profit_ledger()]
                self.assertEqual(len(profit), 1)
                self.assertEqual(profit[0]["date_added"], "2026-06-20")
                self.assertEqual(profit[0]["company"], "Cash Buyer")
                self.assertEqual(profit[0]["sale_price"], 125.0)
                self.assertEqual(profit[0]["profit"], 85.0)
                self.assertEqual(profit[0]["sale_method"], "Venmo")
                self.assertIn("Sale method: Venmo", profit[0]["notes"])
            finally:
                app.CARD_PIPELINE_DIR = old_pipeline
                app.INVENTORY_LEDGER_PATH = old_inventory
                app.PROFIT_LEDGER_PATH = old_profit

    def test_mobile_queue_sync_applies_actions_once(self) -> None:
        class MobileQueueDummy:
            _money_value = app.CardPipelineApp._money_value
            _inventory_record_key = app.CardPipelineApp._inventory_record_key
            _normalize_inventory_record = app.CardPipelineApp._normalize_inventory_record
            _load_inventory_ledger = app.CardPipelineApp._load_inventory_ledger
            _save_inventory_ledger = app.CardPipelineApp._save_inventory_ledger
            _mobile_inventory_payload_record = app.CardPipelineApp._mobile_inventory_payload_record
            _mobile_inventory_json_record = app.CardPipelineApp._mobile_inventory_json_record
            _next_raw_item_id = app.CardPipelineApp._next_raw_item_id
            mobile_inventory_add = app.CardPipelineApp.mobile_inventory_add
            _profit_record_key = app.CardPipelineApp._profit_record_key
            _load_profit_ledger = app.CardPipelineApp._load_profit_ledger
            _save_profit_ledger = app.CardPipelineApp._save_profit_ledger
            _normalize_profit_record = app.CardPipelineApp._normalize_profit_record
            _person_for_profit_record = app.CardPipelineApp._person_for_profit_record
            _append_profit_records = app.CardPipelineApp._append_profit_records
            _profit_record_date = app.CardPipelineApp._profit_record_date
            mobile_expense_add = app.CardPipelineApp.mobile_expense_add
            _load_mobile_action_log = app.CardPipelineApp._load_mobile_action_log
            _save_mobile_action_log = app.CardPipelineApp._save_mobile_action_log
            _apply_mobile_queue_action = app.CardPipelineApp._apply_mobile_queue_action
            mobile_queue_sync = app.CardPipelineApp.mobile_queue_sync
            _append_activity = lambda self, action, summary, details=None: None

            def __init__(self):
                self.events = queue.Queue()
                self.lucas_identity = {"display_name": "Tester", "machine": "Test"}
                self.home_sheet_markers = {}

            def _known_people(self):
                return ["Kevin Hambone"]

            def _enrich_inventory_record_assignment(self, record, force=False):
                return self._normalize_inventory_record(record)

        with TemporaryDirectory() as tmp:
            old_pipeline = app.CARD_PIPELINE_DIR
            old_inventory = app.INVENTORY_LEDGER_PATH
            old_profit = app.PROFIT_LEDGER_PATH
            old_mobile_log = app.MOBILE_ACTION_LOG_PATH
            app.CARD_PIPELINE_DIR = Path(tmp)
            app.INVENTORY_LEDGER_PATH = Path(tmp) / "inventory_ledger.json"
            app.PROFIT_LEDGER_PATH = Path(tmp) / "profit_ledger.json"
            app.MOBILE_ACTION_LOG_PATH = Path(tmp) / "mobile_action_log.json"
            try:
                dummy = MobileQueueDummy()
                payload = {
                    "client_id": "phone-1",
                    "actions": [
                        {
                            "id": "phone-1-add-1",
                            "type": "inventory.add",
                            "payload": {
                                "cert_number": "777888",
                                "grader": "PSA",
                                "card_title": "Queued Mobile Card",
                                "purchase_price": "12.50",
                                "assigned_person": "Kevin Hambone",
                            },
                        },
                        {
                            "id": "phone-1-expense-1",
                            "type": "expense.add",
                            "payload": {
                                "person": "Kevin Hambone",
                                "date": "2026-06-22",
                                "expense_type": "Shipping",
                                "amount": "5",
                            },
                        },
                    ],
                }
                first = dummy.mobile_queue_sync(payload)
                self.assertTrue(first["ok"])
                self.assertEqual(first["applied"], 2)
                self.assertEqual(first["skipped"], 0)
                self.assertEqual(len(dummy._load_inventory_ledger()), 1)
                self.assertEqual(len(dummy._load_profit_ledger()), 1)

                second = dummy.mobile_queue_sync(payload)
                self.assertTrue(second["ok"])
                self.assertEqual(second["applied"], 0)
                self.assertEqual(second["skipped"], 2)
                self.assertEqual(len(dummy._load_inventory_ledger()), 1)
                self.assertEqual(len(dummy._load_profit_ledger()), 1)
            finally:
                app.CARD_PIPELINE_DIR = old_pipeline
                app.INVENTORY_LEDGER_PATH = old_inventory
                app.PROFIT_LEDGER_PATH = old_profit
                app.MOBILE_ACTION_LOG_PATH = old_mobile_log

    def test_mobile_card_identify_returns_search_query_from_photo_ocr(self) -> None:
        class MobilePhotoDummy:
            _photo_card_to_row = app.CardPipelineApp._photo_card_to_row
            _photo_card_has_inventory = app.CardPipelineApp._photo_card_has_inventory
            _load_photo_env = lambda self: None
            _mobile_image_parts = app.CardPipelineApp._mobile_image_parts
            _parse_mobile_quick_card_response = app.CardPipelineApp._parse_mobile_quick_card_response
            _mobile_quick_card_to_row = app.CardPipelineApp._mobile_quick_card_to_row
            _mobile_single_card_quick_read = app.CardPipelineApp._mobile_single_card_quick_read

        dummy = MobilePhotoDummy()
        quick_response = json.dumps(
            {
                "grading_company": "PSA",
                "cert_number": "123456789",
                "player": "Test Player",
                "year": "2024",
                "set": "Prizm",
                "grade": "10",
                "confidence": "high",
            }
        )

        class FakeModels:
            def generate_content(self, **_kwargs):
                return types.SimpleNamespace(text=quick_response)

        class FakeClient:
            models = FakeModels()

        fake_genai = types.SimpleNamespace(Client=lambda api_key: FakeClient())
        fake_genai_types = types.SimpleNamespace(
            Part=types.SimpleNamespace(from_bytes=lambda **kwargs: kwargs),
            GenerateContentConfig=lambda **kwargs: kwargs,
            ThinkingConfig=lambda **kwargs: kwargs,
        )
        with patch.object(app, "genai", fake_genai), \
                patch.object(app, "genai_types", fake_genai_types), \
                patch.object(app, "identify_cards_sync") as fallback_ocr, \
                patch.dict(app.os.environ, {"GOOGLE_API_KEY": "test-key"}):
            result = app.CardPipelineApp.mobile_card_identify(dummy, {"image": "data:image/jpeg;base64,ZmFrZQ=="})

        self.assertTrue(result["ok"])
        self.assertEqual(result["mode"], "quick")
        self.assertEqual(result["query"], "123456789")
        self.assertEqual(result["card"]["grader"], "PSA")
        self.assertIn("Test Player", result["card"]["card_title"])
        fallback_ocr.assert_not_called()


if __name__ == "__main__":
    unittest.main()
from openpyxl import Workbook, load_workbook
