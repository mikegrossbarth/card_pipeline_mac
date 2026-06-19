from __future__ import annotations

import json
import sys
import threading
import time
import types
import unittest
from datetime import datetime
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
import assignment_engine
import bridge_server
import google_sheets_import
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
        self.assertIn("CY value: $87.50", row.notes)

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

        self.assertEqual([row.cy_value for row in rows], [87.5])
        close_cy.assert_called_once()

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
        comps = [
            {"date_sold": "May 1, 2026", "price": "$8.00", "title": "Test Card PSA 10"},
            {"date_sold": "Jun 12, 2026", "price": "$25.00", "title": "Test Card PSA 10"},
            {"date_sold": "Jun 1, 2026", "price": "$10.00", "title": "Test Card PSA 10"},
            {"date_sold": "Apr 15, 2026", "price": "$2.00", "title": "Test Card PSA 10"},
        ]

        self.assertEqual(app.comp_price(comps, app.COMP_STRATEGY_STALE_NEWEST), 25.0)

    def test_date_weighted_averages_when_two_newest_are_within_seven_days(self) -> None:
        comps = [
            {"date_sold": "May 1, 2026", "price": "$8.00", "title": "Test Card PSA 10"},
            {"date_sold": "Jun 12, 2026", "price": "$25.00", "title": "Test Card PSA 10"},
            {"date_sold": "Jun 8, 2026", "price": "$10.00", "title": "Test Card PSA 10"},
            {"date_sold": "Apr 15, 2026", "price": "$2.00", "title": "Test Card PSA 10"},
        ]

        self.assertEqual(app.comp_price(comps, app.COMP_STRATEGY_STALE_NEWEST), 11.25)

    def test_date_weighted_dedupes_same_sale_before_comparing_newest_dates(self) -> None:
        comps = [
            {"date_sold": "Apr 5, 2026", "price": "$11.50", "source": "EBAY", "title": "ZANDGEMPORIUM Pokemon Magnezone Stormfront Holo Rare #5 PSA 7 $12.02"},
            {"date_sold": "Apr 5, 2026", "price": "$12.02", "source": "EBAY", "title": "ZANDGEMPORIUM Pokemon Magnezone Stormfront Holo Rare #5 PSA 7"},
            {"date_sold": "Aug 24, 2025", "price": "$10.50", "source": "EBAY", "title": "ALMAR ENTERPRISES 2008 POKEMON DIAMOND & PEARL STORMFRONT MAGNEZONE LV. 44 #5/100 RARE HOLO PSA 7"},
            {"date_sold": "Oct 2, 2022", "price": "$11.50", "source": "EBAY", "title": "Pokemon Magnezone D&P Stormfront Holo Rare #5 PSA 7 -454"},
        ]

        self.assertEqual(app.comp_price(comps, app.COMP_STRATEGY_STALE_NEWEST), 12.02)

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

                dummy.refresh_inventory_tab(enrich=True, filtered_only=True)

                ledger = json.loads(app.INVENTORY_LEDGER_PATH.read_text(encoding="utf-8"))["items"]
                by_person = {record["assigned_person"]: record for record in ledger}
                self.assertEqual(by_person["Kevin Hambone"]["best_company"], "Kevin Hambone Club")
                self.assertEqual(by_person["Kevin Hambone"]["estimated_payout"], 88)
                self.assertEqual(by_person["Lucas"]["best_company"], "Old")
                self.assertEqual(by_person["Lucas"]["estimated_payout"], 1)
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
            _profit_record_key = app.CardPipelineApp._profit_record_key
            _normalize_profit_record = app.CardPipelineApp._normalize_profit_record
            _load_profit_ledger = app.CardPipelineApp._load_profit_ledger
            _save_profit_ledger = app.CardPipelineApp._save_profit_ledger
            record_profit_sales = app.CardPipelineApp.record_profit_sales
            refresh_inventory_tab = lambda self: None
            refresh_profit_tab = lambda self: None

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
            _general_sold_sheet_name = app.CardPipelineApp._general_sold_sheet_name
            mark_inventory_record_sold = app.CardPipelineApp.mark_inventory_record_sold
            record_profit_sales = app.CardPipelineApp.record_profit_sales
            refresh_profit_tab = lambda self: None

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
                self.assertTrue(dummy.mark_inventory_record_sold(record, "Arena Club", 95))
                inventory = json.loads(app.INVENTORY_LEDGER_PATH.read_text(encoding="utf-8"))["items"]
                profit = json.loads(app.PROFIT_LEDGER_PATH.read_text(encoding="utf-8"))
                self.assertEqual(inventory, [])
                self.assertEqual(len(profit), 1)
                self.assertEqual(profit[0]["company"], "Arena Club")
                self.assertEqual(profit[0]["assigned_person"], "Hambone")
                self.assertEqual(profit[0]["sale_price"], 95.0)
                self.assertEqual(profit[0]["profit"], 55.0)
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

    def test_profit_period_filter_and_chart_series_support_daily_and_overall_views(self) -> None:
        class ProfitDummy:
            _money_value = app.CardPipelineApp._money_value
            _profit_record_date = app.CardPipelineApp._profit_record_date
            _profit_today = lambda self: datetime(2026, 6, 17).date()
            _profit_period_bounds = app.CardPipelineApp._profit_period_bounds
            _profit_period_label = app.CardPipelineApp._profit_period_label
            _profit_graph_label = app.CardPipelineApp._profit_graph_label
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
        dummy.profit_person_var = types.SimpleNamespace(get=lambda: "luc")
        dummy.profit_period_var = types.SimpleNamespace(get=lambda: "5 Days")
        dummy.profit_graph_var = types.SimpleNamespace(get=lambda: "Daily Trend")

        filtered = dummy._filtered_profit_records(rows)
        days, daily_values = dummy._profit_chart_series(filtered)
        dummy.profit_graph_var = types.SimpleNamespace(get=lambda: "Overall Profit")
        overall_days, overall_values = dummy._profit_chart_series(filtered)

        self.assertEqual([record["date_added"] for record in filtered], ["2026-06-17", "2026-06-13"])
        self.assertEqual(days, ["2026-06-13", "2026-06-14", "2026-06-15", "2026-06-16", "2026-06-17"])
        self.assertEqual(daily_values, [20, 0.0, 0.0, 0.0, 30])
        self.assertEqual(overall_days, days)
        self.assertEqual(overall_values, [20, 20.0, 20.0, 20.0, 50.0])

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


if __name__ == "__main__":
    unittest.main()
from openpyxl import Workbook, load_workbook
