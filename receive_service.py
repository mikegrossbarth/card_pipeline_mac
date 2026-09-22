from __future__ import annotations

from pathlib import Path
from typing import Iterable, Protocol


class ReceiveRow(Protocol):
    excel_row: int


def receive_row_ref_key(match: dict[str, object]) -> str:
    sheet_name = str(match.get("sheet") or "").strip()
    workbook_sheet = str(match.get("workbook_sheet") or "").strip()
    try:
        workbook_row = int(match.get("workbook_row") or 0)
    except (TypeError, ValueError):
        workbook_row = 0
    if not sheet_name or not workbook_sheet or workbook_row <= 0:
        return ""
    return f"raw:{sheet_name.lower()}:{workbook_sheet.lower()}:{workbook_row}"


def receive_row_ref(row: object, sheet_source: object = "") -> tuple[str, str, int] | None:
    sheet_name = str(getattr(row, "_receive_sheet", "") or sheet_source or "").strip()
    workbook_sheet = str(getattr(row, "_receive_workbook_sheet", "") or "").strip()
    try:
        workbook_row = int(getattr(row, "_receive_workbook_row", 0) or 0)
    except (TypeError, ValueError):
        workbook_row = 0
    if not sheet_name or not workbook_sheet or workbook_row <= 0:
        return None
    return (Path(sheet_name).name, workbook_sheet, workbook_row)


def receive_row_is_sheet_matched(row: object, sheet_source: object = "") -> bool:
    return receive_row_ref(row, sheet_source) is not None


def receive_rows_missing_sheet_refs(
    rows: Iterable[ReceiveRow],
    sheet_sources: dict[int, str],
) -> list[ReceiveRow]:
    missing: list[ReceiveRow] = []
    for row in rows:
        if receive_row_ref(row, sheet_sources.get(row.excel_row, "")) is None:
            missing.append(row)
    return missing


def receive_target_rows(rows: Iterable[ReceiveRow]) -> list[ReceiveRow]:
    return [
        row
        for row in rows
        if str(getattr(row, "cert_number", "") or "").strip()
        or str(getattr(row, "_receive_sheet", "") or "").strip()
    ]
