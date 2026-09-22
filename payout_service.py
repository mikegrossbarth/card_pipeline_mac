from __future__ import annotations

from intake_io import format_money


def payout_expenses_display(values: dict[str, object]) -> str:
    expense_offsets = float(values.get("unpaid_expenses") or 0.0)
    if abs(expense_offsets) < 0.005:
        return ""
    return f"-{format_money(abs(expense_offsets))}"
