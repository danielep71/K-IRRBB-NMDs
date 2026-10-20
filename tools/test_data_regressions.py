"""Regression cases for the data-contract audit."""
from __future__ import annotations

import tempfile
import unittest
from collections.abc import Callable
from datetime import date
from pathlib import Path
from typing import Any

from test_data_contract import (
    ACCOUNT_FIELDS,
    RATE_FIELDS,
    read_accounts,
    read_market_rates,
    row_error,
    warnings_and_totals,
)


class DataRegressions(unittest.TestCase):
    row = dict(zip(ACCOUNT_FIELDS, ["2025-01-31", "A", "RET_TX", "EUR", "100.0049", "0.01", "0", "2020-01-01", ""]))

    def read(self, fields: list[str], rows: list[list[str]], reader: Callable[[Path], Any]) -> Any:
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "fixture.csv"
            path.write_text(",".join(fields) + "\n" + "\n".join(",".join(row) for row in rows) + "\n")
            return reader(path)

    def test_nonfinite_balance_and_unapproved_currency_fail(self) -> None:
        self.assertEqual(row_error({**self.row, "balance": "9" * 400}), "E04")
        self.assertEqual(row_error({**self.row, "currency": "ZZZ"}), "E07")

    def test_history_dates_checked_even_with_blank_later_date_and_unsorted_input(self) -> None:
        early = {**self.row, "close_date": "2025-02-15"}
        late = {**self.row, "as_of_date": "2025-03-31"}
        for rows in ((early, late), (late, early)):
            good, errors = self.read(ACCOUNT_FIELDS, [list(r.values()) for r in rows], read_accounts)
            self.assertEqual([error["code"] for error in errors], ["E10"])
            self.assertEqual([r["as_of_date"] for _, r in good], ["2025-01-31"])

    def test_contradictory_account_dates_fail(self) -> None:
        for field in ("open_date", "close_date"):
            rows = [{**self.row, field: "2020-01-01" if field == "open_date" else "2025-04-01"},
                    {**self.row, "as_of_date": "2025-02-28", field: "2021-01-01" if field == "open_date" else "2025-05-01"}]
            _, errors = self.read(ACCOUNT_FIELDS, [list(r.values()) for r in rows], read_accounts)
            self.assertEqual([e["code"] for e in errors], ["E11", "E11"])

    def test_segment_reconciliation_preserves_four_decimal_balances(self) -> None:
        rows = [(i, {**self.row, "account_id": str(i), "segment": segment})
                for i, segment in enumerate(("RET_TX", "RET_NTX", "WHS_NFC"))]
        _, totals = warnings_and_totals(rows, date(2025, 1, 31), date(2025, 1, 31))
        self.assertAlmostEqual(sum(t["balance"] for t in totals), 300.0147, places=8)

    def test_market_missing_id_extra_column_and_series_changes_fail(self) -> None:
        valid = ["2025-01-31", "R", "EUR", "1", "0.01"]
        for row, expected in (([valid[0], "", *valid[2:]], "E03"),
                              ([*valid, "extra"], "E01"),
                              (["2025-02-28", "R", "USD", "1", "0.01"], "E08"),
                              (["2025-02-28", "R", "EUR", "12", "0.01"], "E08")):
            with self.subTest(row=row):
                self.assertEqual(self.read(RATE_FIELDS, [valid, row], read_market_rates),
                                 [{"line": 3, "code": expected}])


if __name__ == "__main__":
    unittest.main()
