#!/usr/bin/env python3
"""Generate the synthetic decay fixtures (deterministic; standard library only).

Process (aggregate month-end balance of one segment and currency):

    ln B(t) = ln B(t-1) + mu + sigma * eps(t) + jump * 1[t = break]
    eps(t) ~ N(0, 1) i.i.d., drawn with random.Random(SEED).gauss

Files written next to this script:
    dec_log_short.csv        hand-chosen 8-month series (DEC-LOG-01)
    dec_series_ret_tx.csv    168-month generated series (DEC-E2E-01, backtest)
    dec_series_ret_tx.json   process, parameters and seed of the series above

Usage: python tests/fixtures/decay/generate_fixtures.py
"""
from __future__ import annotations

import calendar
import json
import math
import random
from datetime import date
from pathlib import Path

HERE = Path(__file__).resolve().parent
SEED = 20261008
PROCESS = {
    "segment": "RET_TX", "currency": "EUR", "start": "2012-01-31", "months": 168,
    "b0": 250000000.0, "mu": 0.0015, "sigma": 0.011,
    "break": "2020-03-31", "jump": 0.06,
}
SHORT = [("2025-01-31", 1000.0), ("2025-02-28", 1010.0), ("2025-03-31", 1005.0), ("2025-04-30", 1030.0),
         ("2025-05-31", 1100.0), ("2025-06-30", 1095.0), ("2025-07-31", 1102.0), ("2025-08-31", 1110.0)]


def month_end(year: int, month: int) -> date:
    return date(year, month, calendar.monthrange(year, month)[1])


def series() -> list[tuple[str, float]]:
    rng = random.Random(SEED)
    start = date.fromisoformat(PROCESS["start"])
    rows = []
    level = math.log(PROCESS["b0"])
    for i in range(PROCESS["months"]):
        y = start.year + (start.month - 1 + i) // 12
        m = (start.month - 1 + i) % 12 + 1
        d = month_end(y, m)
        if i > 0:
            level += PROCESS["mu"] + PROCESS["sigma"] * rng.gauss(0.0, 1.0)
            if d.isoformat() == PROCESS["break"]:
                level += PROCESS["jump"]
        rows.append((d.isoformat(), round(math.exp(level), 2)))
    return rows


def write(name: str, rows: list[tuple[str, float]]) -> None:
    lines = ["as_of_date,balance"] + [f"{d},{b:.2f}" for d, b in rows]
    (HERE / name).write_text("\n".join(lines) + "\n", encoding="ascii")


def main() -> int:
    write("dec_log_short.csv", SHORT)
    write("dec_series_ret_tx.csv", series())
    meta = dict(PROCESS, seed=SEED, generator="tests/fixtures/decay/generate_fixtures.py",
                rng="Python random.Random(seed).gauss, CPython 3.13",
                rounding="balances rounded to 2 decimals after exponentiation")
    (HERE / "dec_series_ret_tx.json").write_text(json.dumps(meta, indent=2) + "\n", encoding="ascii")
    print("fixtures written")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
