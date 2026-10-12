#!/usr/bin/env python3
"""Write a large synthetic account panel and market-rate file for timing the import.

Usage: python tools/devbuild/make_large_panel.py [--accounts N] [--months M] [--out DIR]
Default: 50,000 accounts x 16 month ends (about 50 MB) in build/large_panel/.

The files follow docs/methodology/DATA_CONTRACT.md and import with no error:
every account is observed at every month end of the window, with a fixed
segment and currency, a positive balance, a rate within bounds and an opening
date before the window. Values come from a fixed seed, so the same arguments
always write the same bytes. The data is synthetic; it is for measuring speed
(import time and SHA-256 fingerprint), not for any model result.
"""
from __future__ import annotations

import argparse
import calendar
import random
from datetime import date
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
SEGMENTS = ("RET_TX", "RET_NTX", "WHS_NFC")
CURRENCIES = ("EUR", "USD")
ACCOUNT_HEADER = "as_of_date,account_id,segment,currency,balance,customer_rate,indexed,open_date,close_date"
RATE_HEADER = "as_of_date,rate_id,currency,tenor_months,rate"


def month_ends(first_year: int, first_month: int, count: int) -> list[str]:
    out = []
    year, month = first_year, first_month
    for _ in range(count):
        out.append(date(year, month, calendar.monthrange(year, month)[1]).isoformat())
        year, month = (year + 1, 1) if month == 12 else (year, month + 1)
    return out


def write(accounts: int, months: int, out: Path) -> tuple[Path, Path, str, str]:
    rng = random.Random(71)
    dates = month_ends(2024, 1, months)
    out.mkdir(parents=True, exist_ok=True)
    panel = out / "accounts_large.csv"
    rates = out / "market_rates_large.csv"
    profiles = []
    for i in range(1, accounts + 1):
        profiles.append((f"L{i:07d}", rng.choice(SEGMENTS), rng.choice(CURRENCIES), rng.uniform(100, 250000),
                         rng.choice(("", f"{rng.uniform(0, 0.03):.4f}")), rng.choice("01")))
    with panel.open("w", encoding="ascii", newline="\n") as f:
        f.write(ACCOUNT_HEADER + "\n")
        for as_of in dates:
            for account, segment, currency, base, rate, indexed in profiles:
                balance = base * rng.uniform(0.95, 1.05)
                f.write(f"{as_of},{account},{segment},{currency},{balance:.2f},{rate},{indexed},2019-01-01,\n")
    with rates.open("w", encoding="ascii", newline="\n") as f:
        f.write(RATE_HEADER + "\n")
        for as_of in dates:
            f.write(f"{as_of},EUR_1M,EUR,1,{rng.uniform(0, 0.04):.4f}\n")
            f.write(f"{as_of},USD_1M,USD,1,{rng.uniform(0, 0.05):.4f}\n")
    return panel, rates, dates[0], dates[-1]


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--accounts", type=int, default=50000)
    parser.add_argument("--months", type=int, default=16)
    parser.add_argument("--out", type=Path, default=ROOT / "build" / "large_panel")
    args = parser.parse_args()
    panel, rates, first, last = write(args.accounts, args.months, args.out)
    print(f"{panel} ({panel.stat().st_size / 1e6:.1f} MB)")
    print(f"{rates}")
    print(f"Data window: {first} to {last}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
