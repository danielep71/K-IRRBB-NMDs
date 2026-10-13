#!/usr/bin/env python3
"""Generate tests/modules/TEST_Fixtures.bas from the data-contract fixtures and expected files.

Usage: python tools/devbuild/fixture_module.py [--check]

The VBA data cases need the synthetic CSV fixtures and their hand-computed
expectations inside the workbook, without file paths that differ between hosts.
This script writes them as string literals, unchanged: fixture lines as they are
in the file, expected findings, warnings and totals as "|"-separated fields in
the order the contract sorts them. tools/test_vba_harness.py fails if the
committed module differs from this output. --check compares without writing.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
FIXTURES = ROOT / "tests" / "fixtures" / "data_contract"
EXPECTED = ROOT / "tests" / "expected" / "data_contract"
MODULE = ROOT / "tests" / "modules" / "TEST_Fixtures.bas"
CSV_FILES = ["accounts.csv", "accounts_invalid.csv", "market_rates.csv", "availability.csv"]


def expected_lists() -> dict[str, list[str]]:
    summary = json.loads((EXPECTED / "accounts_summary.json").read_text(encoding="utf-8"))
    invalid = json.loads((EXPECTED / "accounts_invalid_findings.json").read_text(encoding="utf-8"))
    availability = json.loads((EXPECTED / "availability_outcomes.json").read_text(encoding="utf-8"))
    visible = [f"{origin}|{account}|{closed}" for origin, accounts in sorted(availability["visible_close_dates"].items())
               for account, closed in sorted(accounts.items())]
    warnings = sorted(summary["warnings"], key=lambda w: (w["code"], w["account_id"], w["as_of_date"]))
    totals = sorted(summary["totals"], key=lambda t: (t["as_of_date"], t["segment"], t["currency"]))
    return {
        "accounts_summary.window": [summary["window"]["first"], summary["window"]["last"]],
        "accounts_summary.counts": [str(summary["rows"]), str(summary["accounts"])],
        "accounts_summary.tolerance": [f'{summary["balance_tolerance"]:.4f}'],
        "accounts_summary.errors": [f'{e["line"]}|{e["code"]}' for e in summary["errors"]],
        "accounts_summary.warnings": [f'{w["code"]}|{w["account_id"]}|{w["as_of_date"]}' for w in warnings],
        "accounts_summary.totals": [f'{t["as_of_date"]}|{t["segment"]}|{t["currency"]}|{t["accounts"]}|'
                                    f'{t["balance"]:.4f}' for t in totals],
        "accounts_invalid_findings.errors": [f'{e["line"]}|{e["code"]}' for e in invalid["errors"]],
        "availability_outcomes.window": [availability["window"]["first"], availability["window"]["last"]],
        "availability_outcomes.counts": [str(availability["rows"])],
        "availability_outcomes.outcomes": [f'{o["account_id"]}|{o["origin"]}|{o["segment"]}|{o["outcome"]}'
                                           for o in sorted(availability["outcomes"],
                                                           key=lambda o: (o["account_id"], o["origin"]))],
        "availability_outcomes.warnings": [f'{w["code"]}|{w["account_id"]}|{w["as_of_date"]}' for w in sorted(
            availability["warnings"], key=lambda w: (w["code"], w["account_id"], w["as_of_date"]))],
        "availability_outcomes.visible": visible,
        "availability_outcomes.origins": sorted(availability["visible_close_dates"]),
        # Independent fingerprints (Python hashlib) of the fixture files as committed.
        "sha256.accounts.csv": [hashlib.sha256((FIXTURES / "accounts.csv").read_bytes()).hexdigest()],
        "sha256.market_rates.csv": [hashlib.sha256((FIXTURES / "market_rates.csv").read_bytes()).hexdigest()],
    }


def fixture_lines(name: str) -> list[str]:
    text = (FIXTURES / name).read_bytes().decode("ascii")
    lines = text.split("\n")
    if lines and lines[-1] == "":
        lines.pop()
    return lines


def _literal(text: str) -> str:
    if any(ord(ch) < 32 or ord(ch) > 126 for ch in text) or len(text) > 900:
        raise ValueError(f"cannot write {text!r} as a VBA literal")
    return '"' + text.replace('"', '""') + '"'


def _function(name: str, items: list[str]) -> list[str]:
    out = [f"Private Function {name}() As Variant"]
    if not items:
        out += [f"    {name} = Array()", "End Function"]
        return out
    out += [f"    Dim a(1 To {len(items)}) As String"]
    out += [f"    a({i}) = {_literal(item)}" for i, item in enumerate(items, 1)]
    out += [f"    {name} = a", "End Function"]
    return out


def _function_name(key: str) -> str:
    return "Data_" + "".join(part.capitalize() for part in key.replace(".", "_").replace("-", "_").split("_"))


def _dispatcher(public: str, entries: dict[str, str]) -> list[str]:
    out = [f"Public Function {public}(ByVal name As String) As Variant", "    Select Case name"]
    for key, function in entries.items():
        out.append(f'        Case "{key}": {public} = {function}()')
    out += ["        Case Else",
            f'            Err.Raise 5, "TEST_Fixtures.{public}", "No fixture data named " & name & "."',
            "    End Select", "End Function"]
    return out


def module_text() -> str:
    fixtures = {name: _function_name(name) for name in CSV_FILES}
    expected = expected_lists()
    names = {key: _function_name(key) for key in expected}
    lines = [
        'Attribute VB_Name = "TEST_Fixtures"',
        "'" + "=" * 78,
        "' TEST_Fixtures",
        "'" + "-" * 78,
        "' PURPOSE",
        "'   GENERATED by tools/devbuild/fixture_module.py from",
        "'   tests/fixtures/data_contract and tests/expected/data_contract; do not edit.",
        "'   Fixture lines are the file's lines without line endings; expected",
        "'   findings, warnings and totals are \"|\"-separated, in contract order.",
        "'   Each function returns a 1-based String array (an empty Variant array",
        "'   when there is nothing to expect).",
        "'" + "=" * 78,
        "Option Explicit",
        "Option Private Module",
        "",
        "",
    ]
    lines += _dispatcher("TEST_FixtureLines", fixtures) + ["", ""]
    lines += _dispatcher("TEST_ExpectedLines", names)
    for name, function in fixtures.items():
        lines += ["", ""] + _function(function, fixture_lines(name))
    for key, function in names.items():
        lines += ["", ""] + _function(function, expected[key])
    return "\n".join(lines) + "\n"


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--check", action="store_true", help="compare with the committed module; write nothing")
    args = parser.parse_args()
    text = module_text()
    if args.check:
        same = MODULE.is_file() and MODULE.read_text(encoding="ascii").replace("\r\n", "\n") == text
        print("TEST_Fixtures.bas matches its sources" if same else "TEST_Fixtures.bas differs: regenerate it")
        return 0 if same else 1
    MODULE.write_bytes(text.encode("ascii"))
    print("wrote", MODULE.relative_to(ROOT))
    return 0


if __name__ == "__main__":
    sys.exit(main())
