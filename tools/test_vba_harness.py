"""Check the VBA regression harness against the register and the contract, without Excel.

tests/modules/TEST_CoreCases.bas states its expected values as literal Expect*
calls. This module reads every one of them and re-derives the expected result
from the independent Python reading of docs/methodology/DATA_CONTRACT.md in
tools/test_data_contract.py and from the harness rules in
docs/methodology/TEST_CASES.md, so a wrong hand-written expectation fails here.
It also keeps the harness registry in tests/modules/TEST_Harness.bas equal to
the register in TEST_CASES.md. It does not execute VBA.
"""
from __future__ import annotations

import math
import re
import sys
import unittest
from datetime import date
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))
import test_data_contract as reference  # noqa: E402

sys.path.insert(0, str(ROOT / "tools" / "workbook"))
import build_template  # noqa: E402

sys.path.insert(0, str(ROOT / "tools" / "devbuild"))
import fixture_module  # noqa: E402

HARNESS = ROOT / "tests" / "modules" / "TEST_Harness.bas"
CORE_CASES = ROOT / "tests" / "modules" / "TEST_CoreCases.bas"
WORKBOOK_CASES = ROOT / "tests" / "modules" / "TEST_WorkbookCases.bas"
DATA_CASES = ROOT / "tests" / "modules" / "TEST_DataCases.bas"
HASH_CASES = ROOT / "tests" / "modules" / "TEST_HashCases.bas"
FIXTURE_MODULE = ROOT / "tests" / "modules" / "TEST_Fixtures.bas"
CORE_PARSE = ROOT / "src" / "core" / "CORE_Parse.bas"
CORE_CODES = ROOT / "src" / "core" / "CORE_Codes.bas"
REGISTER = ROOT / "docs" / "methodology" / "TEST_CASES.md"

VB_OBJECT_ERROR = -2147221504
ARGUMENT = re.compile(r'\s*("(?:[^"]|"")*"|[^,]+?)\s*(?:,|$)')
CONSTANT = re.compile(r"^Public Const (\w+) As \w+ = (.+?)\s*(?:'.*)?$", re.MULTILINE)


def read(path: Path) -> str:
    """VBA sources are cp1252 (as the VBA editor stores them); documentation is UTF-8."""
    return path.read_text(encoding="cp1252" if path.suffix == ".bas" else "utf-8")


def constants(*paths: Path) -> dict[str, object]:
    found: dict[str, object] = {}
    for path in paths:
        for name, literal in CONSTANT.findall(read(path)):
            found[name] = literal_value(literal, {})
    return found


def literal_value(token: str, names: dict[str, object]) -> object:
    """Value of one VBA literal argument: string, number (with optional #), Boolean or constant."""
    if token.startswith('"'):
        return token[1:-1].replace('""', '"')
    if token in ("True", "False"):
        return token == "True"
    if token in names:
        return names[token]
    number = token.rstrip("#")
    if re.fullmatch(r"-?\d+", number):
        return int(number)
    if re.fullmatch(r"-?(\d+\.?\d*|\.\d+)([Ee][+-]?\d+)?", number):
        return float(number)
    raise ValueError(f"not a literal: {token!r}")


def expectations(path: Path = CORE_CASES) -> list[tuple[str, str, list[object]]]:
    """Every Expect* call as (case ID, helper, arguments); fails on a line it cannot read."""
    names = constants(CORE_PARSE)
    calls, case_id = [], None
    for number, line in enumerate(read(path).splitlines(), 1):
        text = line.strip()
        declared = re.fullmatch(r'Const C As String = "([A-Z0-9-]+)"', text)
        if declared:
            case_id = declared.group(1)
        if not text.startswith("Expect"):
            continue
        match = re.fullmatch(r"(Expect\w+) C, (.+)", text)
        if not match or case_id is None:
            raise ValueError(f"{path.name}:{number}: unreadable expectation: {text}")
        tokens = [m.group(1) for m in ARGUMENT.finditer(match.group(2)) if m.group(0)]
        calls.append((case_id, match.group(1), [literal_value(t, names) for t in tokens]))
    return calls


def outcome(built: bool, errored: bool, asserts: int, passed: int) -> str:
    """TEST_CASES.md harness rule: unbuilt NOT RUN, then ERROR, then PASS only on all of at least one."""
    if not built:
        return "NOT RUN"
    if errored:
        return "ERROR"
    return "PASS" if asserts > 0 and passed == asserts else "FAIL"


def within(actual: float, expected: float, kind: str, eps: float) -> bool:
    """TEST_CASES.md tolerance kinds; any other kind is rejected."""
    diff = abs(actual - expected)
    if kind == "abs":
        return diff <= eps
    if kind == "rel":
        return diff <= eps * max(1.0, abs(expected))
    if kind == "exact":
        return actual == expected
    return False


def parse_status(text: str, accepted: bool, ok: int, missing: int, invalid: int) -> int:
    if text == "":
        return missing
    return ok if accepted else invalid


def register_ids() -> list[str]:
    section = read(REGISTER).split('<a id="register"></a>', 1)[1].split("## ➕", 1)[0]
    return re.findall(r"^\| `([A-Z0-9-]+)` \|", section, re.MULTILINE)


def registry() -> list[tuple[str, bool]]:
    return [(case, built == "True")
            for case, built in re.findall(r'^\s*c\.Add Array\("([A-Z0-9-]+)", (True|False),', read(HARNESS), re.MULTILINE)]


def select_cases(text: str, procedure: str) -> set[str]:
    body = text.split(procedure, 1)[1].split("End Select", 1)[0]
    return set(re.findall(r'"([A-Z0-9-]+)"', "".join(re.findall(r"^\s*Case (\"[^\n]*)", body, re.MULTILINE))))


class Expectations(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.calls = expectations()
        cls.codes = constants(CORE_PARSE)
        cls.ok, cls.missing, cls.invalid = (cls.codes[f"CORE_PARSE_{k}"] for k in ("OK", "MISSING", "INVALID"))

    def of(self, helper: str) -> list[tuple[str, list[object]]]:
        found = [(case, args) for case, name, args in self.calls if name == helper]
        self.assertTrue(found, f"no {helper} expectations")
        return found

    def test_every_expect_line_is_read_and_every_helper_is_checked(self) -> None:
        lines = [line for line in read(CORE_CASES).splitlines() if line.strip().startswith("Expect")]
        self.assertEqual(len(self.calls), len(lines))
        checked = {"ExpectOutcome", "ExpectWithin", "ExpectDate", "ExpectMonthEnd", "ExpectDecimal",
                   "ExpectRate", "ExpectSegment", "ExpectCurrency"}
        self.assertEqual({name for _, name, _ in self.calls}, checked)

    def test_result_codes_are_distinct(self) -> None:
        self.assertEqual(len({self.ok, self.missing, self.invalid}), 3)

    def test_harness_outcomes(self) -> None:
        seen = set()
        for case, (built, errored, asserts, passed, expected) in self.of("ExpectOutcome"):
            with self.subTest(case=case, built=built, errored=errored, asserts=asserts, passed=passed):
                self.assertEqual(outcome(built, errored, asserts, passed), expected)
                seen.add(expected)
        self.assertEqual(seen, {"NOT RUN", "ERROR", "FAIL", "PASS"})

    def test_tolerance_rules(self) -> None:
        kinds = set()
        for case, (expected, actual, kind, eps, want) in self.of("ExpectWithin"):
            with self.subTest(case=case, expected=expected, actual=actual, kind=kind, eps=eps):
                self.assertEqual(within(float(actual), float(expected), kind, float(eps)), want)
                kinds.add((kind, want))
        for kind in ("abs", "rel", "exact"):
            self.assertIn((kind, True), kinds)
            self.assertIn((kind, False), kinds)

    def test_dates(self) -> None:
        for case, (text, status, y, m, d) in self.of("ExpectDate"):
            with self.subTest(case=case, text=text):
                parsed = reference.parse_date(text)
                self.assertEqual(status, parse_status(text, parsed is not None, self.ok, self.missing, self.invalid))
                if status == self.ok:
                    self.assertEqual(parsed, date(y, m, d))

    def test_month_ends(self) -> None:
        for case, (y, m, d, want) in self.of("ExpectMonthEnd"):
            with self.subTest(case=case, day=(y, m, d)):
                self.assertEqual(reference.is_month_end(date(y, m, d)), want)

    def test_decimals(self) -> None:
        for case, (text, places, status, value) in self.of("ExpectDecimal"):
            with self.subTest(case=case, text=text):
                self.assertEqual(places, 4, "the reference covers balances (4 places) only")
                accepted = reference.BALANCE.fullmatch(text) is not None and math.isfinite(float(text))
                self.assertEqual(status, parse_status(text, accepted, self.ok, self.missing, self.invalid))
                if status == self.ok:
                    self.assertEqual(float(value), float(text))

    def test_rates(self) -> None:
        low, high = reference.RATE_BOUNDS
        for case, (text, status, value) in self.of("ExpectRate"):
            with self.subTest(case=case, text=text):
                accepted = reference.RATE.fullmatch(text) is not None and low <= float(text) <= high
                self.assertEqual(status, parse_status(text, accepted, self.ok, self.missing, self.invalid))
                if status == self.ok:
                    self.assertEqual(float(value), float(text))

    def test_every_parser_covers_ok_missing_and_invalid(self) -> None:
        position = {"ExpectDate": 1, "ExpectDecimal": 2, "ExpectRate": 1}
        for helper, index in position.items():
            with self.subTest(helper=helper):
                self.assertEqual({args[index] for _, args in self.of(helper)}, {self.ok, self.missing, self.invalid})

    def test_codes(self) -> None:
        for case, (code, want) in self.of("ExpectSegment"):
            with self.subTest(case=case, segment=code):
                self.assertEqual(code in reference.SEGMENTS, want)
        for case, (code, want) in self.of("ExpectCurrency"):
            with self.subTest(case=case, currency=code):
                self.assertEqual(code in reference.SUPPORTED_CURRENCIES, want)


class SourceConstants(unittest.TestCase):
    def test_vba_constants_match_the_reference(self) -> None:
        values = constants(CORE_PARSE, CORE_CODES)
        self.assertEqual((values["CORE_RATE_MIN"], values["CORE_RATE_MAX"]), reference.RATE_BOUNDS)
        self.assertIn("{1,%d}" % values["CORE_RATE_DECIMALS"], reference.RATE.pattern)
        self.assertEqual(set(values["CORE_CURRENCY_ALLOWLIST"].split(",")), reference.SUPPORTED_CURRENCIES)
        segments = re.search(r"^\s*Case (\"[^\n]*)", read(CORE_CODES), re.MULTILINE).group(1)
        self.assertEqual(set(re.findall(r'"([A-Z_]+)"', segments)), reference.SEGMENTS)

    def test_no_case_error_is_vb_object_error_plus_1000(self) -> None:
        self.assertEqual(constants(HARNESS)["TEST_ERR_NO_CASE"], VB_OBJECT_ERROR + 1000)


class Registry(unittest.TestCase):
    def test_registry_lists_the_register_in_order(self) -> None:
        self.assertEqual([case for case, _ in registry()], register_ids())

    def test_built_cases_are_dispatched_and_implemented(self) -> None:
        built = {case for case, flag in registry() if flag}
        self.assertTrue(built)
        self.assertEqual(select_cases(read(HARNESS), "Private Sub DispatchCase"), built)
        implemented = set()
        for path in sorted(HARNESS.parent.glob("TEST_*Cases.bas")):
            entry = re.search(r"Public Sub (TEST_Run\w+Case)\(", read(path)).group(1)
            implemented |= select_cases(read(path), f"Public Sub {entry}")
        self.assertEqual(implemented, built)
        declared = set(re.findall(r'Const C As String = "([A-Z0-9-]+)"',
                                  "".join(read(p) for p in HARNESS.parent.glob("TEST_*Cases.bas"))))
        self.assertEqual(declared, built)


class HashCases(unittest.TestCase):
    """Every SHA-256 literal in TEST_HashCases agrees with Python's hashlib."""

    @classmethod
    def setUpClass(cls) -> None:
        cls.calls = expectations(HASH_CASES)

    def test_every_expect_line_is_read(self) -> None:
        lines = [line for line in read(HASH_CASES).splitlines() if line.strip().startswith("Expect")]
        self.assertEqual(len(self.calls), len(lines))
        self.assertEqual({name for _, name, _ in self.calls}, {"ExpectSha", "ExpectShaRepeated", "ExpectShaPieces"})

    def test_digests(self) -> None:
        import hashlib
        for _, helper, args in self.calls:
            with self.subTest(helper=helper, args=str(args)[:60]):
                if helper == "ExpectSha":
                    message, expected = args
                elif helper == "ExpectShaRepeated":
                    unit, count, expected = args
                    message = unit * count
                else:
                    message, _, expected = args
                self.assertEqual(hashlib.sha256(message.encode("latin-1")).hexdigest(), expected)

    def test_published_examples_are_present(self) -> None:
        published = {"ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad",
                     "248d6a61d20638b8e5c026930c3e6039a33ce45964ff2167f6ecedd419db06c1",
                     "cf5b16a778af8380036ce59e7b0492370b249b11e8f07a51afac45037afee9d1"}
        self.assertTrue(published <= {args[-1] for _, _, args in self.calls})


class DataCases(unittest.TestCase):
    """TEST_DataCases literals agree with the Python reading of DATA_CONTRACT.md; TEST_Fixtures is generated."""

    @classmethod
    def setUpClass(cls) -> None:
        cls.calls = expectations(DATA_CASES)

    def args(self, helper: str) -> list[list[object]]:
        found = [args for _, name, args in self.calls if name == helper]
        self.assertTrue(found, helper)
        return found

    @staticmethod
    def findings(reader, text: str) -> str:
        import tempfile
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "input.csv"
            path.write_text(text.replace("|", "\n") + "\n", encoding="ascii")
            result = reader(path)
        errors = result[1] if isinstance(result, tuple) else result
        return ",".join(f'{e["line"]}:{e["code"]}' for e in errors)

    def test_fixture_module_is_generated_from_the_fixtures(self) -> None:
        self.assertEqual(read(FIXTURE_MODULE).replace("\r\n", "\n"), fixture_module.module_text(),
                         "regenerate with: python tools/devbuild/fixture_module.py")

    def test_every_expect_line_is_read(self) -> None:
        lines = [line for line in read(DATA_CASES).splitlines() if line.strip().startswith("Expect")]
        self.assertEqual(len(self.calls), len(lines))
        self.assertEqual({name for _, name, _ in self.calls},
                         {"ExpectAccountFindings", "ExpectRateFindings", "ExpectMissingExtract", "ExpectWindowGaps",
                          "ExpectOutcomes"})

    def test_account_findings(self) -> None:
        for text, expected in self.args("ExpectAccountFindings"):
            with self.subTest(text=text[:80]):
                self.assertEqual(self.findings(reference.read_accounts, text), expected)

    def test_rate_findings(self) -> None:
        for text, expected in self.args("ExpectRateFindings"):
            with self.subTest(text=text[:80]):
                self.assertEqual(self.findings(reference.read_market_rates, text), expected)

    def test_window_gaps(self) -> None:
        import tempfile
        for text, first, last, expected in self.args("ExpectWindowGaps"):
            with self.subTest(window=(first, last)), tempfile.TemporaryDirectory() as directory:
                path = Path(directory) / "input.csv"
                path.write_text(text.replace("|", "\n") + "\n", encoding="ascii")
                rows, errors = reference.read_accounts(path)
                self.assertEqual(errors, [])
                gaps = reference.missing_extracts(rows, reference.parse_date(first), reference.parse_date(last))
                self.assertEqual(",".join(g["as_of_date"] for g in gaps), expected)

    def test_missing_extract(self) -> None:
        import json
        expected_file = {"accounts.csv": "accounts_summary.json", "availability.csv": "availability_outcomes.json"}
        for fixture, month_end in self.args("ExpectMissingExtract"):
            with self.subTest(fixture=fixture, month_end=month_end):
                spec = json.loads((ROOT / "tests/expected/data_contract" / expected_file[fixture])
                                  .read_text(encoding="utf-8"))
                first = reference.parse_date(spec["window"]["first"])
                last = reference.parse_date(spec["window"]["last"])
                rows, _ = reference.read_accounts(ROOT / "tests/fixtures/data_contract" / fixture)
                kept = [(n, r) for n, r in rows if r["as_of_date"] != month_end]
                self.assertEqual(reference.missing_extracts(kept, first, last),
                                 [{"code": "E12", "as_of_date": month_end}])

    def test_outcomes(self) -> None:
        import tempfile
        for text, first, last, expected in self.args("ExpectOutcomes"):
            with self.subTest(window=(first, last), text=text[:60]), tempfile.TemporaryDirectory() as directory:
                path = Path(directory) / "input.csv"
                path.write_text(text.replace("|", "\n") + "\n", encoding="ascii")
                rows, _ = reference.read_accounts(path)
                outcomes = reference.transition_outcomes(rows, reference.parse_date(first), reference.parse_date(last))
                self.assertEqual(",".join(f'{o["account_id"]}|{o["origin"]}|{o["segment"]}|{o["outcome"]}'
                                          for o in outcomes), expected)


class WorkbookLayout(unittest.TestCase):
    """WB-LAYOUT-01 literals equal the template generator's sheet and name contract."""

    @classmethod
    def setUpClass(cls) -> None:
        cls.calls = expectations(WORKBOOK_CASES)

    def args(self, helper: str) -> list[list[object]]:
        return [args for _, name, args in self.calls if name == helper]

    def test_every_expect_line_is_read(self) -> None:
        lines = [line for line in read(WORKBOOK_CASES).splitlines() if line.strip().startswith("Expect")]
        self.assertEqual(len(self.calls), len(lines))
        self.assertEqual({name for _, name, _ in self.calls}, {"ExpectCounts", "ExpectSheet", "ExpectName"})

    def test_sheets_and_code_names(self) -> None:
        self.assertEqual(self.args("ExpectSheet"),
                         [[i, s.name, s.code_name] for i, s in enumerate(build_template.SHEETS, 1)])

    def test_defined_names(self) -> None:
        self.assertEqual({name: (sheet, ref) for name, sheet, ref in self.args("ExpectName")}, build_template.NAMES)
        self.assertEqual(len(self.args("ExpectName")), len(build_template.NAMES))

    def test_counts(self) -> None:
        self.assertEqual(self.args("ExpectCounts"), [[len(build_template.SHEETS), len(build_template.NAMES)]])


if __name__ == "__main__":
    unittest.main()
