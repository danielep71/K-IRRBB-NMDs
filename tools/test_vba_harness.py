"""Check the VBA regression harness against the register and the contract, without Excel.

tests/modules/TEST_CoreCases.bas states its expected values as literal Expect*
calls. This module reads every one of them and re-derives the expected result
from the independent Python reading of docs/methodology/DATA_CONTRACT.md in
tools/test_data_contract.py and from the harness rules in
docs/methodology/TEST_CASES.md, so a wrong hand-written expectation fails here.
It also keeps the harness registry in tests/modules/TEST_Harness.bas equal to
the register in TEST_CASES.md, and the generated decay data module equal to the
expected files in tests/expected/decay/. It does not execute VBA.
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

HARNESS = ROOT / "tests" / "modules" / "TEST_Harness.bas"
CORE_CASES = ROOT / "tests" / "modules" / "TEST_CoreCases.bas"
CORE_PARSE = ROOT / "src" / "core" / "CORE_Parse.bas"
CORE_CODES = ROOT / "src" / "core" / "CORE_Codes.bas"
REGISTER = ROOT / "docs" / "methodology" / "TEST_CASES.md"
DECAY_CASES = ROOT / "tests" / "modules" / "TEST_DecayCases.bas"
DECAY_DATA = ROOT / "tests" / "modules" / "TEST_DecayData.bas"
DECAY_EXPECTED = ROOT / "tests" / "expected" / "decay"

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


def expectations() -> list[tuple[str, str, list[object]]]:
    """Every Expect* call as (case ID, helper, arguments); fails on a line it cannot read."""
    names = constants(CORE_PARSE)
    calls, case_id = [], None
    for number, line in enumerate(read(CORE_CASES).splitlines(), 1):
        text = line.strip()
        declared = re.fullmatch(r'Const C As String = "([A-Z0-9-]+)"', text)
        if declared:
            case_id = declared.group(1)
        if not text.startswith("Expect"):
            continue
        match = re.fullmatch(r"(Expect\w+) C, (.+)", text)
        if not match or case_id is None:
            raise ValueError(f"TEST_CoreCases.bas:{number}: unreadable expectation: {text}")
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
        core = select_cases(read(CORE_CASES), "Public Sub TEST_RunCoreCase")
        decay = select_cases(read(DECAY_CASES), "Public Sub TEST_RunDecayCase") if DECAY_CASES.exists() else set()
        self.assertFalse(core & decay, "a case is implemented twice")
        self.assertEqual(core | decay, built)
        self.assertEqual({case for case, _, _ in expectations()}, core)


def json_rows(value, prefix: str = "") -> list[tuple[str, float]]:
    """Flatten an expected-file value the way make_expected.py does (dict keys, [i] indices)."""
    if isinstance(value, dict):
        return [row for k, v in value.items() for row in json_rows(v, f"{prefix}.{k}" if prefix else k)]
    if isinstance(value, list):
        return [row for i, v in enumerate(value) for row in json_rows(v, f"{prefix}[{i}]")]
    return [(prefix, float(value))]


@unittest.skipUnless(DECAY_DATA.exists(), "decay cases not on this branch")
class DecayData(unittest.TestCase):
    """TEST_DecayData.bas is generated; it must equal the reviewed expected files exactly."""

    ROW = re.compile(r'v\(\d+, 1\) = "([^"]+)": v\(\d+, 2\) = "([^"]+)": v\(\d+, 3\) = Val\("([^"]+)"\): '
                     r'v\(\d+, 4\) = "([a-z]+)": v\(\d+, 5\) = Val\("([^"]+)"\)')

    def test_generated_module_equals_expected_files(self) -> None:
        import json
        vba = [(c, k, float(v), kind, float(e)) for c, k, v, kind, e in self.ROW.findall(read(DECAY_DATA))]
        files = []
        for path in sorted(DECAY_EXPECTED.glob("*.json")):
            case = json.loads(path.read_text(encoding="ascii"))
            self.assertEqual(case["case"], path.stem)
            for field in ("tool", "script", "method", "date", "produced_by"):
                self.assertTrue(case["provenance"].get(field), f"{path.name}: provenance.{field} missing")
            tol = case["tolerance"]
            files += [(case["case"], k, v, tol["kind"], float(tol["eps"])) for k, v in json_rows(case["expected"])]
        self.assertTrue(vba)
        self.assertEqual(sorted(vba), sorted(files))

    def test_every_decay_expected_case_is_registered(self) -> None:
        registered = set(register_ids())
        for path in DECAY_EXPECTED.glob("*.json"):
            self.assertIn(path.stem, registered)


if __name__ == "__main__":
    unittest.main()
