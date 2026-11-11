#!/usr/bin/env python3
"""No VBA module may declare a name that is a reserved word of VBA.

Excel rejects a declaration such as `ByVal currency As String` (`Currency` is a
type keyword), and every call of that procedure then fails to compile with "Sub
or Function not defined"; LibreOffice accepts it, so the proxy run cannot catch
it. This test reads every module under src/, tests/ and examples/ and rejects a
reserved word declared as a parameter, variable, constant or array.
"""
from __future__ import annotations

import re
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]

# The VBA language reference's keywords that cannot name a variable, plus the
# type names that are keywords (Currency, Date, ...).
RESERVED = frozenset(word.lower() for word in """
    And As Boolean ByRef Byte ByVal Call Case CBool CByte CCur CDate CDbl CDec CInt CLng CLngLng
    CLngPtr Const CSng CStr Currency CVar CVErr Date Debug Decimal Declare DefBool DefByte DefCur
    DefDate DefDbl DefDec DefInt DefLng DefLngLng DefLngPtr DefObj DefSng DefStr DefVar Dim Do
    Double Each Else ElseIf Empty End EndIf Enum Eqv Erase Event Exit False For Friend Function Get
    Global GoSub GoTo If Imp Implements In Integer Is LBound Len LenB Let Like Long LongLong LongPtr
    Loop LSet Me Mod New Next Not Nothing Null Object On Option Optional Or ParamArray Preserve
    Private PtrSafe Public RaiseEvent ReDim Rem Resume Return RSet Select Set Single Static Step Stop
    String Sub Then To True Type TypeOf UBound Until Variant Wend While With WithEvents Xor
""".split())
# The word right after one of these is a declared name, unless it is itself a keyword that
# continues the declaration (Private Sub, Dim WithEvents, ReDim Preserve, ...).
LEADING = re.compile(r"\b(?:ByVal|ByRef|Optional|ParamArray|Dim|Static|Private|Public|Global|Const|ReDim)"
                     r"(?:\s+(?:Preserve|WithEvents|Const|Optional|ByVal|ByRef|ParamArray))*\s+([A-Za-z_]\w*)",
                     re.IGNORECASE)
# A further name in a list: `, name As` or `, name() As`, with an optional passing mode.
LISTED = re.compile(r",\s*(?:(?:ByVal|ByRef|Optional|ParamArray)\s+)*([A-Za-z_]\w*)\s*(?:\(\s*\))?\s+As\b",
                    re.IGNORECASE)
CONTINUING = frozenset({"sub", "function", "property", "type", "enum", "declare", "event", "const",
                        "withevents", "preserve", "optional", "byval", "byref", "paramarray"})


def code_of(line: str) -> str:
    """The line without string literals and without its comment."""
    without_strings = re.sub(r'"(?:[^"]|"")*"', '""', line)
    return without_strings.split("'", 1)[0]


def reserved_names(text: str) -> list[tuple[int, str]]:
    """(line number, word) for every reserved word declared as a name."""
    found = []
    for number, line in enumerate(text.replace("\r\n", "\n").split("\n"), 1):
        code = code_of(line)
        for match in list(LEADING.finditer(code)) + list(LISTED.finditer(code)):
            word = match.group(1).lower()
            if word in RESERVED and word not in CONTINUING:
                found.append((number, match.group(1)))
    return found


def modules() -> list[Path]:
    return sorted(path for folder in ("src", "tests", "examples") for pattern in ("*.bas", "*.cls", "*.frm")
                  for path in (ROOT / folder).rglob(pattern))


class ReservedWordTests(unittest.TestCase):
    def test_no_module_declares_a_reserved_word(self) -> None:
        self.assertTrue(modules())
        for path in modules():
            with self.subTest(module=str(path.relative_to(ROOT))):
                self.assertEqual(reserved_names(path.read_text(encoding="cp1252")), [])

    def test_the_scan_finds_reserved_names(self) -> None:
        cases = {
            "Private Function F(ByVal totals As Collection, _\n    ByVal currency As String) As Long": [(2, "currency")],
            "Public Sub G(ByVal a As Long, date As Date)": [(1, "date")],
            "    Dim string As String, i As Long": [(1, "string")],
            "    Dim x As Long, Next As Long": [(1, "Next")],
            "Private Const Currency As Long = 1": [(1, "Currency")],
            "    ReDim Preserve object(1 To 2)": [(1, "object")],
        }
        for code, expected in cases.items():
            with self.subTest(code=code):
                self.assertEqual(reserved_names(code), expected)

    def test_the_scan_leaves_valid_code_alone(self) -> None:
        for code in ('Private Function IndexOfTotal(ByVal totals As Collection, ByVal currencyCode As String) As Long',
                     '    mTotCurrency(mTotCount) = currencyCode',
                     'Private Const HEADER As String = "as_of_date,segment,currency"',
                     "    Dim key As String ' the currency key",
                     "Private Sub Withdraw(ByVal started As Date)",
                     "Public Function TEST_Summary() As String",
                     "    Dim WithEvents app As Application",
                     "    ReDim Preserve mTotDate(1 To 2 * UBound(mTotDate))"):
            with self.subTest(code=code):
                self.assertEqual(reserved_names(code), [])


if __name__ == "__main__":
    unittest.main()
