#!/usr/bin/env python3
"""No VBA module may spell a floating-point number with more than 15 significant digits.

The VBE keeps a Double literal to 15 significant digits: it rewrites
`-0.007784894002430293` as `-7.78489400243029E-03`, a slightly different value, so
the workbook would not hold the source's number and the source round trip
(`compare_vba_export.py`) fails. LibreOffice keeps every digit, so the proxy run
cannot catch it. This test reads every module under src/, tests/ and examples/.
"""
from __future__ import annotations

import re
import unittest
from pathlib import Path

from compare_vba_export import segments

ROOT = Path(__file__).resolve().parents[1]
FLOAT = re.compile(r"(?<![\w.])((?:\d+\.\d*|\.\d+)(?:e[+-]?\d+)?|\d+e[+-]?\d+)([#!@]?)(?![\w.])", re.IGNORECASE)
MAX_DIGITS = 15


def significant_digits(text: str) -> int:
    mantissa = text.lower().split("e", 1)[0].replace(".", "").lstrip("0").rstrip("0")
    return len(mantissa)


def long_literals(text: str) -> list[tuple[int, str]]:
    """(line number, literal) for every number in code with more than 15 significant digits."""
    found = []
    for number, line in enumerate(text.replace("\r\n", "\n").split("\n"), 1):
        for kind, part in segments(line):
            if kind == "code":
                found += [(number, m.group(1)) for m in FLOAT.finditer(part)
                          if significant_digits(m.group(1)) > MAX_DIGITS]
    return found


class DoubleLiteralTests(unittest.TestCase):
    def test_no_module_spells_more_than_15_significant_digits(self) -> None:
        paths = sorted(path for folder in ("src", "tests", "examples") for pattern in ("*.bas", "*.cls", "*.frm")
                       for path in (ROOT / folder).rglob(pattern))
        self.assertTrue(paths)
        for path in paths:
            with self.subTest(module=str(path.relative_to(ROOT))):
                self.assertEqual(long_literals(path.read_text(encoding="cp1252")), [])

    def test_the_scan(self) -> None:
        self.assertEqual(long_literals("    x = -0.007784894002430293 * q"), [(1, "0.007784894002430293")])
        self.assertEqual(long_literals("    x = 1.2345678901234567E-05"), [(1, "1.2345678901234567E-05")])
        for code in ("    x = -0.00778489400243029 * q", "    x = -7.78489400243029E-03",
                     "    x = 0.32246712907004", "    x = 1234567890123456789",
                     '    s = "0.007784894002430293"', "    x = 1 ' 0.007784894002430293",
                     "    x = 100000000000#", "    y = x1234567890123456.5"):
            with self.subTest(code=code):
                self.assertEqual(long_literals(code), [])


if __name__ == "__main__":
    unittest.main()
