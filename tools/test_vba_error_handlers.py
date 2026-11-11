#!/usr/bin/env python3
"""No VBA error handler may set a new handler before it has ended with Resume.

While a handler runs, VBA does not trap a second error: an `On Error GoTo`
executed inside the handler has no effect until a `Resume` ends it, so an
error in the cleanup escapes to the caller and the first error is lost. The
harness and the importer end their handler with `Resume` before cleaning up
(docs/WORKBOOK_CONTRACT.md: primary and cleanup errors are both kept). This
test reads every module under src/, tests/ and examples/.
"""
from __future__ import annotations

import re
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
TARGET = re.compile(r"^\s*On Error GoTo (\w+)\s*(?:'.*)?$", re.IGNORECASE)
LABEL = re.compile(r"^(\w+):\s*(?:'.*)?$")
END = re.compile(r"^\s*End (?:Sub|Function|Property)\b", re.IGNORECASE)
RESUME = re.compile(r"^\s*Resume\b", re.IGNORECASE)


def nested_handlers(text: str) -> list[tuple[int, str]]:
    """(line number, handler) for each handler that sets a new handler before any Resume."""
    lines = text.replace("\r\n", "\n").split("\n")
    found = []
    targets: set[str] = set()
    current: str | None = None
    resumed = False
    for number, line in enumerate(lines, 1):
        if END.match(line):
            targets, current = set(), None
            continue
        label = LABEL.match(line)
        if label:
            current = label.group(1) if label.group(1).lower() in targets else None
            resumed = False
            continue
        target = TARGET.match(line)
        if current is not None:
            if RESUME.match(line):
                resumed = True
            elif target and target.group(1) != "0" and not resumed:
                found.append((number, current))
        if target and target.group(1) != "0":
            targets.add(target.group(1).lower())
    return found


class ErrorHandlerTests(unittest.TestCase):
    def test_no_handler_sets_a_new_handler_before_resume(self) -> None:
        paths = sorted(path for folder in ("src", "tests", "examples") for pattern in ("*.bas", "*.cls")
                       for path in (ROOT / folder).rglob(pattern))
        self.assertTrue(paths)
        for path in paths:
            with self.subTest(module=str(path.relative_to(ROOT))):
                self.assertEqual(nested_handlers(path.read_text(encoding="cp1252")), [])

    def test_the_scan(self) -> None:
        broken = ("Sub A()\n    On Error GoTo Fail\n    X\n    Exit Sub\nFail:\n    On Error GoTo CleanupFail\n"
                  "    Y\n    Exit Sub\nCleanupFail:\n    Z\nEnd Sub")
        self.assertEqual(nested_handlers(broken), [(6, "Fail")])
        fixed = ("Sub A()\n    On Error GoTo Fail\n    X\n    Exit Sub\nFail:\n    Resume Report\nReport:\n"
                 "    On Error GoTo CleanupFail\n    Y\n    Exit Sub\nCleanupFail:\n    Resume Done\nDone:\n"
                 "    On Error Resume Next\nEnd Sub")
        self.assertEqual(nested_handlers(fixed), [])


if __name__ == "__main__":
    unittest.main()
