"""Regression fixtures for the duplicate Public name rule of check_source."""
from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from check_source import check_public_names, public_members


def module(name: str, body: str) -> str:
    return f'Attribute VB_Name = "{name}"\nOption Explicit\n{body}'


class PublicMembers(unittest.TestCase):
    def test_every_kind_of_public_member_is_found(self) -> None:
        body = ("Public Const LIMIT As Long = 3\n"
                "Public Total As Double, Other As Long\n"
                "Public Function Mean() As Double\nEnd Function\n"
                "Public Sub Run()\nEnd Sub\n"
                "Public Property Get Size() As Long\nEnd Property\n"
                "Public Property Let Size(ByVal v As Long)\nEnd Property\n"
                "Public Enum Field\n    FIELD_A = 1\n    FIELD_B = 2\nEnd Enum\n"
                "Public Type Pair\n    a As Long\nEnd Type\n"
                "Public Declare PtrSafe Function Tick Lib \"kernel32\" () As Long\n"
                "Public Items() As String\n")
        self.assertEqual(public_members(module("CORE_A", body)),
                         {"limit", "total", "mean", "run", "size", "field", "field_a", "field_b", "pair", "tick",
                          "items"})

    def test_private_members_comments_and_strings_are_ignored(self) -> None:
        body = ("Private Function Hidden() As Long\nEnd Function\n"
                "' Public Function Commented() As Long\n"
                "Sub S(): x = \"Public Function InString()\": End Sub\n"
                "Dim local As Long\n")
        self.assertEqual(public_members(module("CORE_A", body)), set())


class DuplicatePublicNames(unittest.TestCase):
    def check(self, files: dict[str, str]) -> list[str]:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            for path, text in files.items():
                (root / path).parent.mkdir(parents=True, exist_ok=True)
                (root / path).write_text(text, encoding="cp1252")
            return check_public_names(root, sorted(files))

    def test_the_same_function_in_two_modules_is_reported(self) -> None:
        files = {"src/core/CORE_Parse.bas": module("CORE_Parse", "Public Function CORE_IsMonthEnd() As Boolean\n"
                                                                 "End Function\n"),
                 "src/core/CORE_Decay.bas": module("CORE_Decay", "Public Function core_ismonthend() As Boolean\n"
                                                                 "End Function\n")}
        self.assertEqual(self.check(files),
                         ["Public name core_ismonthend is declared in more than one module: "
                          "src/core/CORE_Decay.bas, src/core/CORE_Parse.bas"])

    def test_an_enum_member_clashing_with_a_constant_is_reported(self) -> None:
        files = {"src/modules/IRRBB_A.bas": module("IRRBB_A", "Public Enum F\n    IRRBB_X = 1\nEnd Enum\n"),
                 "src/modules/IRRBB_B.bas": module("IRRBB_B", "Public Const IRRBB_X As Long = 2\n")}
        self.assertEqual(len(self.check(files)), 1)

    def test_get_and_let_of_one_property_and_distinct_names_pass(self) -> None:
        files = {"src/core/CORE_A.bas": module("CORE_A", "Public Property Get Size() As Long\nEnd Property\n"
                                                         "Public Property Let Size(ByVal v As Long)\nEnd Property\n"),
                 "src/core/CORE_B.bas": module("CORE_B", "Public Function Other() As Long\nEnd Function\n"
                                                         "Private Function Size() As Long\nEnd Function\n")}
        self.assertEqual(self.check(files), [])


if __name__ == "__main__":
    unittest.main()
