"""Regression fixtures for the src/core host-independence rule."""
from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from check_source import check_component, check_core_host_independence

HEADER = 'Attribute VB_Name = "CORE_Sample"\nOption Explicit\nOption Private Module\n'


def findings(body: str) -> list[str]:
    return check_core_host_independence("src/core/CORE_Sample.bas", HEADER + body)


class CoreHostIndependence(unittest.TestCase):
    def test_pure_numeric_code_passes(self) -> None:
        body = ("Public Function Mean(ByRef values() As Double) As Double\n"
                "    Dim i As Long, total As Double\n"
                "    For i = LBound(values) To UBound(values): total = total + values(i): Next i\n"
                "    Mean = total / (UBound(values) - LBound(values) + 1)\n"
                "End Function\n")
        self.assertEqual(findings(body), [])

    def test_com_automation_is_rejected_regardless_of_progid(self) -> None:
        for expression, name in (
                ('CreateObject("Excel.Application")', "CreateObject"),
                ('GetObject(, "Excel.Application")', "GetObject"),
                ('VBA.Interaction.cReAtEoBjEcT("excel.application")', "cReAtEoBjEcT"),
                ('VBA.GetObject(Class:="Excel.Application")', "GetObject"),
                ('CreateObject("Excel." & "Application")', "CreateObject"),
                ('CreateObject(progId)', "CreateObject"),
                ('GetObject(workbookPath)', "GetObject"),
                ('CreateObject("Scripting.Dictionary")', "CreateObject"),
                ('CreateObject( _\n        progId)', "CreateObject")):
            with self.subTest(expression=expression):
                self.assertEqual(findings(f"Sub S()\n    Set obj = {expression}\nEnd Sub\n"),
                                 [f"src/core/CORE_Sample.bas:5: core must not use COM automation identifier {name}"])

    def test_automation_names_in_comments_strings_and_longer_names_pass(self) -> None:
        body = ("' CreateObject _\nGetObject\n"
                "Rem GetObject\n"
                'Sub S(): s = "CreateObject(""Excel.Application"")"\n'
                '    Dim CreateObjectText As String, GetObjectCount As Long\n'
                "End Sub\n")
        self.assertEqual(findings(body), [])

    def test_comment_continuations_accept_vba_whitespace(self) -> None:
        whitespace = [" ", "\t", "\x19", "\u1680", "\u202f", "\u205f", "\u3000"]
        whitespace.extend(chr(n) for n in range(0x2000, 0x200b))
        for prefix in ("'", "Rem"):
            for space in whitespace:
                with self.subTest(prefix=prefix, space=repr(space)):
                    body = f"{prefix} prose{space}_\nGetObject{space}_\nCreateObject\nRange\n"
                    self.assertEqual(findings(body),
                                     ["src/core/CORE_Sample.bas:7: core must not use Excel host identifier Range"])

    def test_non_continuations_do_not_hide_following_code(self) -> None:
        for suffix in ("_", "\u00a0_", "\v_", "\f_", " _ ", " _\t"):
            with self.subTest(suffix=repr(suffix)):
                self.assertEqual(findings(f"' prose{suffix}\nGetObject\n"),
                                 ["src/core/CORE_Sample.bas:5: core must not use COM automation identifier GetObject"])

    def test_host_adapter_may_use_com_automation(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            path = "src/workbook/WB_Adapter.bas"
            (root / path).parent.mkdir(parents=True)
            (root / path).write_text('Attribute VB_Name = "WB_Adapter"\nOption Explicit\n'
                                     'Sub S()\nSet obj = CreateObject("Excel.Application")\nEnd Sub\n',
                                     encoding="cp1252")
            self.assertEqual(check_component(root, path, {path}, {}), [])

    def test_object_model_use_is_reported_with_line(self) -> None:
        result = findings("Public Sub Load()\n    x = Range(\"A1\").Value\nEnd Sub\n")
        self.assertEqual(result, ["src/core/CORE_Sample.bas:5: core must not use Excel host identifier Range"])

    def test_each_reserved_identifier_is_reported(self) -> None:
        for name in ("Application", "ThisWorkbook", "ActiveSheet", "Worksheets", "Cells",
                     "WorksheetFunction", "MsgBox", "InputBox", "Excel"):
            with self.subTest(name=name):
                self.assertEqual(len(findings(f"Sub S()\n    Call {name}\nEnd Sub\n")), 1)

    def test_identifiers_are_case_insensitive(self) -> None:
        self.assertEqual(len(findings("Sub S()\n    Dim range As Double\nEnd Sub\n")), 1)

    def test_longer_identifiers_are_not_reserved(self) -> None:
        body = "Sub S()\n    Dim rangeMin As Double, cellsUsed As Long, MsgBoxText As String\nEnd Sub\n"
        self.assertEqual(findings(body), [])

    def test_comments_and_strings_are_ignored(self) -> None:
        body = ("' Range and MsgBox may be named in a comment _\n"
                "  and its continuation line: Application\n"
                "Rem ActiveSheet in a Rem statement\n"
                "Sub S(): Rem Worksheets after a colon\n"
                "    s = \"Range \"\"Cells\"\" MsgBox\" ' trailing Application\n"
                "End Sub\n")
        self.assertEqual(findings(body), [])

    def test_code_after_a_string_with_quotes_is_still_checked(self) -> None:
        self.assertEqual(len(findings("Sub S()\n    s = \"a\"\"b\": Set r = Cells(1, 1)\nEnd Sub\n")), 1)

    def test_rule_applies_only_to_src_core(self) -> None:
        body = 'Option Explicit\nSub S()\n    MsgBox "hi"\nEnd Sub\n'
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            for path, name in (("src/core/CORE_Ui.bas", "CORE_Ui"), ("src/modules/IRRBB_Ui.bas", "IRRBB_Ui")):
                (root / path).parent.mkdir(parents=True, exist_ok=True)
                private = "Option Private Module\n" if path.startswith("src/core/") else ""
                (root / path).write_text(f'Attribute VB_Name = "{name}"\n{private}{body}', encoding="cp1252")
            core = check_component(root, "src/core/CORE_Ui.bas", set(), {})
            facade = check_component(root, "src/modules/IRRBB_Ui.bas", set(), {})
        self.assertEqual(core, ["src/core/CORE_Ui.bas:5: core must not use Excel host identifier MsgBox"])
        self.assertEqual(facade, [])


if __name__ == "__main__":
    unittest.main()
