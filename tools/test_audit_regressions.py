"""Negative cases reproduced in the 2026-10-08 repository audit."""
from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import check_source
import test_review_regressions as packages
from check_vba_conditionals import analyze_component as conditionals
from check_vba_jumps import analyze_component as jumps
from check_vba_public_api import parse_component
from test_review_regressions import CORE_PARTS


class LexerRegressions(unittest.TestCase):
    def test_commented_label_does_not_resolve_jump(self):
        for prefix in ("' example", "Rem example", "x = 1: Rem example"):
            for space in (" ", "\t", "\x19", "\u3000"):
                with self.subTest(prefix=prefix, space=space):
                    findings = jumps("Fixture.bas", f"Public Sub Test()\nGoTo Missing\n{prefix}{space}_\nMissing:\nEnd Sub\n")
                    self.assertTrue(any(f.get("target") == "Missing" for f in findings), findings)

    def test_commented_declaration_is_not_public_api(self):
        entries, findings = parse_component("src/modules/IRRBB_Test.bas",
                                             "' example _\nPublic Const Fake As Long = 1\n", True)
        self.assertEqual(entries, [])
        self.assertEqual(findings, [])

    def test_commented_conditional_directive_and_declare_are_ignored(self):
        for hidden in ("#If UNKNOWN Then", 'Public Declare Function F Lib "x" () As Long'):
            self.assertEqual(conditionals("Fixture.bas", "Rem note _\n" + hidden + "\n"), [])

    def test_real_code_after_comment_chain_retains_line_number(self):
        findings = conditionals("Fixture.bas", "' note _\nstill a comment _\nmore comment\n"
                               'Public Declare Function F Lib "x" () As Long\n')
        self.assertEqual([f["line"] for f in findings], [4])

    def test_directive_code_continuation_is_supported(self):
        self.assertEqual(conditionals("Fixture.bas", "#If VBA7 And _\nWin64 Then\n"
                         'Public Declare PtrSafe Function F Lib "x" () As Long\n#End If\n'), [])

    def test_string_literals_and_non_continuations(self):
        self.assertEqual(jumps("Fixture.bas", 'Public Sub Test()\ns = "Rem _": GoTo Valid\nValid:\nEnd Sub'), [])
        for suffix in (" _ ", "\u00a0_", "\v_", "\f_"):
            self.assertTrue(conditionals("Fixture.bas", "' note" + suffix + "\n"
                            'Public Declare Function F Lib "x" () As Long\n'))


class PackageRegressions(unittest.TestCase):
    def test_utf16_and_character_references_cannot_hide_forbidden_targets(self):
        helper = packages.WorkbookPackageTests()
        for part, attr, value in (("_rels/.rels", "Target", "docProps/core.xml"),
                                  ("xl/_rels/workbook.xml.rels", "Target", "vbaProject.bin"),
                                  ("[Content_Types].xml", "PartName", "/docProps/core.xml")):
            for encoding in ("utf-8", "utf-16"):
                xml = f'<?xml version="1.0" encoding="{encoding}"?><Root><Item {attr}="{value}"/></Root>'
                self.assertTrue(helper.check_parts({**CORE_PARTS, part: xml.encode(encoding)}))
        self.assertTrue(helper.check_parts({**CORE_PARTS, "_rels/.rels":
                       '<Root><Item Target="docPr&#111;ps/core.xml"/></Root>'}))

    def test_tracked_workbook_outside_exact_template_path_fails(self):
        with tempfile.TemporaryDirectory() as directory:
            for path, allowed in (("Unexpected.xlsx", False), ("src/workbook/Other.xlsx", False),
                                  ("src/workbook/IRRBB_Template.xlsm", False),
                                  ("src/workbook/IRRBB_Template.xlsx", True)):
                with self.subTest(path=path), patch.object(check_source, "tracked_files", return_value={path}), \
                        patch.object(check_source, "check_storage", return_value=[]), \
                        patch.object(check_source, "check_workbook", return_value=[]), \
                        patch.object(check_source, "check_changelog", return_value=[]), \
                        patch.object(check_source, "check_sync_script", return_value=[]), \
                        patch.object(check_source, "check_document_modules", return_value=[]):
                    report = check_source.run_check(Path(directory))
                    self.assertEqual(report["status"], "pass" if allowed else "fail")




if __name__ == "__main__":
    unittest.main()
