"""Fixtures for the synchronizer and document-module checks of check_source.py.

Ported from the SACCR repository's tools/test_tooling.py: every VBA component
must lie in a $sourceRoots directory, which the synchronizer
(tools/Sync-IRRBB-VBA.ps1) scans against tools/vba-projects.json, and the
document modules in src/workbook/ must match the template's sheets by
CodeName.
"""
import shutil
import subprocess
import tempfile
import unittest
import zipfile
from pathlib import Path

import check_source

ROOT = Path(__file__).resolve().parents[1]
CHANGELOG = "# Changelog\n\n## [Unreleased]\n\n[Unreleased]: https://example.invalid\n"
MODULE = 'Attribute VB_Name = "IRRBB_Test"\r\nOption Explicit\r\n\r\nPublic Sub Run()\r\nEnd Sub\r\n'


class SyncGateTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.git("init", "-q")
        shutil.copy(ROOT / ".gitattributes", self.root / ".gitattributes")
        self.write("CHANGELOG.md", CHANGELOG)
        # The synchronizer and the template are required: by default the
        # synchronizer reads every documented VBA home and the template has no
        # CodeNames, so neither adds findings of its own.
        for home in check_source.VBA_HOMES:
            (self.root / home).mkdir(parents=True, exist_ok=True)
        self.write_sync(*(home.rstrip("/").replace("/", "\\") for home in check_source.VBA_HOMES))
        self.write_workbook("src/workbook/IRRBB_Template.xlsx", {
            "[Content_Types].xml": "<Types/>",
            "_rels/.rels": '<Relationships Target="xl/workbook.xml"/>',
            "xl/workbook.xml": "<workbook/>",
            "xl/_rels/workbook.xml.rels": "<Relationships/>"})

    def git(self, *args: str) -> str:
        return subprocess.run(["git", "-C", str(self.root), *args], check=True,
                              capture_output=True, text=True).stdout

    def write(self, path: str, text: str) -> None:
        target = self.root / path
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(text.encode("cp1252"))

    def write_workbook(self, path: str, parts: dict[str, str]) -> None:
        target = self.root / path
        target.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(target, "w") as package:
            for name, text in parts.items():
                package.writestr(name, text)

    def findings(self) -> list[str]:
        self.git("add", "-A")
        findings: list[str] = check_source.run_check(self.root)["findings"]
        return findings

    def test_fixture_is_clean(self) -> None:
        self.write("src/modules/IRRBB_Test.bas", MODULE)
        self.assertEqual(self.findings(), [])

    def write_sync(self, *directories: str) -> None:
        entries = ",\r\n".join(f'        (Join-Path $repoFullPath "{d}")' for d in directories)
        self.write("tools/Sync-IRRBB-VBA.ps1", f"$sourceRoots = @(\r\n{entries}\r\n    )\r\n")

    def test_component_outside_sync_roots_is_rejected(self) -> None:
        self.write_sync("src\\modules")
        self.write("src/modules/IRRBB_Test.bas", MODULE)
        self.write("tests/modules/TEST_Test.bas", MODULE.replace("IRRBB_Test", "TEST_Test"))
        self.assertEqual(self.findings(), ["tests/modules/TEST_Test.bas: not synchronized; add its "
                                           "directory to $sourceRoots in tools/Sync-IRRBB-VBA.ps1"])
        self.write_sync("src\\modules", "tests\\modules")
        self.assertEqual(self.findings(), [])

    def test_missing_sync_root_is_rejected(self) -> None:
        self.write_sync("src\\modules", "src\\nowhere")
        self.write("src/modules/IRRBB_Test.bas", MODULE)
        self.assertEqual(self.findings(),
                         ["tools/Sync-IRRBB-VBA.ps1: source directory src/nowhere/ does not exist"])

    def test_commented_sync_root_is_not_counted(self) -> None:
        self.write("tests/modules/TEST_Test.bas", MODULE.replace("IRRBB_Test", "TEST_Test"))
        unsynchronized = ("tests/modules/TEST_Test.bas: not synchronized; add its directory to "
                          "$sourceRoots in tools/Sync-IRRBB-VBA.ps1")
        for commented in ('        # (Join-Path $repoFullPath "tests\\modules")',
                          '        <# (Join-Path $repoFullPath "tests\\modules")\r\n        #>'):
            with self.subTest(entry=commented):
                self.write("tools/Sync-IRRBB-VBA.ps1",
                           '$sourceRoots = @(\r\n        (Join-Path $repoFullPath "src\\modules"),'
                           f"\r\n{commented}\r\n    )\r\n")
                self.assertEqual(self.findings(), [unsynchronized])
        self.write("tools/Sync-IRRBB-VBA.ps1",
                   '$sourceRoots = @(\r\n        (Join-Path $repoFullPath "tests\\modules")  # unit tests'
                   "\r\n    )\r\n")
        self.assertEqual(self.findings(), [])

    def test_commented_sync_declaration_is_not_the_active_list(self) -> None:
        # A complete list inside a block comment before an incomplete active
        # list must not hide the missing root.
        self.write("src/modules/IRRBB_Test.bas", MODULE)
        self.write("tests/modules/TEST_Test.bas", MODULE.replace("IRRBB_Test", "TEST_Test"))
        unsynchronized = ("tests/modules/TEST_Test.bas: not synchronized; add its directory to "
                          "$sourceRoots in tools/Sync-IRRBB-VBA.ps1")
        old = ('$sourceRoots = @(\r\n        (Join-Path $repoFullPath "src\\modules"),\r\n'
               '        (Join-Path $repoFullPath "tests\\modules")\r\n    )\r\n')
        active = '$sourceRoots = @(\r\n        (Join-Path $repoFullPath "src\\modules")\r\n    )\r\n'
        for before in (f"<#\r\n{old}#>\r\n", f"<# {old}#>\r\n",
                       "".join(f"# {line}\r\n" for line in old.split("\r\n") if line)):
            with self.subTest(before=before[:12]):
                self.write("tools/Sync-IRRBB-VBA.ps1", before + active)
                self.assertEqual(self.findings(), [unsynchronized])
        # A commented declaration and no active list: no list is found.
        self.write("tools/Sync-IRRBB-VBA.ps1", f"<#\r\n{old}#>\r\n")
        self.assertIn("tools/Sync-IRRBB-VBA.ps1: $sourceRoots list not found", self.findings())
        # A complete active list after a block comment is accepted.
        self.write("tools/Sync-IRRBB-VBA.ps1", "<# synchronizer\r\n   notes #>\r\n" + old)
        self.assertEqual(self.findings(), [])
        # A quoted "<#" before the list does not open a comment, also in a
        # string or here-string over several lines.
        for before in ('$label = "<# not a comment"\r\n',
                       '$label = "first line\r\n<# second line"\r\n',
                       "$help = @'\r\n<# text\r\n'@\r\n",
                       '$help = @"\r\n  <# text $x\r\n"@\r\n'):
            with self.subTest(string=before[:16]):
                self.write("tools/Sync-IRRBB-VBA.ps1", before + old)
                self.assertEqual(self.findings(), [])
        # A declaration inside a here-string is text, not the active list.
        self.write("tools/Sync-IRRBB-VBA.ps1", "$doc = @'\r\n" + old + "'@\r\n" + active)
        self.assertEqual(self.findings(), [unsynchronized])

    def test_backtick_escapes_in_sync_preamble(self) -> None:
        # In a double-quoted string a backtick escapes the next character, so
        # `" does not close the string and a following <# is text; `` is a
        # literal backtick and the quote after it does close the string. In a
        # single-quoted string a backtick is literal.
        self.write("src/modules/IRRBB_Test.bas", MODULE)
        self.write("tests/modules/TEST_Test.bas", MODULE.replace("IRRBB_Test", "TEST_Test"))
        unsynchronized = ("tests/modules/TEST_Test.bas: not synchronized; add its directory to "
                          "$sourceRoots in tools/Sync-IRRBB-VBA.ps1")
        old = ('$sourceRoots = @(\r\n        (Join-Path $repoFullPath "src\\modules"),\r\n'
               '        (Join-Path $repoFullPath "tests\\modules")\r\n    )\r\n')
        active = '$sourceRoots = @(\r\n        (Join-Path $repoFullPath "src\\modules")\r\n    )\r\n'
        for before in ('$x = "literal `"<#`""\r\n',
                       '$x = "`"# not a comment `"<# nor this"\r\n',
                       '$x = "first `"\r\n<# still text"\r\n'):
            with self.subTest(string=before[:24]):
                self.write("tools/Sync-IRRBB-VBA.ps1", before + old)
                self.assertEqual(self.findings(), [])
        # After an escaped backtick the quote closes the string, so the <#
        # opens a real comment and the complete list inside it is not active;
        # the same holds after a single-quoted string ending in a backtick.
        for before in (f'$x = "C:\\path``" <#\r\n{old}#>\r\n',
                       f"$y = 'C:\\path`' <#\r\n{old}#>\r\n"):
            with self.subTest(string=before[:24]):
                self.write("tools/Sync-IRRBB-VBA.ps1", before + active)
                self.assertEqual(self.findings(), [unsynchronized])

    def test_unsupported_sync_entry_is_rejected(self) -> None:
        self.write("tools/Sync-IRRBB-VBA.ps1",
                   '$sourceRoots = @(\r\n        (Join-Path $repoFullPath "src\\modules"),\r\n'
                   "        (Get-Item $env:SOURCE)\r\n")
        found = self.findings()
        self.assertIn("tools/Sync-IRRBB-VBA.ps1: unsupported $sourceRoots entry: (Get-Item $env:SOURCE)", found)
        self.assertIn("tools/Sync-IRRBB-VBA.ps1: $sourceRoots list is not closed", found)

    def test_required_files_must_exist(self) -> None:
        for path in ("tools/Sync-IRRBB-VBA.ps1", "src/workbook/IRRBB_Template.xlsx"):
            with self.subTest(path=path):
                self.findings()
                (self.root / path).unlink()
                self.assertIn(f"{path}: tracked but missing from the working tree",
                              check_source.run_check(self.root)["findings"])
                self.assertIn(f"{path}: required file is not tracked", self.findings())
                self.setUp()

    def test_document_modules_match_template_sheets(self) -> None:
        self.write_workbook("src/workbook/IRRBB_Template.xlsx", {
            "[Content_Types].xml": "<Types/>",
            "_rels/.rels": '<Relationships Target="xl/workbook.xml"/>',
            "xl/workbook.xml": '<workbook><workbookPr codeName="ThisWorkbook"/></workbook>',
            "xl/_rels/workbook.xml.rels": "<Relationships/>",
            "xl/worksheets/sheet1.xml": '<worksheet><sheetPr codeName="shInputs"/></worksheet>'})
        document = 'VERSION 1.0 CLASS\r\nAttribute VB_Name = "{0}"\r\nOption Explicit\r\n'
        self.write("src/workbook/ThisWorkbook.cls", document.format("ThisWorkbook"))
        self.write("src/workbook/shOutputs.cls", document.format("shOutputs"))
        found = self.findings()
        self.assertIn("src/workbook/shOutputs.cls: no sheet or workbook with this CodeName in "
                      "src/workbook/IRRBB_Template.xlsx", found)
        self.assertIn("src/workbook/IRRBB_Template.xlsx: CodeName shInputs has no document module "
                      "src/workbook/shInputs.cls", found)
        (self.root / "src/workbook/shOutputs.cls").unlink()
        self.write("src/workbook/shInputs.cls", document.format("shInputs"))
        self.assertEqual(self.findings(), [])


if __name__ == "__main__":
    unittest.main()
