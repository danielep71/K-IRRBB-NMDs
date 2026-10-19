"""Tests for tools/compare_vba_export.py, without Excel (issue #6).

Each test writes an "export" from the repository source at HEAD, with the
CRLF line endings the VBE uses, changes it the way a faulty round trip
would, and runs the comparison.
"""
from __future__ import annotations

import hashlib
import io
import json
import subprocess
import tempfile
import unittest
import zipfile
from contextlib import redirect_stderr, redirect_stdout
from pathlib import Path
from typing import Any, Callable

import compare_vba_export as tool

ROOT = Path(__file__).resolve().parent.parent
MANIFEST = json.loads((ROOT / "tools" / "vba-projects.json").read_text(encoding="utf-8"))
PRODUCTION = len(MANIFEST["configurations"]["production"]["modules"]) + len(MANIFEST["document_modules"])


def source(path: str) -> bytes:
    return subprocess.run(["git", "-C", str(ROOT), "show", f"HEAD:{path}"],
                          capture_output=True, check=True).stdout


class CompareExportTests(unittest.TestCase):
    def setUp(self) -> None:
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.folder = Path(temp.name)
        # The stand-in workbooks live beside the export, in a folder of this
        # test's own, so that parallel runs cannot remove each other's files.
        books = tempfile.TemporaryDirectory()
        self.addCleanup(books.cleanup)
        self.books = Path(books.name)

    def export(self, mode: str) -> None:
        for path in [*MANIFEST["configurations"][mode]["modules"], *MANIFEST["document_modules"]]:
            (self.folder / Path(path).name).write_bytes(source(path).replace(b"\n", b"\r\n"))

    def run_tool(self, mode: str, *extra: str) -> tuple[int, str]:
        out, err = io.StringIO(), io.StringIO()
        with redirect_stdout(out), redirect_stderr(err):
            code = tool.main(["--root", str(ROOT), "--export", str(self.folder), "--mode", mode, *extra])
        return code, out.getvalue() + err.getvalue()

    def workbook(self, name: str, project: bytes = b"vba project") -> Path:
        """A stand-in for a built .xlsm: a zip holding a VBA project part."""
        path = self.books / name
        with zipfile.ZipFile(path, "w") as archive:
            archive.writestr("[Content_Types].xml", "<Types/>")
            archive.writestr("xl/vbaProject.bin", project)
        return path

    def write_manifest(self, workbook: Path) -> None:
        """export-manifest.json as Export-IRRBB-VBA.ps1 writes it."""
        components = []
        for path in sorted(self.folder.iterdir()):
            if path.suffix in (".bas", ".cls", ".frm", ".frx"):
                name = path.stem if path.suffix == ".frx" else tool.parse(path.read_bytes(), path.name).name
                components.append({"name": name, "file": path.name,
                                   "sha256": hashlib.sha256(path.read_bytes()).hexdigest().upper()})
        record = {"format": "irrbb-vba-export-1",
                  "workbook": {"file": workbook.name, "sha256": hashlib.sha256(workbook.read_bytes()).hexdigest()},
                  "components": components}
        (self.folder / "export-manifest.json").write_text(json.dumps(record, indent=4), encoding="utf-8")

    def edit_manifest(self, change: Callable[[dict[str, Any]], object]) -> None:
        path = self.folder / "export-manifest.json"
        record = json.loads(path.read_text(encoding="utf-8"))
        change(record)
        path.write_text(json.dumps(record), encoding="utf-8")

    def edit(self, name: str, old: bytes, new: bytes) -> None:
        path = self.folder / name
        data = path.read_bytes()
        self.assertIn(old, data)
        path.write_bytes(data.replace(old, new, 1))

    def test_exact_export_passes_in_both_modes(self) -> None:
        for mode in ("production", "development"):
            with self.subTest(mode=mode):
                for path in self.folder.iterdir():
                    path.unlink()
                self.export(mode)
                code, out = self.run_tool(mode)
                self.assertEqual(code, 0, out)
                self.assertIn("RESULT=PASS", out)

    def test_trailing_blank_lines_and_line_endings_are_ignored(self) -> None:
        self.export("production")
        path = self.folder / "CORE_Codes.bas"
        path.write_bytes(path.read_bytes().replace(b"\r\n", b"\n") + b"\r\n\r\n")
        self.assertEqual(self.run_tool("production")[0], 0)

    def test_changed_code_fails_with_its_line(self) -> None:
        self.export("production")
        self.edit("WB_Import.bas", b'"WB_Import.WB_ImportData"', b'"WB_Import.WB_Import"')
        code, out = self.run_tool("production")
        self.assertEqual(code, 1)
        self.assertIn("src/workbook/WB_Import.bas: code differs at line", out)
        self.assertIn("'WB_Import.WB_Import'", out.replace('"', "'"))

    def test_identifier_case_is_ignored_and_counted(self) -> None:
        # The VBE gives each name one spelling across the project.
        self.export("production")
        self.edit("CORE_Import.bas", b"Private Const BALANCE_DECIMALS As Long = 4",
                  b"Private Const Balance_Decimals As Long = 4")
        self.edit("CORE_Codes.bas", b"Option Explicit", b"option explicit")
        code, out = self.run_tool("production")
        self.assertEqual(code, 0, out)
        self.assertIn("Re-cased:  2 line(s)", out)

    def test_case_change_in_a_string_or_comment_fails(self) -> None:
        for path, old, new in (("src/workbook/WB_Import.bas", b'"WB_Import.WB_ImportData"',
                                b'"wb_import.wb_importdata"'),
                               ("src/core/CORE_Parse.bas", b"' PURPOSE", b"' purpose")):
            with self.subTest(path=path):
                for item in self.folder.iterdir():
                    item.unlink()
                self.export("production")
                self.edit(Path(path).name, old, new)
                code, out = self.run_tool("production")
                self.assertEqual(code, 1, out)
                self.assertIn(f"{path}: code differs at line", out)

    def test_fold_case_keeps_strings_and_comments(self) -> None:
        self.assertEqual(tool.fold_case("    MsgBox \"It's OK\", , MSG_TITLE 'Show It"),
                         "    msgbox \"It's OK\", , msg_title 'Show It")
        self.assertEqual(tool.fold_case("    Rem Keep THIS"), "    rem Keep THIS")
        self.assertEqual(tool.fold_case("    X = 1: Rem Keep THIS"), "    x = 1: rem Keep THIS")
        self.assertEqual(tool.fold_case("Done: REM it's KEPT"), "done: rem it's KEPT")
        self.assertEqual(tool.fold_case("10 Rem Keep THIS"), "10 rem Keep THIS")
        self.assertEqual(tool.fold_case("    Remaining = Rem2"), "    remaining = rem2")
        self.assertEqual(tool.fold_case('    S = """A"" B"'), '    s = """A"" B"')

    def test_changed_attribute_of_a_standard_module_fails(self) -> None:
        self.export("production")
        self.edit("CORE_Codes.bas", b'Attribute VB_Name = "CORE_Codes"\r\n',
                  b'Attribute VB_Name = "CORE_Codes"\r\nAttribute VB_Description = "x"\r\n')
        code, out = self.run_tool("production")
        self.assertEqual(code, 1)
        self.assertIn("src/core/CORE_Codes.bas: attributes differ", out)

    def test_document_module_attributes_are_not_compared(self) -> None:
        # Excel writes them and the synchronizer replaces only the code. An
        # export from Excel can lack VB_Base, VB_TemplateDerived and
        # VB_Customizable on the sheet modules.
        self.export("production")
        for line in (b'Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"\r\n',
                     b"Attribute VB_TemplateDerived = False\r\n", b"Attribute VB_Customizable = True\r\n"):
            self.edit("shConfig.cls", line, b"")
        code, out = self.run_tool("production")
        self.assertEqual(code, 0, out)

    def test_missing_and_extra_components_fail(self) -> None:
        self.export("development")
        (self.folder / "IRRBB_Import.bas").unlink()
        code, out = self.run_tool("production")
        self.assertEqual(code, 1)
        self.assertIn("missing from the export: IRRBB_Import", out)
        self.assertIn("not in the source for this mode: TEST_Harness", out)

    def test_component_without_vb_name_is_an_error(self) -> None:
        self.export("production")
        (self.folder / "Stray.bas").write_bytes(b"Option Explicit\r\n")
        code, out = self.run_tool("production")
        self.assertEqual(code, 2)
        self.assertIn("Stray.bas: expected one Attribute VB_Name line, found 0", out)
    def test_bound_export_is_verified(self) -> None:
        workbook = self.workbook("A.xlsm")
        self.export("production")
        self.write_manifest(workbook)
        code, out = self.run_tool("production", "--workbook", str(workbook))
        self.assertEqual(code, 0, out)
        self.assertIn("Binding:   VERIFIED", out)
        self.assertIn(f"RESULT=PASS; components={PRODUCTION}; differences=0; binding=VERIFIED", out)

    def test_renamed_workbook_is_verified_with_a_note(self) -> None:
        # A download is often renamed ("IRRBB_0.2.0 (1).xlsm"): the bytes
        # bind it, the name only gets a note.
        self.export("production")
        self.write_manifest(self.workbook("IRRBB_0.2.0.xlsm"))
        renamed = self.workbook("IRRBB_0.2.0 (1).xlsm")
        code, out = self.run_tool("production", "--workbook", str(renamed))
        self.assertEqual(code, 0, out)
        self.assertIn("NOTE export-manifest.json names the workbook IRRBB_0.2.0.xlsm", out)
        self.assertIn("binding=VERIFIED", out)

    def test_export_compared_with_another_workbook_fails(self) -> None:
        # The review's case: the export of workbook A presented with workbook B.
        self.export("production")
        self.write_manifest(self.workbook("A.xlsm", b"project A"))
        code, out = self.run_tool("production", "--workbook", str(self.workbook("B.xlsm", b"project B")))
        self.assertEqual(code, 1, out)
        self.assertIn("B.xlsm is not the workbook the export was made from", out)
        self.assertNotIn("NOTE", out)  # another workbook, not a renamed one
        self.assertIn("differences=0; binding=FAIL", out)

    def test_non_workbook_fails_even_when_the_manifest_names_it(self) -> None:
        readme = ROOT / "README.md"
        self.export("production")
        self.write_manifest(readme)
        code, out = self.run_tool("production", "--workbook", str(readme))
        self.assertEqual(code, 1, out)
        self.assertIn("README.md is not a macro-enabled workbook", out)

    def test_component_changed_after_the_export_fails_the_binding(self) -> None:
        # A file replaced by one that still matches the source: parity holds,
        # the binding does not.
        workbook = self.workbook("A.xlsm")
        self.export("production")
        self.write_manifest(workbook)
        path = self.folder / "CORE_Codes.bas"
        path.write_bytes(path.read_bytes().replace(b"\r\n", b"\n"))
        code, out = self.run_tool("production", "--workbook", str(workbook))
        self.assertEqual(code, 1, out)
        self.assertIn("CORE_Codes.bas changed after the export", out)
        self.assertIn("differences=0; binding=FAIL", out)

    def test_form_resources_are_bound(self) -> None:
        # The VBE writes a form's binary properties to a .frx beside its .frm:
        # a changed or unlisted .frx fails the binding like any component.
        workbook = self.workbook("A.xlsm")
        self.export("production")
        (self.folder / "UF_Picker.frx").write_bytes(b"form resources")
        self.write_manifest(workbook)
        self.assertEqual(self.run_tool("production", "--workbook", str(workbook))[0], 0)
        (self.folder / "UF_Picker.frx").write_bytes(b"other resources")
        code, out = self.run_tool("production", "--workbook", str(workbook))
        self.assertEqual(code, 1, out)
        self.assertIn("UF_Picker.frx changed after the export", out)
        self.write_manifest(workbook)
        (self.folder / "UF_Other.frx").write_bytes(b"form resources")
        code, out = self.run_tool("production", "--workbook", str(workbook))
        self.assertEqual(code, 1, out)
        self.assertIn("UF_Other.frx is in the export but not in export-manifest.json", out)

    def test_modified_manifest_fails(self) -> None:
        workbook = self.workbook("A.xlsm")
        cases = {
            "dropped entry": (lambda r: r["components"].pop(0), "is in the export but not in export-manifest.json"),
            "extra entry": (lambda r: r["components"].append(
                {"name": "TEST_Harness", "file": "TEST_Harness.bas", "sha256": "0" * 64}),
                "lists TEST_Harness.bas, which is not in the export"),
            "renamed component": (lambda r: r["components"][0].update(name="Other"), "names Other"),
            "workbook digest": (lambda r: r["workbook"].update(sha256="0" * 64),
                                "is not the workbook the export was made from"),
            "format": (lambda r: r.update(format="other"), "is not a valid export manifest"),
            "missing key": (lambda r: r.pop("workbook"), "is not a valid export manifest"),
            "missing workbook name": (lambda r: r["workbook"].pop("file"), "is not a valid export manifest"),
            "empty workbook name": (lambda r: r["workbook"].update(file=""), "workbook.file is not a file name"),
        }
        for label, (change, message) in cases.items():
            with self.subTest(label):
                for item in self.folder.iterdir():
                    item.unlink()
                self.export("production")
                self.write_manifest(workbook)
                self.edit_manifest(change)
                code, out = self.run_tool("production", "--workbook", str(workbook))
                self.assertEqual(code, 1, out)
                self.assertIn(message, out)
                self.assertIn("binding=FAIL", out)

    def test_export_without_manifest_is_unverified(self) -> None:
        # An export made before the manifest: source parity only.
        self.export("production")
        code, out = self.run_tool("production", "--workbook", str(self.workbook("A.xlsm")))
        self.assertEqual(code, 0, out)
        self.assertIn("no export-manifest.json: the export is not bound to a workbook", out)
        self.assertIn(f"RESULT=PASS; components={PRODUCTION}; differences=0; binding=UNVERIFIED", out)

    def test_manifest_without_workbook_is_unverified_but_still_checked(self) -> None:
        self.export("production")
        self.write_manifest(self.workbook("A.xlsm"))
        code, out = self.run_tool("production")
        self.assertEqual(code, 0, out)
        self.assertIn("binding=UNVERIFIED", out)
        (self.folder / "WB_Import.bas").unlink()
        code, out = self.run_tool("production")
        self.assertEqual(code, 1, out)
        self.assertIn("lists WB_Import.bas, which is not in the export", out)

    def test_missing_workbook_is_an_error(self) -> None:
        self.export("production")
        code, out = self.run_tool("production", "--workbook", str(self.folder / "absent.xlsm"))
        self.assertEqual(code, 2, out)
        self.assertIn("workbook not found", out)

    def test_userform_export_header_is_parsed(self) -> None:
        # A form export: VERSION 5.00, a title-case Begin ... End block with a
        # nested block, then the attributes and the code.
        raw = ("VERSION 5.00\r\n"
               "Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_Picker \r\n"
               "   Caption         =   \"Pick\"\r\n"
               "   Begin {0BA686C6-F7D3-101A-993E-0000C0EF6F5E} Ctl \r\n"
               "   End\r\n"
               "End\r\n"
               'Attribute VB_Name = "UF_Picker"\r\n'
               "Attribute VB_GlobalNameSpace = False\r\n"
               "Option Explicit\r\n").encode("cp1252")
        component = tool.parse(raw, "UF_Picker.frm")
        self.assertEqual(component.name, "UF_Picker")
        self.assertEqual(component.attributes,
                         ('Attribute VB_Name = "UF_Picker"', "Attribute VB_GlobalNameSpace = False"))
        self.assertEqual(component.body, ("Option Explicit",))


if __name__ == "__main__":
    unittest.main()
