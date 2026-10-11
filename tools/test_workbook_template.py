"""Check the workbook template against its generator, its contract and the harness, without Excel.

src/workbook/IRRBB_Template.xlsx must be exactly what tools/workbook/build_template.py
produces, the generator must agree with the tables in docs/WORKBOOK_CONTRACT.md,
the package must be sanitized and formula-free, and the Checks names must cover
the areas TEST_Harness writes. Standard library only.
"""
from __future__ import annotations

import io
import re
import sys
import unittest
import zipfile
from datetime import date
from pathlib import Path
from xml.etree import ElementTree

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools" / "workbook"))
sys.path.insert(0, str(ROOT / "tools"))
import build_template  # noqa: E402
import test_data_contract as reference  # noqa: E402

CONTRACT = ROOT / "docs" / "WORKBOOK_CONTRACT.md"
HARNESS = ROOT / "tests" / "modules" / "TEST_Harness.bas"
NS = {"m": build_template.MAIN_NS}


def table(heading: str) -> list[list[str]]:
    """Body rows of the first Markdown table after a heading, cells stripped of backticks."""
    text = CONTRACT.read_text(encoding="utf-8").split(heading, 1)[1]
    rows = []
    for line in text.splitlines():
        if line.startswith("|"):
            rows.append([cell.strip().strip("`") for cell in line.strip("|").split("|")])
        elif rows:
            break
    return rows[2:]


class Template(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.blob = build_template.TEMPLATE.read_bytes()
        cls.package = zipfile.ZipFile(io.BytesIO(cls.blob))

    def xml(self, name: str) -> ElementTree.Element:
        return ElementTree.fromstring(self.package.read(name))

    def sheet(self, name: str) -> ElementTree.Element:
        index = [s.name for s in build_template.SHEETS].index(name) + 1
        return self.xml(f"xl/worksheets/sheet{index}.xml")

    def strings(self) -> list[str]:
        return [si.findtext("m:t", namespaces=NS) for si in self.xml("xl/sharedStrings.xml")]

    def value(self, sheet: str, ref: str) -> str | None:
        cell = self.sheet(sheet).find(f".//m:c[@r='{ref}']", NS)
        if cell is None or cell.find("m:v", NS) is None:
            return None
        return self.strings()[int(cell.findtext("m:v", namespaces=NS))]

    def test_committed_template_is_the_generator_output(self) -> None:
        self.assertEqual(build_template.members(self.blob), build_template.members(build_template.package()),
                         "rebuild with: python tools/workbook/build_template.py")

    def test_package_is_macro_free_and_sanitized(self) -> None:
        names = self.package.namelist()
        for forbidden in ("docProps/", "vbaProject", "externalLink", "customXml", "printerSettings", "comments"):
            self.assertFalse([n for n in names if forbidden.lower() in n.lower()], forbidden)
        types = self.package.read("[Content_Types].xml").decode()
        self.assertIn("spreadsheetml.sheet.main+xml", types)
        self.assertNotIn("macroEnabled", types)
        for info in self.package.infolist():
            self.assertEqual(info.date_time, build_template.ZIP_TIME, info.filename)

    def test_shared_string_counts(self) -> None:
        table = self.xml("xl/sharedStrings.xml")
        references = sum(len(self.sheet(s.name).findall(".//m:c[@t='s']", NS)) for s in build_template.SHEETS)
        self.assertEqual(int(table.get("count")), references)
        self.assertEqual(int(table.get("uniqueCount")), len(table))
        self.assertLess(len(table), references)

    def test_no_cell_formulas(self) -> None:
        for s in build_template.SHEETS:
            with self.subTest(sheet=s.name):
                self.assertIsNone(self.sheet(s.name).find(".//m:f", NS))

    def test_sheets_match_the_contract(self) -> None:
        documented = [row for row in table("## 📑 Sheets")]
        generated = [[s.name, s.code_name, s.role, s.owner, s.status] for s in build_template.SHEETS]
        self.assertEqual(documented, generated)
        workbook = self.xml("xl/workbook.xml")
        self.assertEqual([e.get("name") for e in workbook.find("m:sheets", NS)], [s.name for s in build_template.SHEETS])
        self.assertEqual(workbook.find("m:workbookPr", NS).get("codeName"), "ThisWorkbook")
        for s in build_template.SHEETS:
            self.assertEqual(self.sheet(s.name).find("m:sheetPr", NS).get("codeName"), s.code_name)

    def test_readme_maps_every_sheet(self) -> None:
        for row, s in enumerate(build_template.SHEETS, 13):
            self.assertEqual([self.value("Readme", f"{c}{row}") for c in "ABCDE"],
                             [s.name, s.code_name, s.role, s.owner, s.status])

    def test_names_match_the_contract(self) -> None:
        documented = {}
        for name, refers, _, _ in table("## 🏷️ Named areas"):
            if name == "<Sheet>_Status":
                for s in build_template.SHEETS:
                    if s.role == "Output":
                        documented[f"{s.name}_Status"] = (s.name, "$B$3")
            else:
                sheet, ref = refers.split("!")
                documented[name] = (sheet, ref)
        self.assertEqual(documented, build_template.NAMES)
        defined = {e.get("name"): e.text for e in self.xml("xl/workbook.xml").find("m:definedNames", NS)}
        self.assertEqual(defined, {n: f"{s}!{r}" for n, (s, r) in build_template.NAMES.items()})

    def test_every_output_sheet_starts_not_current(self) -> None:
        for s in build_template.SHEETS:
            if s.role != "Output":
                continue
            with self.subTest(sheet=s.name):
                self.assertEqual(build_template.NAMES[f"{s.name}_Status"], (s.name, "$B$3"))
                status = self.value(s.name, "B3")
                self.assertTrue(status.startswith("NOT RUN"), status)
                rules = self.sheet(s.name).findall("m:conditionalFormatting[@sqref='B3']/m:cfRule", NS)
                self.assertEqual(len(rules), 3)
            if s.status.startswith("Placeholder"):
                self.assertTrue(self.value(s.name, "A5").startswith("Reserved: no function writes this sheet"))

    def test_validation_lists_follow_the_data_contract(self) -> None:
        self.assertEqual({code for code, _ in build_template.SEGMENTS}, reference.SEGMENTS)
        self.assertEqual([self.value("Segments", f"A{row}") for row in (6, 7, 8)],
                         [code for code, _ in build_template.SEGMENTS])
        epoch = date(1899, 12, 30)  # serial 2958465 = 9999-12-31; serial 1 = 1900-01-01 (1900 leap-year rule)
        self.assertEqual(build_template.LAST_SERIAL, (date(9999, 12, 31) - epoch).days)
        self.assertEqual(build_template.FIRST_SERIAL, 1)
        config = self.sheet("Config").findall(".//m:dataValidation", NS)
        self.assertEqual({(v.get("sqref"), v.get("type")) for v in config}, {("B8:B9", "date"), ("B11", "list")})
        dates = next(v for v in config if v.get("type") == "date")
        self.assertEqual((dates.findtext("m:formula1", namespaces=NS), dates.findtext("m:formula2", namespaces=NS)),
                         ("1", "2958465"))

    def test_checks_names_cover_the_harness_areas(self) -> None:
        source = HARNESS.read_text(encoding="cp1252")
        case_row = int(re.search(r"CASE_ROW As Long = (\d+)", source).group(1))
        assert_col = int(re.search(r"ASSERT_COL As Long = (\d+)", source).group(1))
        case_rows = int(re.search(r"\.Cells\(CASE_ROW \+ (\d+), 6\)\)\.ClearContents", source).group(1))
        assert_rows = int(re.search(r"\.Cells\(CASE_ROW \+ (\d+), ASSERT_COL \+ 6\)\)\.ClearContents",
                                    source).group(1))
        column = "ABCDEFGHIJKLMN"
        self.assertEqual(build_template.NAMES["Checks_Cases"],
                         ("Checks", f"$A${case_row}:$F${case_row + case_rows}"))
        self.assertEqual(build_template.NAMES["Checks_Assertions"],
                         ("Checks", f"${column[assert_col - 1]}${case_row}:"
                                    f"${column[assert_col + 5]}${case_row + assert_rows}"))
        self.assertEqual([self.value("Checks", f"A{row}") for row in range(3, 12)],
                         ["Status", "Started", "Finished", "Source commit", "Host", "Cases", "Assertions",
                          "Primary error", "Cleanup error"])

    def test_development_stamp_does_not_leak_into_the_template(self) -> None:
        build_template.package(["sha", "time", "builder", "components", "references"], ["note"])
        self.assertEqual(build_template.members(build_template.package()), build_template.members(self.blob))
        self.assertEqual(self.value("Readme", "A5"), "Source commit")
        self.assertEqual(self.value("Readme", "B5"), build_template.TEMPLATE_STAMP[0])


if __name__ == "__main__":
    unittest.main()
