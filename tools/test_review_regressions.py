"""Regressions for review findings on PRs #1 and #15; no Excel execution."""
from __future__ import annotations

import csv
import json
import lzma
import os
import shutil
import subprocess
import tempfile
import textwrap
import unittest
import warnings
import zipfile
from pathlib import Path

from check_source import check_workbook

ROOT = Path(__file__).resolve().parents[1]
CORE_PARTS = {
    "[Content_Types].xml": '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>',
    "_rels/.rels": '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>',
    "xl/workbook.xml": '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>',
    "xl/_rels/workbook.xml.rels": '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>',
}


class WorkbookPackageTests(unittest.TestCase):
    def check_parts(self, parts):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            with zipfile.ZipFile(root / "template.xlsx", "w") as archive:
                for name, value in parts.items():
                    archive.writestr(name, value)
            return check_workbook(root, "template.xlsx")

    def test_empty_archive_rejected(self):
        self.assertTrue(self.check_parts({}))

    def test_each_required_part_is_required(self):
        for missing in CORE_PARTS:
            with self.subTest(missing=missing):
                findings = self.check_parts({k: v for k, v in CORE_PARTS.items() if k != missing})
                self.assertIn(missing, " ".join(findings))

    def test_structural_fixture_accepted_without_claiming_excel_validity(self):
        # Deliberately a structural fixture, not a distributable Excel workbook.
        self.assertEqual(self.check_parts(CORE_PARTS), [])

    def test_malformed_core_xml_rejected(self):
        self.assertTrue(self.check_parts({**CORE_PARTS, "xl/workbook.xml": "<broken"}))

    def test_unsupported_xml_encodings_reported(self):
        for xml in ('<?xml version="1.0" encoding="NO-SUCH-ENCODING"?><workbook/>',
                    '<?xml version="1.0" encoding="UTF-32"?><workbook/>'):
            with self.subTest(xml=xml):
                self.assertTrue(self.check_parts({**CORE_PARTS, "xl/workbook.xml": xml}))

    def test_binary_workbook_uses_binary_parts(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            path = root / "template.xlsb"
            with zipfile.ZipFile(path, "w") as archive:
                archive.writestr("[Content_Types].xml", CORE_PARTS["[Content_Types].xml"])
                archive.writestr("_rels/.rels", CORE_PARTS["_rels/.rels"])
                # Synthetic binary bytes exercise structural dispatch, not BIFF validity.
                archive.writestr("xl/workbook.bin", b"\x83\x01\x00\x84\x01\x00")
                archive.writestr("xl/_rels/workbook.bin.rels", CORE_PARTS["xl/_rels/workbook.xml.rels"])
            self.assertEqual(check_workbook(root, path.name), [])
            self.assertTrue(check_workbook(root, "missing.xlsb"))

    def test_forbidden_parts_still_rejected(self):
        for name in ("xl/vbaProject.bin", "docProps/core.xml"):
            with self.subTest(name=name):
                self.assertTrue(self.check_parts({**CORE_PARTS, name: "forbidden"}))

    def test_corruption_outside_core_parts_rejected(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            path = root / "template.xlsx"
            with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_STORED) as archive:
                for name, value in CORE_PARTS.items():
                    archive.writestr(name, value)
                archive.writestr("xl/worksheets/sheet1.xml", b"UNIQUE_PAYLOAD_FOR_CRC_TEST")
            original = path.read_bytes()
            corrupted = original.replace(b"UNIQUE_PAYLOAD_FOR_CRC_TEST", b"BROKEN_PAYLOAD_FOR_CRC_TEST")
            self.assertNotEqual(original, corrupted)
            path.write_bytes(corrupted)
            self.assertIn("corrupt", " ".join(check_workbook(root, path.name)))

    def test_clean_duplicate_cannot_hide_corrupt_first_member(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            path = root / "template.xlsx"
            with warnings.catch_warnings():
                warnings.simplefilter("ignore", UserWarning)
                with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_STORED) as archive:
                    archive.writestr("xl/workbook.xml", "FIRST_WORKBOOK_COPY")
                    for name, value in CORE_PARTS.items():
                        archive.writestr(name, value)
            path.write_bytes(path.read_bytes().replace(b"FIRST_WORKBOOK_COPY", b"WRONG_WORKBOOK_COPY"))
            with zipfile.ZipFile(path) as archive:
                # Demonstrate the name-based testzip blind spot.
                self.assertIsNone(archive.testzip())
                with self.assertRaises(zipfile.BadZipFile):
                    archive.read(archive.infolist()[0])
            self.assertIn("duplicate", " ".join(check_workbook(root, path.name)))

    def test_lzma_corruption_returns_a_finding(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            path = root / "template.xlsx"
            with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_LZMA) as archive:
                for name, value in CORE_PARTS.items():
                    archive.writestr(name, value)
            with zipfile.ZipFile(path) as archive:
                entry = archive.infolist()[0]
                offset = entry.header_offset + 30 + len(entry.filename.encode()) + len(entry.extra)
            damaged = bytearray(path.read_bytes())
            # ZIP LZMA prefix: version (2), properties length (2), properties (5).
            damaged[offset + 4:offset + 9] = b"\xff" * 5
            path.write_bytes(damaged)
            with zipfile.ZipFile(path) as archive:
                with self.assertRaises(lzma.LZMAError):
                    archive.testzip()
            self.assertIn("not a readable workbook package", " ".join(check_workbook(root, path.name)))


class LabelDefaultsTests(unittest.TestCase):
    @unittest.skipUnless(shutil.which("node"), "Node is optional for local Python-only checks")
    def test_commands_without_policy(self):
        for script, arguments in (("labels-sync.mjs", ["--mode", "validate"]),
                                  ("labels-drift.mjs", ["--self-test"])):
            with self.subTest(script=script):
                result = subprocess.run(["node", str(ROOT / ".github/scripts" / script), *arguments],
                                        cwd=ROOT, capture_output=True, text=True, timeout=60)
                self.assertEqual(result.returncode, 0, result.stdout + result.stderr)


class TrafficMetricTests(unittest.TestCase):
    def test_old_and_current_schema_and_same_day_rerun(self):
        workflow = (ROOT / ".github/workflows/daily-traffic.yml").read_text()
        build = workflow.split("      - name: Build CSV history\n", 1)[1]
        source = textwrap.dedent(build.split("          python3 << 'PY'\n", 1)[1]
                                 .split("          PY\n", 1)[0])
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            fixtures = {
                "views": {"count": 5, "uniques": 3, "views": []},
                "clones": {"count": 2, "uniques": 1, "clones": []},
                "referrers": [], "paths": [],
                "repo": {"stargazers_count": 0, "open_issues_count": 7},
            }
            for name, value in fixtures.items():
                (root / (name + ".json")).write_text(json.dumps(value))
            (root / "data").mkdir()
            history = root / "data/traffic.csv"
            history.write_text("snapshot_date,open_issues\n2000-01-01,12\n")
            source = source.replace('"/tmp/', json.dumps(str(root) + "/")[:-1])
            previous = Path.cwd()
            try:
                os.chdir(root)
                exec(compile(source, "daily-traffic-build", "exec"), {})
                exec(compile(source, "daily-traffic-build", "exec"), {})
            finally:
                os.chdir(previous)
            with history.open(newline="") as handle:
                reader = csv.DictReader(handle)
                rows = list(reader)
                self.assertNotIn("open_issues", reader.fieldnames)
                self.assertIn("open_issues_and_prs", reader.fieldnames)
            self.assertEqual(len(rows), 2)
            self.assertEqual(rows[0]["open_issues_and_prs"], "12")
            self.assertEqual(rows[1]["open_issues_and_prs"], "7")
            self.assertEqual(rows[1]["views_14d"], "5")


if __name__ == "__main__":
    unittest.main()
