"""Documentation conventions: the label catalog and milestone titles; no network access."""
from __future__ import annotations

import json
import re
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
# A milestone title as GitHub shows it: "vX.Y.Z - Title".
TITLE = re.compile(r"v\d+\.\d+\.\d+ - \S")
# The same version followed by an en or em dash instead of " - ".
DASHED_TITLE = re.compile(r"v\d+\.\d+\.\d+ [–—] ")


def read(path: str) -> str:
    return (ROOT / path).read_text(encoding="utf-8")


class LabelCatalogTests(unittest.TestCase):
    def test_labels_md_names_every_catalog_label(self) -> None:
        catalog = {label["name"] for label in json.loads(read(".github/labels.json"))["core"]}
        named = set(re.findall(r"`([^`]+)`", read("docs/LABELS.md")))
        self.assertEqual(catalog - named, set())


class MilestoneTitleTests(unittest.TestCase):
    def test_guides_use_the_github_title_form(self) -> None:
        for guide in sorted((ROOT / "docs" / "Milestones").glob("v*.md")):
            with self.subTest(guide=guide.name):
                heading = guide.read_text(encoding="utf-8").splitlines()[0]
                self.assertRegex(heading, r"^# \S+ " + re.escape(guide.stem) + r" - \S")

    def test_no_dashed_title_in_the_entry_documents(self) -> None:
        for path in ("README.md", "CONTRIBUTING.md", "INSTALLATION.md"):
            with self.subTest(path=path):
                self.assertIsNone(DASHED_TITLE.search(read(path)))
                self.assertRegex(read(path), TITLE)


if __name__ == "__main__":
    unittest.main()
