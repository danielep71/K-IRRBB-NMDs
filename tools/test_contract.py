"""Regression fixtures for the repository's foundational source contracts."""
from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from check_source import OPTION_EXPLICIT, VB_NAME, check_changelog, check_version


class SourceContracts(unittest.TestCase):
    def test_component_header(self) -> None:
        sample = 'Attribute VB_Name = "CORE_Rates"\nOption Explicit\n'
        self.assertEqual(VB_NAME.findall(sample), ["CORE_Rates"])
        self.assertIsNotNone(OPTION_EXPLICIT.search(sample))

    def test_missing_changelog_fails(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            self.assertTrue(check_changelog(Path(directory)))

    def test_unreleased_changelog_is_valid(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / "CHANGELOG.md").write_text(
                "## [Unreleased]\n\n[Unreleased]: https://example.org/repo\n", encoding="utf-8"
            )
            self.assertEqual(check_changelog(root), [])
            self.assertEqual(check_version(root), [])


if __name__ == "__main__":
    unittest.main()
