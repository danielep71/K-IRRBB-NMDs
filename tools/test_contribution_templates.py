"""The feature form and the pull request template ask for what CONTRIBUTING and GOVERNANCE require."""
from __future__ import annotations

import re
import unittest
from pathlib import Path

from check_issues import form_field

ROOT = Path(__file__).resolve().parents[1]
# CONTRIBUTING.md: a feature issue defines the methodological basis, observable
# contract, compatibility, non-goals and validation; each maps to a form field.
FEATURE_FIELDS = {
    "methodological basis": "methodology",
    "observable contract": "contract",
    "compatibility": "compatibility",
    "non-goals": "boundary",
    "validation": "validation",
}


def read(path: str) -> str:
    return (ROOT / path).read_text(encoding="utf-8")


class FeatureFormTests(unittest.TestCase):
    def test_contributing_names_the_mapped_elements(self) -> None:
        row = next(line for line in read("CONTRIBUTING.md").splitlines() if "Model or feature change" in line)
        for element in FEATURE_FIELDS:
            self.assertIn(element, row)

    def test_each_element_is_a_required_feature_field(self) -> None:
        form = read(".github/ISSUE_TEMPLATE/feature.yml")
        for element, field in FEATURE_FIELDS.items():
            with self.subTest(element=element):
                item = form_field(form, field)
                self.assertIsNotNone(item, f"feature.yml has no {field} field")
                assert item is not None
                self.assertRegex(item, r"(?m)^\s*validations:\s*\n\s*required:\s*true\s*$")


class PullRequestTemplateTests(unittest.TestCase):
    def test_template_names_the_definition_of_done_command(self) -> None:
        done = next(line for line in read("docs/GOVERNANCE.md").splitlines() if "A change is complete only when" in line)
        command = re.search(r"`(python tools/check\.py[^`]*)`", done)
        assert command is not None
        self.assertEqual(command.group(1), "python tools/check.py --ci")
        template = read(".github/PULL_REQUEST_TEMPLATE.md")
        self.assertEqual(re.findall(r"`(python tools/check\.py[^`]*)`", template), [command.group(1)])


if __name__ == "__main__":
    unittest.main()
