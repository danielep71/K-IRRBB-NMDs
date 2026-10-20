"""Workflow action pins and the Dependabot target branch (#116); no network."""
from __future__ import annotations

import re
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
USES = re.compile(r"^\s*(?:-\s+)?uses:\s*(\S+)(.*)$", re.M)


class WorkflowPinTests(unittest.TestCase):
    def test_every_action_is_pinned_by_full_sha_with_its_version(self) -> None:
        # GitHub accepts both extensions for workflow files.
        folder = ROOT / ".github/workflows"
        workflows = sorted([*folder.glob("*.yml"), *folder.glob("*.yaml")])
        self.assertTrue(workflows)
        for path in workflows:
            for match in USES.finditer(path.read_text(encoding="utf-8")):
                reference, rest = match.groups()
                with self.subTest(workflow=path.name, uses=reference):
                    self.assertRegex(reference, r"^[\w.-]+/[\w./-]+@[0-9a-f]{40}$")
                    self.assertRegex(rest, r"#\s*v\d+(\.\d+)*\s*$")

    def test_dependabot_targets_the_active_release_branch(self) -> None:
        config = (ROOT / ".github/dependabot.yml").read_text(encoding="utf-8")
        target = re.search(r'^\s*target-branch:\s*"([^"]+)"', config, re.M)
        active = re.search(r"^\| Active branch \| `([^`]+)` \|$",
                           (ROOT / "RELEASING.md").read_text(encoding="utf-8"), re.M)
        self.assertIsNotNone(target)
        self.assertIsNotNone(active)
        assert target is not None and active is not None
        self.assertEqual(target.group(1), active.group(1))


if __name__ == "__main__":
    unittest.main()
