"""Workflow pins, the Python toolchain and the Dependabot configuration; no network."""
from __future__ import annotations

import re
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
USES = re.compile(r"^\s*(?:-\s+)?uses:\s*(\S+)(.*)$", re.M)
JOB = re.compile(r"^  ([A-Za-z_][\w-]*):\s*$", re.M)
SETUP_PYTHON = re.compile(r"uses:\s*actions/setup-python@\S+[^\n]*\n\s+with:\s*\n\s+python-version:\s*\"([^\"]+)\"")
RUNS_PYTHON = re.compile(r"^\s*(?:-\s+)?(?:run:\s*)?python3?\s", re.M)
# The job that checks the documented minimum Python instead of the CI version.
MINIMUM_JOB = ("static-checks.yml", "minimum-python")


def workflows() -> list[Path]:
    # GitHub accepts both extensions for workflow files.
    folder = ROOT / ".github/workflows"
    return sorted([*folder.glob("*.yml"), *folder.glob("*.yaml")])


def jobs(text: str) -> dict[str, str]:
    """The text of each job under the top-level jobs: key."""
    body = text.split("\njobs:\n", 1)[1]
    starts = list(JOB.finditer(body))
    return {m.group(1): body[m.end():(starts[i + 1].start() if i + 1 < len(starts) else len(body))]
            for i, m in enumerate(starts)}


class WorkflowPinTests(unittest.TestCase):
    def test_every_action_is_pinned_by_full_sha_with_its_version(self) -> None:
        self.assertTrue(workflows())
        for path in workflows():
            for match in USES.finditer(path.read_text(encoding="utf-8")):
                reference, rest = match.groups()
                with self.subTest(workflow=path.name, uses=reference):
                    self.assertRegex(reference, r"^[\w.-]+/[\w./-]+@[0-9a-f]{40}$")
                    self.assertRegex(rest, r"#\s*v\d+(\.\d+)*\s*$")


class PythonToolchainTests(unittest.TestCase):
    def test_python_versions_are_exact_and_shared(self) -> None:
        # A floating version lets a new release change CI without a commit.
        minimum = re.search(r'^python_version = "(\d+\.\d+)"$',
                            (ROOT / "pyproject.toml").read_text(encoding="utf-8"), re.M)
        assert minimum is not None
        versions: set[str] = set()
        for path in workflows():
            for job, text in jobs(path.read_text(encoding="utf-8")).items():
                for version in SETUP_PYTHON.findall(text):
                    with self.subTest(workflow=path.name, job=job):
                        self.assertRegex(version, r"^\d+\.\d+\.\d+$")
                        if (path.name, job) == MINIMUM_JOB:
                            self.assertTrue(version.startswith(minimum.group(1) + "."), version)
                        else:
                            versions.add(version)
        self.assertEqual(len(versions), 1, f"one Python version outside the minimum job, found {versions}")

    def test_every_job_that_runs_python_sets_it_up(self) -> None:
        for path in workflows():
            for job, text in jobs(path.read_text(encoding="utf-8")).items():
                if RUNS_PYTHON.search(text):
                    with self.subTest(workflow=path.name, job=job):
                        self.assertRegex(text, r"uses:\s*actions/setup-python@")

    def test_the_minimum_job_exists(self) -> None:
        workflow, job = MINIMUM_JOB
        self.assertIn(job, jobs((ROOT / ".github/workflows" / workflow).read_text(encoding="utf-8")))


class DependabotTests(unittest.TestCase):
    CONFIG = ROOT / ".github/dependabot.yml"

    def test_each_pinned_dependency_is_covered_once(self) -> None:
        config = self.CONFIG.read_text(encoding="utf-8")
        entries = re.findall(r'package-ecosystem:\s*"([^"]+)"\s*\n\s*directory:\s*"([^"]+)"', config)
        self.assertEqual(sorted(entries), [("github-actions", "/"), ("pip", "/tools")])

    def test_every_update_targets_the_active_release_branch(self) -> None:
        targets = re.findall(r'^\s*target-branch:\s*"([^"]+)"', self.CONFIG.read_text(encoding="utf-8"), re.M)
        active = re.search(r"^\| Active branch \| `([^`]+)` \|$",
                           (ROOT / "RELEASING.md").read_text(encoding="utf-8"), re.M)
        assert active is not None
        self.assertEqual(len(targets), 2)
        self.assertEqual(set(targets), {active.group(1)})


if __name__ == "__main__":
    unittest.main()
