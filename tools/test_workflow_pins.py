"""Workflow pins, the Python toolchain and the Dependabot configuration; no network."""
from __future__ import annotations

import re
import subprocess
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


class MutatingJobTests(unittest.TestCase):
    """Jobs that change the repository run only from the default branch."""

    DEFAULT_BRANCH = "github.ref == format('refs/heads/{0}', github.event.repository.default_branch)"

    def test_label_reconcile_runs_only_from_the_default_branch(self) -> None:
        text = (ROOT / ".github/workflows/labels-sync.yml").read_text(encoding="utf-8")
        job = jobs(text)["reconcile"]
        condition = re.search(r"^    if: >-\n((?:      .+\n)+)", job, re.M)
        assert condition is not None, "the reconcile job has no multi-line if condition"
        clauses = [line.strip() for line in condition.group(1).splitlines()]
        self.assertEqual(clauses, ["github.event_name != 'pull_request' &&", self.DEFAULT_BRANCH])
        self.assertIn("\n  workflow_dispatch:", text)
    def test_every_update_is_assigned_and_labelled_for_triage(self) -> None:
        config = self.CONFIG.read_text(encoding="utf-8")
        entries = re.split(r"^  - package-ecosystem:", config, flags=re.M)[1:]
        self.assertEqual(len(entries), 2)
        for entry in entries:
            with self.subTest(entry=entry.split("\n", 1)[0].strip()):
                self.assertRegex(entry, r'\n    assignees:\n      - "danielep71"\n')
                self.assertRegex(entry, r'\n    labels:\n      - "ci"\n      - "P3"\n')


class ActiveBranchTests(unittest.TestCase):
    # Only these files may name the active release branch: Dependabot needs
    # the literal name, RELEASING.md holds the current-status table, and
    # CHANGELOG.md keeps history and its comparison link.
    ALLOWED = {".github/dependabot.yml", "RELEASING.md", "CHANGELOG.md"}

    def test_the_branch_name_lives_only_in_its_allowed_files(self) -> None:
        active = re.search(r"^\| Active branch \| `([^`]+)` \|$",
                           (ROOT / "RELEASING.md").read_text(encoding="utf-8"), re.M)
        assert active is not None
        tracked = subprocess.run(["git", "-C", str(ROOT), "ls-files", "-z"], capture_output=True,
                                 check=True).stdout.decode("utf-8").split("\0")
        found = []
        for name in tracked:
            path = ROOT / name
            if not name or name in self.ALLOWED or not path.is_file():
                continue
            try:
                text = path.read_text(encoding="utf-8")
            except UnicodeDecodeError:
                continue
            if active.group(1) in text:
                found.append(name)
        self.assertEqual(found, [], "name the active release branch through RELEASING.md instead")


if __name__ == "__main__":
    unittest.main()
