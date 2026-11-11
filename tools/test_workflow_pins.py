"""Workflow pins, the Python toolchain and the Dependabot configuration; no network."""
from __future__ import annotations

import re
import subprocess
import tomllib
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
USES = re.compile(r"^\s*(?:-\s+)?uses:\s*(\S+)(.*)$", re.M)
JOB = re.compile(r"^  ([A-Za-z_][\w-]*):\s*$", re.M)
STEP = re.compile(r"^      - ", re.M)
PYTHON_VERSION = re.compile(r"^\s+python-version:\s*(.*?)\s*$", re.M)
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


def setup_python_versions(job: str) -> list[str]:
    """The python-version of each setup-python step of a job, whatever the key order.

    A step that sets no python-version, or sets it through python-version-file,
    yields "" so that the caller rejects it.
    """
    versions = []
    for step in STEP.split(job)[1:]:
        if not re.search(r"^\s*(?:-\s+)?uses:\s*actions/setup-python@", step, re.M):
            continue
        found = PYTHON_VERSION.findall(step)
        exact = len(found) == 1 and "python-version-file" not in step
        versions.append(found[0].strip('"') if exact and found[0].startswith('"') else "")
    return versions


def automation_files(tests: bool = False) -> list[str]:
    """Tracked workflows, workflow scripts and tools; the tests in tools/ only on request."""
    tracked = subprocess.run(["git", "-C", str(ROOT), "ls-files", "-z", ".github", "tools"],
                             capture_output=True, check=True).stdout.decode("utf-8").split("\0")
    return [name for name in tracked
            if re.search(r"\.(?:py|mjs|js|ya?ml|ps1|cmd|sh)$", name)
            and (tests or not re.match(r"tools/(?:.+/)?test_", name))]


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
                for version in setup_python_versions(text):
                    with self.subTest(workflow=path.name, job=job):
                        self.assertRegex(version, r"^\d+\.\d+\.\d+$")
                        if (path.name, job) == MINIMUM_JOB:
                            self.assertTrue(version.startswith(minimum.group(1) + "."), version)
                        else:
                            versions.add(version)
        self.assertEqual(len(versions), 1, f"one Python version outside the minimum job, found {versions}")

    def test_setup_python_steps_are_read_whatever_their_layout(self) -> None:
        uses = "      - uses: actions/setup-python@" + "0" * 40 + " # v1\n"
        cases = {
            '        with:\n          python-version: "3.13.18"\n': ["3.13.18"],
            '        with:\n          cache: pip\n          python-version: "3.13.18"\n': ["3.13.18"],
            '        with:\n          python-version: 3.13\n': [""],
            "        with:\n          python-version-file: .python-version\n": [""],
            "        with:\n          cache: pip\n": [""],
            "": [""],
        }
        for body, expected in cases.items():
            with self.subTest(body=body):
                self.assertEqual(setup_python_versions(uses + body), expected)
        named = ("      - name: Set up Python\n        uses: actions/setup-python@" + "0" * 40
                 + ' # v1\n        with:\n          python-version: "3.11.18"\n'
                 "      - run: python -V\n")
        self.assertEqual(setup_python_versions(named), ["3.11.18"])

    def test_every_job_that_runs_python_sets_it_up(self) -> None:
        for path in workflows():
            for job, text in jobs(path.read_text(encoding="utf-8")).items():
                if RUNS_PYTHON.search(text):
                    with self.subTest(workflow=path.name, job=job):
                        self.assertRegex(text, r"uses:\s*actions/setup-python@")

    def test_the_minimum_job_exists(self) -> None:
        workflow, job = MINIMUM_JOB
        self.assertIn(job, jobs((ROOT / ".github/workflows" / workflow).read_text(encoding="utf-8")))


class RequiredCheckTests(unittest.TestCase):
    """The required check, Repository integrity, runs the portable gates on every pull request."""

    TEXT = (ROOT / ".github/workflows/static-checks.yml").read_text(encoding="utf-8")

    def test_it_triggers_on_pull_requests_and_pushes_to_main_and_release_branches(self) -> None:
        for event in ("pull_request", "push"):
            with self.subTest(event=event):
                block = re.search(rf"^  {event}:\n    branches:\n((?:      - .+\n)+)", self.TEXT, re.M)
                assert block is not None
                self.assertEqual(re.findall(r"- (.+)", block.group(1)), ["main", '"release/**"'])
        self.assertNotRegex(self.TEXT, r"(?m)^\s+(?:paths|paths-ignore|branches-ignore):")

    def test_the_integrity_job_runs_the_portable_gates_on_the_exact_commit(self) -> None:
        integrity = [text for text in jobs(self.TEXT).values()
                     if re.search(r"^    name: Repository integrity$", text, re.M)]
        self.assertEqual(len(integrity), 1)
        job = integrity[0]
        self.assertNotRegex(job, r"(?m)^    if:")
        self.assertIn('test "$(git rev-parse HEAD)" = "$CANDIDATE_SHA"', job)
        self.assertIn("args=(--ci)", job)
        self.assertIn('args+=(--base "$PR_BASE_SHA")', job)
        self.assertRegex(job, r'(?m)^          python tools/check\.py "\$\{args\[@\]\}"$')


class QualityToolTests(unittest.TestCase):
    def test_every_tool_and_workflow_script_is_linted_and_type_checked(self) -> None:
        config = tomllib.loads((ROOT / "pyproject.toml").read_text(encoding="utf-8"))
        typed = [str(path).rstrip("/") for path in config["tool"]["mypy"]["files"]]
        workflow = (ROOT / ".github/workflows/static-checks.yml").read_text(encoding="utf-8")
        command = re.search(r"^          ruff check ([^|]+?) 2>&1 \| tee test-results/ruff\.txt$", workflow, re.M)
        assert command is not None
        linted = [path.rstrip("/") for path in command.group(1).split()]
        tracked = [name for name in automation_files(tests=True) if name.endswith(".py")]
        self.assertIn(".github/scripts/pr-milestones.py", tracked)
        for name in tracked:
            with self.subTest(file=name):
                self.assertTrue(any(name.startswith(path + "/") for path in typed), typed)
                self.assertTrue(any(name.startswith(path + "/") for path in linted), linted)


class ApiVersionTests(unittest.TestCase):
    # The traffic export was written and verified against 2026-03-10; every
    # other script requests 2022-11-28. A change to either is deliberate.
    EXPECTED = {".github/workflows/daily-traffic.yml": "2026-03-10"}
    DEFAULT = "2022-11-28"

    def test_each_script_requests_its_recorded_rest_api_version(self) -> None:
        versions: dict[str, set[str]] = {}
        for name in automation_files():
            text = (ROOT / name).read_text(encoding="utf-8")
            if "X-GitHub-Api-Version" not in text:
                continue
            found = set(re.findall(r"(?:X-GitHub-Api-Version\W+|API_VERSION\W+)(\d{4}-\d{2}-\d{2})", text))
            with self.subTest(file=name):
                self.assertTrue(found, "sends X-GitHub-Api-Version without a literal version")
            versions[name] = found
        self.assertIn(".github/workflows/daily-traffic.yml", versions)
        for name, found in versions.items():
            with self.subTest(file=name):
                self.assertEqual(found, {self.EXPECTED.get(name, self.DEFAULT)})


class TriageAutomationTests(unittest.TestCase):
    """docs/GOVERNANCE.md: no automation changes an issue's priority or milestone."""

    # Commands and API calls that change the labels or milestone of an
    # existing issue or pull request.
    FORBIDDEN = re.compile(r"\bgh\s+(?:issue|pr)\s+edit\b|--(?:add|remove)-label\b"
                           r"|\b(?:updateIssue|updatePullRequest|addLabelsToLabelable"
                           r"|removeLabelsFromLabelable|clearLabelsFromLabelable)\b")
    # An endpoint of one issue (issues/<number>, its labels included).
    ONE_ISSUE = re.compile(r"issues/(?:\d|\{|\$)")
    # Changes only pull requests and refuses an issue (test_pr_milestones.py).
    PULL_REQUEST_GUARD = ".github/scripts/pr-milestones.py"

    @staticmethod
    def commands(text: str) -> list[str]:
        """Lines with shell continuations joined."""
        return re.sub(r"\s*\\\n\s*", " ", text).splitlines()

    def test_no_automation_changes_labels_or_milestones_of_existing_items(self) -> None:
        single: list[str] = []
        for name in automation_files():
            text = (ROOT / name).read_text(encoding="utf-8")
            for line in self.commands(text):
                with self.subTest(file=name, line=line.strip()):
                    self.assertIsNone(self.FORBIDDEN.search(line))
                    if re.search(r"--milestone(?![\w-])", line):
                        # Filing a new issue (the traffic alert) is allowed.
                        self.assertRegex(line, r"\bgh issue create\b")
            if self.ONE_ISSUE.search(text):
                single.append(name)
        self.assertEqual(single, [self.PULL_REQUEST_GUARD])

    def test_the_guard_only_clears_a_milestone_after_checking_for_a_pull_request(self) -> None:
        text = (ROOT / self.PULL_REQUEST_GUARD).read_text(encoding="utf-8")
        writes = re.findall(r'api\([^)]*"(PATCH|POST|PUT|DELETE)"[^)]*\)', text)
        self.assertEqual(writes, ["PATCH"])
        self.assertIn('api(endpoint, "PATCH", {"milestone": None})', text)
        self.assertLess(text.index('if "pull_request" not in item:'), text.index('"PATCH"'))

    def test_the_scan_sees_a_forbidden_call(self) -> None:
        for line in ("gh issue edit 5 --milestone v1", "gh pr edit 5 --add-label P1",
                     "gh api graphql -f query='mutation { updateIssue(input: {}) { clientMutationId } }'"):
            with self.subTest(line=line):
                self.assertIsNotNone(self.FORBIDDEN.search(line))
        self.assertIsNotNone(self.ONE_ISSUE.search("repos/o/r/issues/${number}/labels"))
        self.assertEqual(self.commands("gh issue create --title t \\\n  --milestone m"),
                         ["gh issue create --title t --milestone m"])


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

    def test_every_update_is_assigned_and_labelled_for_triage(self) -> None:
        config = self.CONFIG.read_text(encoding="utf-8")
        entries = re.split(r"^  - package-ecosystem:", config, flags=re.M)[1:]
        self.assertEqual(len(entries), 2)
        for entry in entries:
            with self.subTest(entry=entry.split("\n", 1)[0].strip()):
                self.assertRegex(entry, r'\n    assignees:\n      - "danielep71"\n')
                self.assertRegex(entry, r'\n    labels:\n      - "ci"\n      - "P3"\n')

    def test_codeql_action_steps_update_together(self) -> None:
        config = self.CONFIG.read_text(encoding="utf-8")
        actions = re.split(r"^  - package-ecosystem:", config, flags=re.M)[1]
        self.assertTrue(actions.lstrip().startswith('"github-actions"'))
        self.assertRegex(actions, r'\n    groups:\n      codeql-action:\n        patterns:\n'
                                  r'          - "github/codeql-action\*"\n')


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


class ConcurrencyTests(unittest.TestCase):
    def test_static_checks_cancel_only_superseded_pull_request_runs(self) -> None:
        text = (ROOT / ".github/workflows/static-checks.yml").read_text(encoding="utf-8")
        block = re.search(r"^concurrency:\n  group: (.+)\n  cancel-in-progress: (.+)\n", text, re.M)
        assert block is not None
        group, cancel = block.groups()
        self.assertEqual(group, "static-checks-${{ github.event_name == 'pull_request' && "
                                "format('pr-{0}', github.event.pull_request.number) || "
                                "format('run-{0}', github.run_id) }}")
        self.assertEqual(cancel, "${{ github.event_name == 'pull_request' }}")
        self.assertNotIn("github.ref", group)

    def test_a_label_reconcile_is_never_cancelled_midway(self) -> None:
        # labels-sync.yml promises an exact post-run match; only pull-request
        # validation runs may be cancelled by a newer commit.
        text = (ROOT / ".github/workflows/labels-sync.yml").read_text(encoding="utf-8")
        block = re.search(r"^concurrency:\n  group: (.+)\n  cancel-in-progress: (.+)\n", text, re.M)
        assert block is not None
        self.assertEqual(block.group(2), "${{ github.event_name == 'pull_request' }}")
        self.assertNotRegex(text, r"(?m)^\s+concurrency:")


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
