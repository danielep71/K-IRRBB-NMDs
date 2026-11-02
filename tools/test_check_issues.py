"""Offline coverage of the issue triage check, its workflow and the alert milestone rule."""
from __future__ import annotations

import contextlib
import io
import json
import re
import shutil
import tempfile
import textwrap
import unittest
from pathlib import Path
from typing import Any

import check_issues
from check_issues import Issue
from test_workflow_pins import jobs

ROOT = Path(__file__).resolve().parents[1]
REPOSITORY = "owner/repo"
PERSON = "danielep71"
CATALOG = check_issues.catalog_labels(ROOT)
READY = "## Problem\n\nx\n\n## Dependencies\n\nNone.\n\n- Target end date: {} (Europe/Rome)\n"


def issue(number: int, *labels: str, body: str | None = None, target: str = "2026-11-10",
          assignees: tuple[str, ...] = (PERSON,), milestone: str | None = "v0.2.0",
          due: str | None = "2026-11-20", milestone_state: str = "open",
          state: str = "open") -> Issue:
    text = READY.format(target) if body is None else body
    return Issue(number, state, labels or ("P3",), assignees, milestone, text, due, milestone_state)


def rules(issues: list[Issue], states: dict[int, str] | None = None,
          pulls: frozenset[int] = frozenset()) -> list[tuple[int, str]]:
    known = states if states is not None else {i.number: i.state for i in issues}
    return [(f.number, f.rule) for f in
            check_issues.audit(issues, known, REPOSITORY, CATALOG, pulls, assignee=PERSON)]


def depends(*lines: str, target: str = "2026-11-10") -> str:
    return ("## Problem\n\nx\n\n## Dependencies\n\n" + "\n".join(lines)
            + f"\n\n## Planning\n\n- Target end date: {target}\n")


class MetadataTests(unittest.TestCase):
    def test_a_triaged_issue_has_no_findings(self) -> None:
        self.assertEqual(rules([issue(1)]), [])

    def test_assignee_priority_and_milestone_are_required(self) -> None:
        self.assertEqual(rules([issue(1, assignees=())]), [(1, "assignee")])
        self.assertEqual(rules([issue(1, "documentation")]), [(1, "priority")])
        self.assertEqual(rules([issue(1, "P2", "P3")]), [(1, "priority")])
        self.assertEqual(rules([issue(1, milestone=None, due=None)]), [(1, "milestone")])
        self.assertEqual(rules([issue(1, milestone_state="closed")]), [(1, "milestone")])

    def test_the_milestone_needs_a_due_date(self) -> None:
        self.assertEqual(rules([issue(1, due=None)]), [(1, "milestone-due")])

    def test_p1_is_reserved_for_bugs(self) -> None:
        self.assertEqual(rules([issue(1, "P1", "enhancement")]), [(1, "p1-reserved")])
        self.assertEqual(rules([issue(1, "P1", "bug")]), [])

    def test_labels_outside_the_catalog_are_reported(self) -> None:
        self.assertEqual(rules([issue(1, "P3", "wishlist")]), [(1, "catalog")])

    def test_closed_issues_are_not_audited(self) -> None:
        self.assertEqual(rules([issue(1, "P1", "P2", state="closed", assignees=())]), [])


class TargetDateTests(unittest.TestCase):
    def test_the_target_must_be_a_real_date_before_the_due_date(self) -> None:
        for target in ("", "soon", "2026-02-30", "10/11/2026"):
            with self.subTest(target=target):
                self.assertEqual(rules([issue(1, target=target)]), [(1, "target-date")])
        self.assertEqual(rules([issue(1, target="2026-11-20")]), [(1, "target-date")])
        self.assertEqual(rules([issue(1, target="2026-11-19")]), [])

    def test_the_issue_form_field_is_read(self) -> None:
        body = ("### Problem\n\nx\n\n### Target end date (Europe/Rome)\n\n2026-11-10\n\n"
                "### Dependencies\n\nNone.\n")
        self.assertEqual(rules([issue(1, body=body)]), [])
        empty = body.replace("2026-11-10", "_No response_")
        self.assertEqual(rules([issue(1, body=empty)]), [(1, "target-date")])

    def test_an_invalid_planning_line_does_not_fall_back_to_the_form(self) -> None:
        body = ("- Target end date: later\n\n### Target end date\n\n2026-11-10\n\n"
                "## Dependencies\n\nNone.\n")
        self.assertEqual(rules([issue(1, body=body)]), [(1, "target-date")])

    def test_crlf_bodies_are_read(self) -> None:
        self.assertEqual(rules([issue(1, body=READY.format("2026-11-10").replace("\n", "\r\n"))]), [])


class DependencyTests(unittest.TestCase):
    def test_a_dependencies_section_is_required(self) -> None:
        body = "## Problem\n\nx\n\n- Target end date: 2026-11-10\n"
        self.assertEqual(rules([issue(1, body=body)]), [(1, "dependencies")])

    def test_blocked_matches_open_prerequisites(self) -> None:
        open_dep = issue(1, body=depends("- #2 must land first"))
        self.assertEqual(rules([open_dep, issue(2)]), [(1, "blocked")])
        blocked = issue(1, "P3", "blocked", body=depends("- #2 must land first"))
        self.assertEqual(rules([blocked, issue(2)]), [])
        self.assertEqual(rules([blocked, issue(2, state="closed")]), [(1, "blocked")])

    def test_external_conditions_count_as_blockers(self) -> None:
        body = depends("External condition: the publisher releases the data")
        self.assertEqual(rules([issue(1, body=body)]), [(1, "blocked")])
        self.assertEqual(rules([issue(1, "P3", "blocked", body=body)]), [])

    def test_missing_numbers_are_reported_and_pull_requests_are_not_prerequisites(self) -> None:
        self.assertEqual(rules([issue(1, body=depends("After #7."))]), [(1, "dependencies")])
        self.assertEqual(rules([issue(1, body=depends("After #7."))], pulls=frozenset({7})), [])

    def test_issue_urls_and_references_outside_the_section_are_handled(self) -> None:
        url = f"https://github.com/{REPOSITORY}/issues/2"
        self.assertEqual(rules([issue(1, body=depends(f"- {url}")), issue(2)]), [(1, "blocked")])
        body = "See #2.\n\n" + depends("None.")
        self.assertEqual(rules([issue(1, body=body), issue(2)]), [])


class ConfigTests(unittest.TestCase):
    def copy(self) -> Path:
        folder = Path(tempfile.mkdtemp())
        self.addCleanup(shutil.rmtree, folder)
        shutil.copytree(ROOT / ".github", folder / ".github")
        return folder

    def test_the_repository_configuration_passes(self) -> None:
        self.assertEqual(check_issues.config_findings(ROOT), [])

    def test_each_form_rule_is_enforced(self) -> None:
        cases = {
            'labels: ["bug", "P2"]': ('labels: ["bug"]', "must set exactly the priority P2"),
            "    id: dependencies\n": ("    id: notes\n", "has no dependencies field"),
            "    id: target_end_date\n": ("    id: target\n", "has no target_end_date field"),
        }
        for old, (new, message) in cases.items():
            with self.subTest(change=new):
                root = self.copy()
                form = root / ".github/ISSUE_TEMPLATE/bug.yml"
                form.write_text(form.read_text(encoding="utf-8").replace(old, new), encoding="utf-8")
                problems = check_issues.config_findings(root)
                self.assertTrue(any(message in problem for problem in problems), problems)

    def test_blank_issues_must_stay_disabled(self) -> None:
        root = self.copy()
        config = root / ".github/ISSUE_TEMPLATE/config.yml"
        config.write_text(config.read_text(encoding="utf-8").replace("false", "true"), encoding="utf-8")
        self.assertIn("ISSUE_TEMPLATE/config.yml must set blank_issues_enabled: false",
                      check_issues.config_findings(root))


class CommandLineTests(unittest.TestCase):
    def run_main(self, items: list[dict[str, Any]]) -> tuple[int, str, dict[str, Any]]:
        folder = Path(tempfile.mkdtemp())
        self.addCleanup(shutil.rmtree, folder)
        source, output = folder / "issues.json", folder / "out/triage.json"
        source.write_text(json.dumps(items), encoding="utf-8")
        stdout = io.StringIO()
        with contextlib.redirect_stdout(stdout):
            code = check_issues.main(["--input", str(source), "--repo-name", REPOSITORY,
                                      "--output", str(output), "--summary", str(folder / "sum.md")])
        return code, stdout.getvalue(), json.loads(output.read_text(encoding="utf-8"))

    def item(self, number: int, **changes: Any) -> dict[str, Any]:
        base: dict[str, Any] = {
            "number": number, "state": "open", "labels": [{"name": "P3"}],
            "assignees": [{"login": PERSON}], "body": READY.format("2026-11-10"),
            "milestone": {"title": "v0.2.0", "due_on": "2026-11-20T08:00:00Z", "state": "open"}}
        return {**base, **changes}

    def test_findings_fail_the_run_and_are_reported(self) -> None:
        code, out, report = self.run_main([self.item(1), self.item(2, labels=[]),
                                           self.item(3, pull_request={}, labels=[])])
        self.assertEqual(code, 1)
        self.assertIn("| #2 | `priority` |", out)
        self.assertEqual(report["status"], "fail")
        self.assertEqual([f["number"] for f in report["findings"]], [2])

    def test_a_clean_repository_passes(self) -> None:
        code, out, report = self.run_main([self.item(1)])
        self.assertEqual((code, report["status"]), (0, "pass"))
        self.assertIn("open issues audited: 1", out)


def workflow(name: str) -> str:
    return (ROOT / ".github/workflows" / name).read_text(encoding="utf-8")


class TriageWorkflowTests(unittest.TestCase):
    def test_it_runs_on_issue_and_milestone_events_with_read_only_permissions(self) -> None:
        text = workflow("issue-triage.yml")
        for event in ("issues:", "milestone:", "workflow_dispatch:"):
            self.assertIn(f"\n  {event}", text)
        granted = re.findall(r"^\s+([\w-]+): (read|write)$", text, re.M)
        self.assertTrue(granted)
        self.assertTrue(all(level == "read" for _, level in granted), granted)
        self.assertIn("--summary \"$GITHUB_STEP_SUMMARY\"", jobs(text)["audit"])

    def test_no_workflow_changes_priorities_or_milestones_of_issues(self) -> None:
        self.assertFalse((ROOT / ".github/workflows/issue-metadata.yml").exists())
        self.assertFalse((ROOT / ".github/scripts/issue-metadata.py").exists())
        for path in sorted((ROOT / ".github").rglob("*")):
            if path.is_file() and path.suffix in {".yml", ".yaml", ".py", ".mjs"}:
                with self.subTest(file=path.name):
                    self.assertNotIn("ISSUE_MILESTONE_NUMBER", path.read_text(encoding="utf-8"))


class AlertMilestoneTests(unittest.TestCase):
    """The traffic alert files its issue under the open milestone due soonest after today."""

    choose: Any
    source: str

    @classmethod
    def setUpClass(cls) -> None:
        step = workflow("daily-traffic.yml").split(
            "      - name: Create alert issue on spike or new referrer\n", 1)[1]
        source = textwrap.dedent(step.split("          python3 << 'PY'\n", 1)[1].split("          PY\n", 1)[0])
        function = re.search(r"^def active_milestone\(.*?(?=^\S)", source, re.M | re.S)
        assert function is not None
        namespace: dict[str, Any] = {}
        exec(compile(function.group(0), "daily-traffic-alert", "exec"), namespace)
        cls.choose = staticmethod(namespace["active_milestone"])
        cls.source = source

    def test_the_earliest_due_date_after_today_wins(self) -> None:
        milestones = [{"title": "late", "state": "open", "due_on": "2027-01-10T08:00:00Z"},
                      {"title": "undated", "state": "open", "due_on": None},
                      {"title": "today", "state": "open", "due_on": "2026-11-02T08:00:00Z"},
                      {"title": "next", "state": "open", "due_on": "2026-11-20T08:00:00Z"}]
        self.assertEqual(self.choose(milestones, "2026-11-02")["title"], "next")

    def test_no_milestone_due_after_today_fails_clearly(self) -> None:
        for milestones in ([], [{"title": "x", "state": "open", "due_on": None}],
                           [{"title": "x", "state": "open", "due_on": "2026-11-01T00:00:00Z"}]):
            with self.subTest(milestones=milestones), self.assertRaisesRegex(RuntimeError, "due date"):
                self.choose(milestones, "2026-11-02")

    def test_it_is_chosen_before_any_early_exit(self) -> None:
        self.assertLess(self.source.index("selected = active_milestone("),
                        self.source.index("sys.exit(0)"))

    def test_the_alert_body_meets_the_triage_rules(self) -> None:
        self.assertIn('"## Dependencies"', self.source)
        self.assertIn('f"- Target end date: {target} (Europe/Rome)"', self.source)
        self.assertIn("datetime.date.fromisoformat(milestone_due) - datetime.timedelta(days=1)",
                      self.source)


if __name__ == "__main__":
    unittest.main()
