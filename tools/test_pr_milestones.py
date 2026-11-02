"""Offline coverage of the pull-request milestone guard and its workflows; no GitHub calls."""
from __future__ import annotations

import copy
import importlib.util
import re
import unittest
from pathlib import Path
from typing import Any
from unittest.mock import patch

from test_workflow_pins import jobs

ROOT = Path(__file__).resolve().parents[1]
spec = importlib.util.spec_from_file_location("pr_milestones", ROOT / ".github/scripts/pr-milestones.py")
assert spec is not None and spec.loader is not None  # a .py path always has a source loader
guard: Any = importlib.util.module_from_spec(spec)
spec.loader.exec_module(guard)


class FakeRepository:
    """Issues and pull requests in one numbered collection, as the GitHub API keeps them."""

    def __init__(self, items: list[dict[str, Any]], keep_milestone: bool = False) -> None:
        self.items = {item["number"]: copy.deepcopy(item) for item in items}
        self.keep_milestone = keep_milestone
        self.calls: list[tuple[str, str, Any]] = []

    def api(self, endpoint: str, method: str = "GET", payload: Any = None) -> Any:
        self.calls.append((endpoint, method, payload))
        if "/pulls?" in endpoint:
            page = int(re.search(r"[?&]page=(\d+)", endpoint).group(1))  # type: ignore[union-attr]
            pulls = [item for _, item in sorted(self.items.items()) if "pull_request" in item]
            return copy.deepcopy(pulls[(page - 1) * guard.PAGE_SIZE:page * guard.PAGE_SIZE])
        item = self.items[int(endpoint.rsplit("/", 1)[1])]
        if method == "PATCH" and not self.keep_milestone:
            item["milestone"] = payload["milestone"]
        return copy.deepcopy(item)

    def writes(self) -> list[tuple[str, str, Any]]:
        return [call for call in self.calls if call[1] != "GET"]


def pull(number: int, milestone: int | None = None) -> dict[str, Any]:
    return {"number": number, "pull_request": {}, "milestone": None if milestone is None else {"number": milestone}}


def issue(number: int, milestone: int | None = None) -> dict[str, Any]:
    return {"number": number, "milestone": None if milestone is None else {"number": milestone}}


class GuardTests(unittest.TestCase):
    def run_guard(self, repository: FakeRepository, number: int | None) -> int:
        with patch.object(guard, "api", side_effect=repository.api):
            return int(guard.reconcile("owner/repo", number))

    def test_event_clears_only_its_pull_request(self) -> None:
        repository = FakeRepository([pull(1, 2), pull(2, 2)])
        self.assertEqual(self.run_guard(repository, 1), 1)
        self.assertIsNone(repository.items[1]["milestone"])
        self.assertEqual(repository.items[2]["milestone"], {"number": 2})
        self.assertEqual(repository.writes(), [("repos/owner/repo/issues/1", "PATCH", {"milestone": None})])

    def test_pull_request_without_milestone_is_left_alone(self) -> None:
        repository = FakeRepository([pull(1)])
        self.assertEqual(self.run_guard(repository, 1), 0)
        self.assertEqual(repository.writes(), [])

    def test_an_issue_is_never_changed(self) -> None:
        repository = FakeRepository([issue(5, 2)])
        with self.assertRaisesRegex(RuntimeError, "is an issue"):
            self.run_guard(repository, 5)
        self.assertEqual(repository.writes(), [])
        self.assertEqual(repository.items[5]["milestone"], {"number": 2})

    def test_full_scan_reads_every_page_and_clears_each_milestone(self) -> None:
        items = [pull(n, 3 if n % 50 == 0 else None) for n in range(1, 251)] + [issue(300, 3)]
        repository = FakeRepository(items)
        self.assertEqual(self.run_guard(repository, None), 5)
        self.assertTrue(all(item["milestone"] is None for item in repository.items.values()
                            if "pull_request" in item))
        self.assertEqual(repository.items[300]["milestone"], {"number": 3})
        pages = [endpoint for endpoint, _, _ in repository.calls if "/pulls?" in endpoint]
        self.assertEqual(len(pages), 3)
        self.assertTrue(all("state=all" in page and "direction=asc" in page for page in pages))

    def test_a_milestone_that_survives_the_update_fails(self) -> None:
        repository = FakeRepository([pull(1, 2)], keep_milestone=True)
        with self.assertRaisesRegex(RuntimeError, "still has a milestone"):
            self.run_guard(repository, 1)

    def test_page_limit_fails_instead_of_stopping_silently(self) -> None:
        repository = FakeRepository([pull(n) for n in range(1, 4)])
        with patch.object(guard, "PAGE_SIZE", 1), patch.object(guard, "MAX_PAGES", 2):
            with self.assertRaisesRegex(RuntimeError, "more than 2 pages"):
                self.run_guard(repository, None)

    def test_failure_detail_redacts_tokens(self) -> None:
        detail = guard.failure_detail("HTTP 401 ghp_" + "a" * 36 + " " + "x" * 600)
        self.assertNotIn("ghp_", detail)
        self.assertIn("[REDACTED]", detail)
        self.assertLessEqual(len(detail), guard.DETAIL_LIMIT + 3)


def workflow(name: str) -> str:
    return (ROOT / ".github/workflows" / name).read_text(encoding="utf-8")


class GuardWorkflowTests(unittest.TestCase):
    def test_pull_request_code_is_never_checked_out(self) -> None:
        text = workflow("pr-milestones.yml")
        self.assertIn("pull_request_target:", text)
        self.assertIn("ref: ${{ github.event.repository.default_branch }}", text)
        self.assertNotIn("github.event.pull_request.head", text)
        self.assertRegex(text, r"(?m)^permissions: \{\}$")

    def test_the_job_has_only_the_permissions_it_needs(self) -> None:
        job = jobs(workflow("pr-milestones.yml"))["reconcile"]
        granted = re.findall(r"^      ([\w-]+): (\w+)$", job, re.M)
        self.assertEqual(sorted(granted), [("contents", "read"), ("issues", "read"),
                                           ("pull-requests", "write")])


class CodeQLWorkflowTests(unittest.TestCase):
    text = workflow("codeql.yml")

    def test_init_and_analyze_share_one_pin(self) -> None:
        pins = set(re.findall(r"uses: github/codeql-action/(?:init|analyze)@([0-9a-f]{40})", self.text))
        self.assertEqual(len(pins), 1, pins)

    def test_it_covers_the_branches_the_static_checks_cover(self) -> None:
        def branches(text: str) -> list[list[str]]:
            return [re.findall(r"- (\S+)", block) for block in
                    re.findall(r"\n  (?:push|pull_request):\n    branches:\n((?:      - .+\n)+)", text)]
        self.assertEqual(branches(self.text), branches(workflow("static-checks.yml")))

    def test_pull_requests_never_upload_results(self) -> None:
        analyze, pr = jobs(self.text)["analyze"], jobs(self.text)["analyze-pr"]
        self.assertIn("&& github.event_name != 'pull_request'", analyze)
        self.assertIn("security-events: write", analyze)
        self.assertIn("&& github.event_name == 'pull_request'", pr)
        self.assertNotIn("security-events", pr)
        self.assertIn("upload: never", pr)

    def test_both_jobs_scan_the_same_languages(self) -> None:
        languages = {job: re.findall(r"language: \[([^\]]+)\]", text) for job, text in jobs(self.text).items()}
        expected = ["actions, javascript-typescript, python"]
        self.assertEqual(languages, {"availability": [], "analyze": expected, "analyze-pr": expected})

    def test_analysis_waits_for_the_availability_check(self) -> None:
        for job in ("analyze", "analyze-pr"):
            with self.subTest(job=job):
                text = jobs(self.text)[job]
                self.assertIn("needs: availability", text)
                self.assertIn("needs.availability.outputs.enabled == 'true'", text)


if __name__ == "__main__":
    unittest.main()
