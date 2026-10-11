"""Offline coverage of metadata reconciliation; no GitHub calls."""
from __future__ import annotations

import copy
import importlib.util
import os
import subprocess
import unittest
from pathlib import Path
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
spec = importlib.util.spec_from_file_location("issue_metadata", ROOT / ".github/scripts/issue-metadata.py")
policy = importlib.util.module_from_spec(spec)
spec.loader.exec_module(policy)


class IssueMetadataTests(unittest.TestCase):
    def simulate(self, issue, milestone_state="open"):
        current = copy.deepcopy(issue)
        calls = []

        def api(endpoint, method="GET", payload=None, paginate=False):
            calls.append((endpoint, method, payload))
            if "/milestones/" in endpoint:
                return {"number": 1, "state": milestone_state}
            if method == "GET":
                return copy.deepcopy(current)
            if endpoint.endswith("/assignees"):
                current["assignees"] += [{"login": name} for name in payload["assignees"]]
            elif endpoint.endswith("/labels"):
                current["labels"] += [{"name": name} for name in payload["labels"]]
            elif method == "DELETE":
                current["labels"] = [label for label in current["labels"]
                                     if label["name"] != endpoint.rsplit("/", 1)[1]]
            elif method == "PATCH":
                current["milestone"] = {"number": payload["milestone"]}
            return None

        with patch.object(policy, "api", side_effect=api):
            policy.reconcile("owner/repo", 10, 1)
        return current, calls

    def test_open_and_closed_issues_get_missing_metadata(self):
        for state in ("open", "closed"):
            result, _ = self.simulate({"state": state, "assignees": [], "labels": [{"name": "bug"}], "milestone": None})
            self.assertEqual(result, {"state": state, "assignees": [{"login": "danielep71"}],
                                     "labels": [{"name": "bug"}, {"name": "P3"}], "milestone": {"number": 1}})

    def test_conflicting_priorities_keep_highest_and_preserve_existing_milestone(self):
        result, calls = self.simulate({"assignees": [{"login": "danielep71"}, {"login": "other"}],
                                      "labels": [{"name": p} for p in ("P3", "bug", "P1", "P2")],
                                      "milestone": {"number": 99, "state": "closed"}})
        self.assertEqual(result["labels"], [{"name": "bug"}, {"name": "P1"}])
        self.assertEqual(result["milestone"]["number"], 99)
        self.assertEqual(len(result["assignees"]), 2)
        self.assertFalse(any(method in {"POST", "PATCH"} for _, method, _ in calls))

    def test_correct_issue_is_idempotent_and_pr_is_ignored(self):
        issue = {"assignees": [{"login": "danielep71"}], "labels": [{"name": "P2"}], "milestone": {"number": 1}}
        for current in (issue, {"pull_request": {}}):
            result, calls = self.simulate(current)
            self.assertEqual(result, current)
            self.assertTrue(all(method == "GET" for _, method, _ in calls))

    def test_closed_default_milestone_is_rejected(self):
        with self.assertRaisesRegex(RuntimeError, "must be open"):
            self.simulate({"assignees": [{"login": "danielep71"}], "labels": [{"name": "P2"}],
                           "milestone": None}, "closed")


class IssueMetadataEntryTests(unittest.TestCase):
    """Event routing in main() and API failure reporting (#41)."""

    def run_main(self, **env):
        environment = {"GITHUB_REPOSITORY": "owner/repo", "ISSUE_MILESTONE_NUMBER": "2", **env}
        with patch.dict(os.environ, environment, clear=True):
            return policy.main()

    def test_pull_request_event_is_skipped_before_any_api_call(self):
        with patch.object(policy, "api", side_effect=AssertionError("no API call expected")), \
                patch.object(policy, "reconcile", side_effect=AssertionError("no reconcile expected")):
            self.assertEqual(self.run_main(ISSUE_NUMBER="38", ISSUE_IS_PULL_REQUEST="true"), 0)

    def test_genuine_issue_event_is_reconciled(self):
        with patch.object(policy, "reconcile") as reconcile:
            self.assertEqual(self.run_main(ISSUE_NUMBER="41", ISSUE_IS_PULL_REQUEST=""), 0)
        reconcile.assert_called_once_with("owner/repo", 41, 2)

    def test_sweep_reconciles_every_issue_and_no_pull_request(self):
        pages = [[{"number": 1}, {"number": 38, "pull_request": {}}], [{"number": 41}]]
        with patch.object(policy, "api", return_value=pages) as api, \
                patch.object(policy, "reconcile") as reconcile:
            self.assertEqual(self.run_main(), 0)
        api.assert_called_once_with("repos/owner/repo/issues?state=all&per_page=100", paginate=True)
        self.assertEqual([call.args[1] for call in reconcile.call_args_list], [1, 41])

    def test_api_failure_reports_status_and_redacts_tokens(self):
        token = "ghs_" + "A" * 36
        failed = subprocess.CompletedProcess(
            [], 1, stdout='{"message": "Not Found", "note": "' + token + '"}',
            stderr="gh: Not Found (HTTP 404)\n")
        with patch.object(policy.subprocess, "run", return_value=failed):
            with self.assertRaises(RuntimeError) as raised:
                policy.api("repos/owner/repo/issues/38")
        message = str(raised.exception)
        self.assertIn("GET repos/owner/repo/issues/38 failed (gh exit 1)", message)
        self.assertIn("HTTP 404", message)
        self.assertIn("[REDACTED]", message)
        self.assertNotIn(token, message)

    def test_failure_detail_is_bounded(self):
        self.assertLessEqual(len(policy.failure_detail("x" * 5000)), policy.DETAIL_LIMIT + 3)

    def test_workflow_passes_pull_request_flag_from_event_payload(self):
        workflow = (ROOT / ".github/workflows/issue-metadata.yml").read_text(encoding="utf-8")
        self.assertIn("ISSUE_IS_PULL_REQUEST: ${{ github.event.issue.pull_request && 'true' || '' }}",
                      workflow)


if __name__ == "__main__":
    unittest.main()
