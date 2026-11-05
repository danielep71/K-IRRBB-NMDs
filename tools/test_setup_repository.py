"""The settings baseline script against a fake GitHub API; no network."""
from __future__ import annotations

import argparse
import copy
import io
import json
import tempfile
import unittest
from contextlib import redirect_stdout
from datetime import datetime, timezone
from pathlib import Path
from typing import Any
from unittest.mock import patch

import setup_repository as setup

NOW = datetime(2026, 11, 2, 7, 0, tzinfo=timezone.utc)
ROOT = Path(__file__).resolve().parents[1]


class FakeAPI:
    def __init__(self) -> None:
        self.repo: dict[str, Any] = {
            "full_name": "owner/repo", "visibility": "public", "default_branch": "main",
            "permissions": {"admin": True}, "allow_merge_commit": True, "allow_squash_merge": True,
            "allow_rebase_merge": True, "delete_branch_on_merge": False,
            "allow_update_branch": False, "has_projects": True, "has_downloads": True}
        self.rules: list[dict[str, Any]] = []
        self.writes: list[tuple[str, str, Any]] = []
        self.green = True
        self.ignore_patch = False
        self.reporting = False
        self.alerts = False
        self.fixes = False

    def not_found(self, endpoint: str) -> Any:
        raise setup.SetupError(f"GET {endpoint} failed (HTTP 404). Check gh authentication.")

    def switch(self, method: str, endpoint: str, attribute: str, payload: Any) -> Any:
        if method == "GET":
            if attribute == "reporting":
                return {"enabled": self.reporting}
            if attribute == "alerts":
                return None if self.alerts else self.not_found(endpoint)
            return {"enabled": self.fixes, "paused": False} if self.alerts else self.not_found(endpoint)
        self.writes.append((method, endpoint, payload))
        setattr(self, attribute, method == "PUT")
        return None

    def write(self, method: str, endpoint: str, payload: Any) -> Any:
        self.writes.append((method, endpoint, copy.deepcopy(payload)))
        if method == "POST":
            self.rules.append(dict(copy.deepcopy(payload), id=len(self.rules) + 1, source_type="Repository"))
        elif method == "PUT":
            number = int(endpoint.rsplit("/", 1)[1])
            self.rules = [dict(copy.deepcopy(payload), id=number, source_type="Repository")
                          if r["id"] == number else r for r in self.rules]
        elif method == "PATCH" and not self.ignore_patch:
            payload = copy.deepcopy(payload)
            analysis = payload.pop("security_and_analysis", None)
            if analysis:
                self.repo.setdefault("security_and_analysis", {}).update(analysis)
            self.repo.update(payload)
        return {}

    def __call__(self, method: str, endpoint: str, payload: Any = None) -> Any:
        switches = {"/private-vulnerability-reporting": "reporting",
                    "/vulnerability-alerts": "alerts", "/automated-security-fixes": "fixes"}
        for suffix, attribute in switches.items():
            if endpoint.endswith(suffix):
                return self.switch(method, endpoint, attribute, payload)
        if method != "GET":
            return self.write(method, endpoint, payload)
        if "/check-runs?" in endpoint:
            return {"check_runs": [{"name": "Repository integrity", "app": {"id": 15368},
                                    "status": "completed",
                                    "conclusion": "success" if self.green else "failure"}]}
        if "/branches/" in endpoint:
            return {"commit": {"sha": "abc"}}
        if "/rulesets?" in endpoint:
            return copy.deepcopy(self.rules)
        if "/rulesets/" in endpoint:
            number = int(endpoint.rsplit("/", 1)[1])
            return copy.deepcopy(next(r for r in self.rules if r["id"] == number))
        return copy.deepcopy(self.repo)


class SetupTests(unittest.TestCase):
    def setUp(self) -> None:
        self.api = FakeAPI()
        folder = tempfile.mkdtemp()
        self.addCleanup(__import__("shutil").rmtree, folder)
        self.report = Path(folder) / "evidence" / "settings.md"
        self.args = argparse.Namespace(repo="owner/repo", config=setup.CONFIG, apply=False,
                                       check=False, update_existing=False, report=self.report)

    def run_setup(self) -> tuple[int, str]:
        output = io.StringIO()
        with redirect_stdout(output):
            code = setup.run(self.args, self.api, NOW)
        return code, output.getvalue()

    def test_preview_never_writes_but_records_evidence(self) -> None:
        self.assertEqual(self.run_setup()[0], 0)
        self.assertEqual(self.api.writes, [])
        self.assertIn("- Result: **DRIFT**", self.report.read_text(encoding="utf-8"))

    def test_check_reports_drift_without_writing(self) -> None:
        self.args.check = True
        code, output = self.run_setup()
        self.assertEqual(code, 1)
        self.assertIn("DRIFT", output)
        self.assertEqual(self.api.writes, [])

    def test_apply_reads_back_and_a_second_run_changes_nothing(self) -> None:
        self.args.apply = True
        self.assertEqual(self.run_setup()[0], 0)
        methods = [w[0] for w in self.api.writes]
        self.assertEqual(methods, ["POST"] * 5 + ["PUT", "PUT", "PATCH", "PATCH"])
        self.assertEqual(self.api.writes[-1][2], {
            "delete_branch_on_merge": True, "allow_update_branch": True,
            "allow_rebase_merge": False, "has_projects": False, "has_downloads": False})
        self.api.writes.clear()
        self.args.apply, self.args.check = False, True
        self.assertEqual(self.run_setup()[0], 0)
        self.assertEqual(self.api.writes, [])
        self.assertIn("- Result: **VERIFIED**", self.report.read_text(encoding="utf-8"))

    def test_dependabot_security_updates_are_turned_off(self) -> None:
        # Their pull requests target the default branch, which takes work only
        # through the release branch.
        self.api.alerts = self.api.fixes = True
        self.args.apply = True
        self.assertEqual(self.run_setup()[0], 0)
        self.assertIn(("DELETE", "repos/owner/repo/automated-security-fixes", None), self.api.writes)
        self.assertFalse(self.api.fixes)

    def test_without_admin_rights_nothing_runs(self) -> None:
        self.args.apply = True
        self.api.repo["permissions"]["admin"] = False
        with self.assertRaisesRegex(setup.SetupError, "Admin rights"):
            self.run_setup()
        self.assertEqual(self.api.writes, [])

    def test_a_failed_required_check_blocks_every_write(self) -> None:
        self.args.apply = True
        self.api.green = False
        with self.assertRaisesRegex(setup.SetupError, "has not succeeded"):
            self.run_setup()
        self.assertEqual(self.api.writes, [])

    def test_disabled_merge_methods_block_every_write(self) -> None:
        self.args.apply = True
        for setting, message in (("allow_merge_commit", "Merge commits"),
                                 ("allow_squash_merge", "Squash merging")):
            with self.subTest(setting=setting):
                self.api = FakeAPI()
                self.api.repo[setting] = False
                with self.assertRaisesRegex(setup.SetupError, message):
                    self.run_setup()
                self.assertEqual(self.api.writes, [])

    def test_replacing_a_named_ruleset_needs_the_explicit_flag(self) -> None:
        rule = copy.deepcopy(setup.load_config(setup.CONFIG)["rulesets"][0])
        rule.update(id=1, source_type="Repository", bypass_actors=[])
        self.api.rules = [rule]
        self.args.apply = True
        with self.assertRaisesRegex(setup.SetupError, "named ruleset differs"):
            self.run_setup()
        self.assertEqual(self.api.writes, [])
        self.args.update_existing = True
        self.assertEqual(self.run_setup()[0], 0)
        self.assertEqual(self.api.rules[0]["bypass_actors"][0]["bypass_mode"], "pull_request")

    def test_unmanaged_rulesets_are_named_and_never_touched(self) -> None:
        other = {"name": "Legacy branch rule", "id": 10, "source_type": "Repository",
                 "target": "branch", "enforcement": "active",
                 "conditions": {"ref_name": {"include": ["refs/heads/release/**"], "exclude": []}},
                 "rules": [{"type": "deletion"}], "bypass_actors": []}
        self.api.rules = [copy.deepcopy(other)]
        code, output = self.run_setup()
        self.assertEqual(code, 0)
        self.assertIn('unmanaged ruleset "Legacy branch rule" '
                      "(branch: refs/heads/release/**; active)", output)
        self.args.apply = True
        self.run_setup()
        self.assertEqual(self.api.rules[0], other)
        self.assertFalse(any(endpoint.endswith("/rulesets/10") for _, endpoint, _ in self.api.writes))

    def test_private_repositories_skip_public_only_features(self) -> None:
        self.api.repo["visibility"] = "private"
        self.args.apply = True
        code, output = self.run_setup()
        self.assertEqual(code, 0)
        self.assertIn("apply only to public repositories", output)
        self.assertFalse(any("private-vulnerability" in endpoint or "security_and_analysis" in str(payload)
                             for _, endpoint, payload in self.api.writes))

    def test_dependent_security_features_are_rejected(self) -> None:
        config = setup.load_config(setup.CONFIG)
        for disabled, enabled, message in (("secret_scanning", "secret_scanning_push_protection",
                                            "Push protection"),
                                           ("vulnerability_alerts", "automated_security_fixes",
                                            "security updates")):
            broken = copy.deepcopy(config)
            broken["security"][disabled], broken["security"][enabled] = False, True
            with self.subTest(disabled=disabled), \
                    patch.object(Path, "read_text", return_value=json.dumps(broken)), \
                    self.assertRaisesRegex(setup.SetupError, message):
                setup.load_config(Path("settings.json"))

    def test_the_baseline_rejects_excluded_refs_and_bad_bypass_modes(self) -> None:
        config = setup.load_config(setup.CONFIG)
        excluded = copy.deepcopy(config)
        excluded["rulesets"][0]["conditions"]["ref_name"]["exclude"].append("x")
        exempt = copy.deepcopy(config)
        exempt["rulesets"][0]["bypass_actors"][0]["bypass_mode"] = "exempt"
        for broken, message in ((excluded, "exclude"), (exempt, "bypass actor")):
            with self.subTest(message=message), \
                    patch.object(Path, "read_text", return_value=json.dumps(broken)), \
                    self.assertRaisesRegex(setup.SetupError, message):
                setup.load_config(Path("settings.json"))

    def test_duplicate_live_names_are_rejected_before_any_write(self) -> None:
        name = setup.load_config(setup.CONFIG)["rulesets"][0]["name"]
        self.api.rules = [{"name": name, "id": 1}, {"name": name, "id": 2}]
        with self.assertRaisesRegex(setup.SetupError, "Ambiguous"):
            self.run_setup()
        self.assertEqual(self.api.writes, [])

    def test_read_back_detects_ignored_changes(self) -> None:
        self.args.apply = True
        self.api.ignore_patch = True
        with self.assertRaisesRegex(setup.SetupError, "Read-back"):
            self.run_setup()
        self.assertIn("- Result: **DRIFT**", self.report.read_text(encoding="utf-8"))

    def test_evidence_records_each_managed_setting(self) -> None:
        self.args.check = True
        self.run_setup()
        text = self.report.read_text(encoding="utf-8")
        self.assertIn("- Checked: **2026-11-02T07:00:00Z**", text)
        self.assertIn("| Protect traffic history | — | missing |", text)
        self.assertIn("| `allow_rebase_merge` | off | on |", text)
        self.assertIn("| `automated_security_fixes` | off | off |", text)

    def test_the_default_evidence_file_is_dated(self) -> None:
        self.assertEqual(setup.evidence_path(None, "check", NOW),
                         Path("test-results/repository-settings-check-20261102T070000Z.md"))


class BaselineTests(unittest.TestCase):
    def test_rule_order_is_not_drift_but_an_extra_rule_is(self) -> None:
        rule = copy.deepcopy(setup.load_config(setup.CONFIG)["rulesets"][0])
        reordered = copy.deepcopy(rule)
        reordered["rules"].reverse()
        self.assertEqual(setup.rule_spec(rule), setup.rule_spec(reordered))
        reordered["rules"].append({"type": "creation"})
        self.assertNotEqual(setup.rule_spec(rule), setup.rule_spec(reordered))

    def test_the_baseline_matches_the_branch_policy(self) -> None:
        rules = {r["name"]: r for r in setup.load_config(setup.CONFIG)["rulesets"]}

        def methods(name: str) -> list[str]:
            pull = next(rule for rule in rules[name]["rules"] if rule["type"] == "pull_request")
            return list(pull["parameters"]["allowed_merge_methods"])

        self.assertEqual(methods("Protect main"), ["merge"])
        self.assertEqual(methods("Protect release branches"), ["squash"])
        self.assertEqual([a["bypass_mode"] for a in rules["Protect main"]["bypass_actors"]], ["pull_request"])
        self.assertEqual([a["bypass_mode"] for a in rules["Protect release branches"]["bypass_actors"]],
                         ["always"])
        self.assertEqual(rules["Protect release tags"]["conditions"]["ref_name"]["include"], ["refs/tags/v*"])
        self.assertEqual(rules["Protect traffic history"]["conditions"]["ref_name"]["include"],
                         ["refs/heads/traffic-history"])

    def test_release_branches_cannot_be_deleted_even_by_an_admin(self) -> None:
        # Admins may bypass "Protect release branches", and merging an
        # integration pull request as an admin would let "delete head
        # branches on merge" remove the release branch. A deletion rule with
        # no bypass keeps it.
        keepers = [r for r in setup.load_config(setup.CONFIG)["rulesets"]
                   if r["target"] == "branch" and not r["bypass_actors"]
                   and "refs/heads/release/**" in r["conditions"]["ref_name"]["include"]
                   and {"type": "deletion"} in r["rules"]]
        self.assertEqual(len(keepers), 1)

    def test_the_required_check_is_the_integrity_job(self) -> None:
        workflow = (ROOT / ".github/workflows/static-checks.yml").read_text(encoding="utf-8")
        self.assertIn("    name: Repository integrity\n", workflow)
        contexts = {check["context"] for ruleset in setup.load_config(setup.CONFIG)["rulesets"]
                    for rule in ruleset["rules"] if rule["type"] == "required_status_checks"
                    for check in rule["parameters"]["required_status_checks"]}
        self.assertEqual(contexts, {"Repository integrity"})


class CommandLineTests(unittest.TestCase):
    def test_pagination_reads_beyond_the_first_page(self) -> None:
        calls: list[str] = []

        def call(method: str, endpoint: str) -> Any:
            calls.append(endpoint)
            return list(range(100)) if endpoint.endswith("&page=1") else [100]

        self.assertEqual(len(setup.pages(call, "repos/x/y/rulesets?includes_parents=false")), 101)
        self.assertEqual(len(calls), 2)

    def test_missing_reraises_errors_other_than_404(self) -> None:
        def forbidden(method: str, endpoint: str, payload: Any = None) -> Any:
            raise setup.SetupError(f"{method} {endpoint} failed (HTTP 403).")

        with self.assertRaisesRegex(setup.SetupError, "HTTP 403"):
            setup.missing(forbidden, "repos/a/b/vulnerability-alerts")

    def test_the_cli_call_is_explicit_and_hides_server_detail(self) -> None:
        ok = argparse.Namespace(returncode=0, stdout="{}", stderr="")
        with patch("setup_repository.subprocess.run", return_value=ok) as run:
            setup.api("PATCH", "repos/a/b", {"has_projects": False})
        command = run.call_args.args[0]
        self.assertIn("github.com", command)
        self.assertEqual(command[-2:], ["--input", "-"])
        self.assertEqual(run.call_args.kwargs["input"], '{"has_projects": false}')
        self.assertEqual(run.call_args.kwargs["encoding"], "utf-8")
        failed = argparse.Namespace(returncode=1, stdout="", stderr="secret token HTTP 403")
        with patch("setup_repository.subprocess.run", return_value=failed), \
                self.assertRaises(setup.SetupError) as caught:
            setup.api("GET", "repos/a/b")
        self.assertIn("HTTP 403", str(caught.exception))
        self.assertNotIn("secret token", str(caught.exception))

    def test_update_existing_requires_apply_and_the_repository_is_explicit(self) -> None:
        with redirect_stdout(io.StringIO()), patch("sys.stderr", io.StringIO()), \
                self.assertRaises(SystemExit):
            setup.parse_args(["--repo", "owner/repo", "--update-existing"])
        args = setup.parse_args(["--repo", "../repo"])
        with self.assertRaisesRegex(setup.SetupError, "explicit owner/name"):
            setup.run(args, FakeAPI(), NOW)


if __name__ == "__main__":
    unittest.main()
