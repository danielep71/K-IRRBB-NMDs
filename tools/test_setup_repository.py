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
        self.paused = False
        self.environments: dict[str, dict[str, Any]] = {}
        self.branch_policies: dict[str, list[dict[str, Any]]] = {}
        self.next_policy = 100
        self.default_setup = "configured"
        self.scanning_status: int | None = None

    def not_found(self, endpoint: str) -> Any:
        raise setup.SetupError(f"GET {endpoint} failed (HTTP 404). Check gh authentication.")

    def switch(self, method: str, endpoint: str, attribute: str, payload: Any) -> Any:
        if method == "GET":
            if attribute == "reporting":
                return {"enabled": self.reporting}
            if attribute == "alerts":
                return None if self.alerts else self.not_found(endpoint)
            return ({"enabled": self.fixes, "paused": self.paused} if self.alerts
                    else self.not_found(endpoint))
        self.writes.append((method, endpoint, payload))
        setattr(self, attribute, method == "PUT")
        return None

    def environment(self, method: str, endpoint: str, payload: Any) -> Any:
        """GitHub's environment API: a PUT replaces every protection rule it does not name."""
        path = endpoint.split("?")[0].split("/environments/", 1)[1].split("/")
        name = path[0]
        if method != "GET":
            self.writes.append((method, endpoint, copy.deepcopy(payload)))
        if len(path) == 1:
            if method == "GET":
                return (copy.deepcopy(self.environments[name]) if name in self.environments
                        else self.not_found(endpoint))
            rules: list[dict[str, Any]] = []
            if payload.get("wait_timer"):
                rules.append({"id": 1, "type": "wait_timer", "wait_timer": payload["wait_timer"]})
            if payload.get("reviewers"):
                rules.append({"id": 2, "type": "required_reviewers",
                              "prevent_self_review": payload.get("prevent_self_review", False),
                              "reviewers": [{"type": r["type"], "reviewer": {"id": r["id"]}}
                                            for r in payload["reviewers"]]})
            policy = payload.get("deployment_branch_policy")
            if not (policy or {}).get("custom_branch_policies"):
                self.branch_policies.pop(name, None)
            self.environments[name] = {"name": name, "deployment_branch_policy": policy,
                                       "protection_rules": rules}
            return {}
        policy = self.environments.get(name, {}).get("deployment_branch_policy") or {}
        if not policy.get("custom_branch_policies"):
            return self.not_found(endpoint)
        current = self.branch_policies.setdefault(name, [])
        if method == "GET":
            return {"total_count": len(current), "branch_policies": copy.deepcopy(current)}
        if method == "POST":
            self.next_policy += 1
            current.append(dict(payload, id=self.next_policy))
        else:
            number = int(path[-1])
            self.branch_policies[name] = [p for p in current if p["id"] != number]
        return {}

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
        if "/environments/" in endpoint:
            return self.environment(method, endpoint, payload)
        if endpoint.endswith("/code-scanning/default-setup"):
            if method == "GET":
                if self.scanning_status is not None:
                    raise setup.SetupError(f"GET {endpoint} failed (HTTP {self.scanning_status}).")
                return {"state": self.default_setup, "languages": []}
            self.writes.append((method, endpoint, copy.deepcopy(payload)))
            self.default_setup = payload["state"]
            return {}
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
        # Rulesets, security features, the environment and its branch, the
        # CodeQL default setup, then the general settings.
        self.assertEqual(methods, ["POST"] * 5 + ["PUT", "PUT", "PATCH", "PUT", "POST", "PATCH", "PATCH"])
        self.assertEqual(self.api.writes[8:11], [
            ("PUT", "repos/owner/repo/environments/analytics",
             {"deployment_branch_policy": {"protected_branches": False, "custom_branch_policies": True}}),
            ("POST", "repos/owner/repo/environments/analytics/deployment-branch-policies",
             {"name": "main", "type": "branch"}),
            ("PATCH", "repos/owner/repo/code-scanning/default-setup", {"state": "not-configured"})])
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

    def wanting_security_updates(self) -> None:
        config = setup.load_config(setup.CONFIG)
        config["security"]["automated_security_fixes"] = True
        path = self.report.parent.parent / "settings.json"
        path.write_text(json.dumps(config), encoding="utf-8")
        self.args.config = path

    def test_paused_security_updates_are_drift_when_wanted(self) -> None:
        self.wanting_security_updates()
        self.args.apply = True
        self.assertEqual(self.run_setup()[0], 0)
        self.api.paused = True
        self.args.apply, self.args.check = False, True
        self.assertEqual(self.run_setup()[0], 1)
        self.assertIn("| `automated_security_fixes` | on | on (paused) |",
                      self.report.read_text(encoding="utf-8"))

    def test_security_updates_that_stay_paused_fail_the_read_back(self) -> None:
        self.wanting_security_updates()
        self.api.alerts = self.api.fixes = self.api.paused = True
        self.args.apply = True
        with self.assertRaisesRegex(setup.SetupError, "Read-back"):
            self.run_setup()

    def test_paused_security_updates_are_turned_off_when_unwanted(self) -> None:
        # Paused updates resume on their own, so "off" means not enabled at all.
        self.api.alerts = self.api.fixes = self.api.paused = True
        self.args.check = True
        self.assertEqual(self.run_setup()[0], 1)
        self.assertIn("| `automated_security_fixes` | off | on (paused) |",
                      self.report.read_text(encoding="utf-8"))
        self.args.check, self.args.apply = False, True
        self.assertEqual(self.run_setup()[0], 0)
        self.assertIn(("DELETE", "repos/owner/repo/automated-security-fixes", None), self.api.writes)

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
        rule.update(id=1, source_type="Repository", enforcement="disabled")
        self.api.rules = [rule]
        self.args.apply = True
        with self.assertRaisesRegex(setup.SetupError, "named ruleset differs"):
            self.run_setup()
        self.assertEqual(self.api.writes, [])
        self.args.update_existing = True
        self.assertEqual(self.run_setup()[0], 0)
        self.assertEqual(self.api.rules[0]["enforcement"], "active")

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

    def test_private_repositories_without_code_scanning_skip_the_default_setup(self) -> None:
        self.api.repo["visibility"] = "private"
        for status in (403, 404):
            with self.subTest(status=status):
                self.api.scanning_status = status
                self.api.writes.clear()
                code, output = self.run_setup()
                self.assertEqual(code, 0)
                self.assertIn("code scanning is not available", output)
                self.assertIn("| `code_scanning_default_setup` | not-configured | "
                              "not available (code scanning unavailable) |",
                              self.report.read_text(encoding="utf-8"))
                self.args.apply = True
                self.run_setup()
                self.args.apply = False
                self.assertFalse(any("code-scanning" in endpoint for _, endpoint, _ in self.api.writes))

    def test_code_scanning_errors_fail_on_public_repositories_and_for_other_statuses(self) -> None:
        for visibility, status in (("public", 403), ("public", 404), ("private", 401)):
            with self.subTest(visibility=visibility, status=status):
                self.api.repo["visibility"] = visibility
                self.api.scanning_status = status
                with self.assertRaisesRegex(setup.SetupError, rf"HTTP {status}"):
                    self.run_setup()

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

    def test_the_baseline_rejects_excluded_refs_and_bypass_actors(self) -> None:
        config = setup.load_config(setup.CONFIG)
        excluded = copy.deepcopy(config)
        excluded["rulesets"][0]["conditions"]["ref_name"]["exclude"].append("x")
        bypassed = copy.deepcopy(config)
        bypassed["rulesets"][0]["bypass_actors"] = [
            {"actor_id": 5, "actor_type": "RepositoryRole", "bypass_mode": "pull_request"}]
        for broken, message in ((excluded, "exclude"), (bypassed, "bypass actors")):
            with self.subTest(message=message), \
                    patch.object(Path, "read_text", return_value=json.dumps(broken)), \
                    self.assertRaisesRegex(setup.SetupError, message):
                setup.load_config(Path("settings.json"))

    def test_the_baseline_rejects_invalid_environments_and_codeql_setup(self) -> None:
        config = setup.load_config(setup.CONFIG)
        cases: list[tuple[str, Any, str]] = [
            ("environments", ["analytics"], "object keyed by name"),
            ("environments", {"analytics": {"deployment_branches": ["main"], "reviewers": []}},
             "Invalid baseline environment"),
            ("environments", {"../x": {"deployment_branches": ["main"]}}, "Invalid baseline environment"),
            ("environments", {"analytics": {"deployment_branches": []}}, "unique deployment branches"),
            ("environments", {"analytics": {"deployment_branches": ["main", "main"]}},
             "unique deployment branches"),
            ("environments", {"analytics": {"deployment_branches": [""]}}, "unique deployment branches"),
            ("code_scanning_default_setup", "configured", "default setup"),
            ("code_scanning_default_setup", None, "default setup")]
        for key, value, message in cases:
            broken = copy.deepcopy(config)
            broken[key] = value
            with self.subTest(value=value), \
                    patch.object(Path, "read_text", return_value=json.dumps(broken)), \
                    self.assertRaisesRegex(setup.SetupError, message):
                setup.load_config(Path("settings.json"))
        del config["environments"]
        with patch.object(Path, "read_text", return_value=json.dumps(config)), \
                self.assertRaisesRegex(setup.SetupError, "Unsupported configuration schema"):
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

    def analytics(self, policy: Any, branches: list[dict[str, Any]] | None = None) -> None:
        """An existing analytics environment with a reviewer, a wait timer and a secret."""
        self.api.environments["analytics"] = {
            "name": "analytics", "deployment_branch_policy": policy, "secrets_count": 1,
            "protection_rules": [
                {"id": 1, "type": "wait_timer", "wait_timer": 5},
                {"id": 2, "type": "required_reviewers", "prevent_self_review": True,
                 "reviewers": [{"type": "User", "reviewer": {"id": 7, "login": "me"}}]},
                {"id": 3, "type": "branch_policy"}]}
        if branches is not None:
            self.api.branch_policies["analytics"] = branches

    def environment_writes(self) -> list[tuple[str, str, Any]]:
        return [w for w in self.api.writes if "/environments/" in w[1]]

    def test_a_missing_environment_is_created_for_main_only(self) -> None:
        self.args.check = True
        self.assertEqual(self.run_setup()[0], 1)
        self.assertIn("| `analytics` deployment branches | main | missing |",
                      self.report.read_text(encoding="utf-8"))
        self.args.check, self.args.apply = False, True
        self.assertEqual(self.run_setup()[0], 0)
        self.assertEqual(self.api.branch_policies["analytics"],
                         [{"name": "main", "type": "branch", "id": 101}])
        self.assertIn("| `analytics` deployment branches | main | main |",
                      self.report.read_text(encoding="utf-8"))

    def test_a_wrong_policy_is_updated_keeping_reviewers_and_the_wait_timer(self) -> None:
        self.analytics({"protected_branches": True, "custom_branch_policies": False})
        self.args.check = True
        self.assertEqual(self.run_setup()[0], 1)
        self.assertIn("| `analytics` deployment branches | main | protected branches |",
                      self.report.read_text(encoding="utf-8"))
        self.args.check, self.args.apply = False, True
        self.assertEqual(self.run_setup()[0], 0)
        self.assertEqual(self.environment_writes()[0], (
            "PUT", "repos/owner/repo/environments/analytics",
            {"deployment_branch_policy": {"protected_branches": False, "custom_branch_policies": True},
             "wait_timer": 5, "prevent_self_review": True, "reviewers": [{"type": "User", "id": 7}]}))
        rules = self.api.environments["analytics"]["protection_rules"]
        self.assertEqual([r["wait_timer"] for r in rules if r["type"] == "wait_timer"], [5])
        self.assertEqual([(r["prevent_self_review"], r["reviewers"][0]["reviewer"]["id"])
                          for r in rules if r["type"] == "required_reviewers"],
                         [(True, 7)])
        # Secrets and variables have their own endpoints, which are never called.
        self.assertFalse(any("/secrets" in e or "/variables" in e for _, e, _ in self.api.writes))

    def test_extra_branch_policies_are_deleted(self) -> None:
        self.analytics({"protected_branches": False, "custom_branch_policies": True},
                       [{"id": 1, "name": "main", "type": "branch"},
                        {"id": 2, "name": "release/*", "type": "branch"},
                        {"id": 3, "name": "main", "type": "tag"}])
        self.args.apply = True
        self.assertEqual(self.run_setup()[0], 0)
        self.assertEqual(self.environment_writes(), [
            ("DELETE", "repos/owner/repo/environments/analytics/deployment-branch-policies/2", None),
            ("DELETE", "repos/owner/repo/environments/analytics/deployment-branch-policies/3", None)])
        self.assertEqual(self.api.branch_policies["analytics"], [{"id": 1, "name": "main", "type": "branch"}])

    def test_the_codeql_default_setup_is_drift_until_turned_off(self) -> None:
        self.args.check = True
        self.assertEqual(self.run_setup()[0], 1)
        self.assertIn("| `code_scanning_default_setup` | not-configured | configured |",
                      self.report.read_text(encoding="utf-8"))
        self.args.check, self.args.apply = False, True
        self.assertEqual(self.run_setup()[0], 0)
        self.assertIn(("PATCH", "repos/owner/repo/code-scanning/default-setup",
                       {"state": "not-configured"}), self.api.writes)
        self.assertEqual(self.api.default_setup, "not-configured")

    def test_a_verified_environment_and_codeql_setup_need_no_write(self) -> None:
        self.args.apply = True
        self.run_setup()
        self.api.writes.clear()
        self.analytics({"protected_branches": False, "custom_branch_policies": True},
                       [{"id": 1, "name": "main", "type": "branch"}])
        self.assertEqual(self.run_setup()[0], 0)
        self.assertEqual(self.api.writes, [])
        text = self.report.read_text(encoding="utf-8")
        self.assertIn("- Result: **VERIFIED**", text)
        self.assertIn("| `code_scanning_default_setup` | not-configured | not-configured |", text)

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
        # Task pull requests are squashed; levelling a release branch with
        # main after an integration is a merge-commit pull request.
        self.assertEqual(methods("Protect release branches"), ["squash", "merge"])
        self.assertEqual(rules["Protect release tags"]["conditions"]["ref_name"]["include"], ["refs/tags/v*"])
        self.assertEqual(rules["Protect traffic history"]["conditions"]["ref_name"]["include"],
                         ["refs/heads/traffic-history"])

    def test_no_ruleset_can_be_bypassed_and_release_checks_are_strict(self) -> None:
        rulesets = setup.load_config(setup.CONFIG)["rulesets"]
        self.assertEqual({r["name"]: r["bypass_actors"] for r in rulesets if r["bypass_actors"]}, {})
        for name in ("Protect main", "Protect release branches"):
            with self.subTest(ruleset=name):
                ruleset = next(r for r in rulesets if r["name"] == name)
                checks = next(rule for rule in ruleset["rules"]
                              if rule["type"] == "required_status_checks")
                self.assertTrue(checks["parameters"]["strict_required_status_checks_policy"])

    def test_only_release_branch_creation_skips_the_required_check(self) -> None:
        # docs/GOVERNANCE.md: a release branch is opened from a commit that
        # already passed the check on main; every other change needs it.
        rulesets = {r["name"]: r for r in setup.load_config(setup.CONFIG)["rulesets"]}
        skipped = {name: next(rule for rule in ruleset["rules"]
                              if rule["type"] == "required_status_checks"
                              )["parameters"]["do_not_enforce_on_create"]
                   for name, ruleset in rulesets.items()
                   if any(rule["type"] == "required_status_checks" for rule in ruleset["rules"])}
        self.assertEqual(skipped, {"Protect main": False, "Protect release branches": True})
        governance = (ROOT / "docs/GOVERNANCE.md").read_text(encoding="utf-8")
        self.assertIn('"Protect release branches" sets `do_not_enforce_on_create`',
                      " ".join(governance.split()))

    def test_release_branches_cannot_be_deleted_even_by_an_admin(self) -> None:
        # Admins may bypass "Protect release branches", and merging an
        # integration pull request as an admin would let "delete head
        # branches on merge" remove the release branch. A deletion rule with
        # no bypass keeps it.
        keepers = [r for r in setup.load_config(setup.CONFIG)["rulesets"]
                   if r["target"] == "branch" and not r["bypass_actors"]
                   and r["conditions"]["ref_name"]["include"] == ["refs/heads/release/**"]
                   and r["rules"] == [{"type": "deletion"}]]
        self.assertEqual(len(keepers), 1)

    def test_the_traffic_environment_deploys_only_from_main(self) -> None:
        workflow = (ROOT / ".github/workflows/daily-traffic.yml").read_text(encoding="utf-8")
        self.assertIn("    environment: analytics\n", workflow)
        config = setup.load_config(setup.CONFIG)
        self.assertEqual(config["environments"], {"analytics": {"deployment_branches": ["main"]}})
        self.assertEqual(config["code_scanning_default_setup"], "not-configured")

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
