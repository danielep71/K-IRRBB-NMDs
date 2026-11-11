#!/usr/bin/env python3
"""Compare, apply or check the repository settings baseline through the GitHub CLI.

The baseline is docs/repository-settings.json. Without a mode the script previews
the difference and never writes; --apply writes the difference and reads it back;
--check verifies read-only and exits 1 on drift. Every run writes a dated
evidence file. It needs an authenticated gh with admin rights on the repository,
and it changes only the rulesets and environments the baseline names; it never
touches environment secrets or variables.
"""
from __future__ import annotations

import argparse
import json
import os
import re
import subprocess
import sys
from collections.abc import Callable
from datetime import datetime, timezone
from pathlib import Path
from typing import Any
from urllib.parse import quote

ROOT = Path(__file__).resolve().parents[1]
CONFIG = ROOT / "docs/repository-settings.json"
EVIDENCE_DIR = Path("test-results")
RULE_KEYS = ("name", "target", "enforcement", "conditions", "rules", "bypass_actors")
GENERAL_KEYS = {"delete_branch_on_merge", "allow_update_branch", "allow_merge_commit",
                "allow_squash_merge", "allow_rebase_merge", "has_projects", "has_downloads"}
# Security features and whether GitHub offers them free only on public repositories.
SECURITY_PUBLIC_ONLY = {
    "private_vulnerability_reporting": True,
    "vulnerability_alerts": False,
    "automated_security_fixes": False,
    "secret_scanning": True,
    "secret_scanning_push_protection": True,
}
ANALYSIS_KEYS = ("secret_scanning", "secret_scanning_push_protection")
CONFIG_KEYS = {"schema_version", "repository", "security", "environments",
               "code_scanning_default_setup", "rulesets"}
# Only main may deploy: custom branch policies, not "protected branches".
CUSTOM_BRANCHES = {"protected_branches": False, "custom_branch_policies": True}

Json = Any
Call = Callable[..., Json]
Plan = list[tuple[str, str, Json]]


class SetupError(Exception):
    pass


def api(method: str, endpoint: str, payload: Json = None) -> Json:
    command = ["gh", "api", "--hostname", "github.com", "--method", method,
               "-H", "Accept: application/vnd.github+json",
               "-H", "X-GitHub-Api-Version: 2022-11-28", endpoint]
    if payload is not None:
        command += ["--input", "-"]
    # Never inherit HTTP debug logging: it can expose authentication headers.
    env = dict(os.environ)
    env.pop("GH_DEBUG", None)
    # gh writes UTF-8; never decode with the locale code page (cp1252 on Windows).
    result = subprocess.run(command, input=json.dumps(payload) if payload is not None else None,
                            text=True, encoding="utf-8", capture_output=True, env=env, check=False)
    if result.returncode:
        # The CLI's stderr may hold server or credential detail; report the status only.
        match = re.search(r"HTTP (\d{3})", result.stderr)
        status = match.group(1) if match else "unknown"
        raise SetupError(f"{method} {endpoint} failed (HTTP {status}). Check gh authentication, "
                         "repository access and admin rights. Nothing is rolled back; run the "
                         "preview to inspect partial changes.")
    return json.loads(result.stdout) if result.stdout.strip() else None


def pages(call: Call, endpoint: str, key: str | None = None) -> list[Json]:
    result: list[Json] = []
    page = 1
    while True:
        separator = "&" if "?" in endpoint else "?"
        value = call("GET", f"{endpoint}{separator}per_page=100&page={page}")
        batch = value[key] if key else value
        result.extend(batch)
        if len(batch) < 100:
            return result
        page += 1


def canonical(value: Json) -> Json:
    """Rule arrays are sets for this baseline: ignore the API's ordering, not its values."""
    if isinstance(value, dict):
        return {k: canonical(v) for k, v in sorted(value.items())}
    if isinstance(value, list):
        return sorted((canonical(x) for x in value), key=lambda x: json.dumps(x, sort_keys=True))
    return value


def rule_spec(rule: dict[str, Json]) -> Json:
    # Ignore response metadata, but keep unknown rules and parameters in the comparison.
    return canonical({k: rule.get(k, [] if k == "bypass_actors" else None) for k in RULE_KEYS})


def check_ruleset(rule: dict[str, Json]) -> None:
    if set(rule) != set(RULE_KEYS) or rule["target"] not in ("branch", "tag") \
            or rule["enforcement"] != "active":
        raise SetupError(f"Invalid baseline ruleset: {rule.get('name')!r}.")
    if rule["conditions"]["ref_name"]["exclude"]:
        raise SetupError(f"Baseline ruleset {rule['name']!r} must not exclude refs.")
    # No one, admins included, may bypass the documented route.
    if rule["bypass_actors"]:
        raise SetupError(f"Baseline ruleset {rule['name']!r} must not have bypass actors.")


def check_environments(environments: Json) -> None:
    if not isinstance(environments, dict):
        raise SetupError("Baseline environments must be an object keyed by name.")
    for name, environment in environments.items():
        if not re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9 ._-]*", name) \
                or not isinstance(environment, dict) or set(environment) != {"deployment_branches"}:
            raise SetupError(f"Invalid baseline environment: {name!r}.")
        branches = environment["deployment_branches"]
        if not isinstance(branches, list) or not branches \
                or any(not isinstance(b, str) or not b.strip() for b in branches) \
                or len(set(branches)) != len(branches):
            raise SetupError(f"Baseline environment {name!r} must list unique deployment branches.")


def load_config(path: Path) -> dict[str, Json]:
    value = json.loads(path.read_text(encoding="utf-8"))
    if set(value) != CONFIG_KEYS or value["schema_version"] != 1:
        raise SetupError("Unsupported configuration schema.")
    security = value["security"]
    if set(security) != set(SECURITY_PUBLIC_ONLY) or any(type(v) is not bool for v in security.values()):
        raise SetupError("The configuration must hold exactly the five boolean security features.")
    if security["secret_scanning_push_protection"] and not security["secret_scanning"]:
        raise SetupError("Push protection requires secret scanning.")
    if security["automated_security_fixes"] and not security["vulnerability_alerts"]:
        raise SetupError("Dependabot security updates require vulnerability alerts.")
    general = value["repository"]
    if set(general) != GENERAL_KEYS or any(type(v) is not bool for v in general.values()):
        raise SetupError("The configuration must hold exactly the seven boolean repository settings.")
    check_environments(value["environments"])
    # .github/workflows/codeql.yml uploads its own results; the two cannot run together.
    if value["code_scanning_default_setup"] != "not-configured":
        raise SetupError('The CodeQL default setup must be "not-configured".')
    names = [rule.get("name") for rule in value["rulesets"]]
    for rule in value["rulesets"]:
        check_ruleset(rule)
    if not names or len(set(names)) != len(names):
        raise SetupError("Baseline rulesets must have unique names.")
    return dict(value)


def inspect(call: Call, repo: str, config: dict[str, Json],
            report: bool = True) -> tuple[dict[str, Json], Plan]:
    root = f"repos/{repo}"
    metadata = call("GET", root)
    if metadata.get("archived") or metadata.get("disabled"):
        raise SetupError("The repository is archived or disabled.")
    if not metadata.get("permissions", {}).get("admin"):
        raise SetupError("Admin rights on the repository are required.")
    current = pages(call, root + "/rulesets?includes_parents=false")
    inherited = pages(call, root + "/rulesets?includes_parents=true")
    print(f"Repository: {metadata['full_name']}; visibility: {metadata['visibility']}; "
          f"default branch: {metadata['default_branch']}")
    if report and any(r.get("source_type") != "Repository" for r in inherited):
        print("NOTE: inherited rulesets also apply; this script neither changes nor checks them.")
    if report:
        report_unmanaged(call, root, current, config)
    plan: Plan = []
    for desired in config["rulesets"]:
        matches = [r for r in current if r["name"] == desired["name"]]
        if len(matches) > 1:
            raise SetupError(f"Ambiguous duplicate ruleset name: {desired['name']}")
        if not matches:
            plan.append(("POST", root + "/rulesets", desired))
        else:
            endpoint = root + "/rulesets/" + str(matches[0]["id"])
            if rule_spec(call("GET", endpoint)) != rule_spec(desired):
                plan.append(("PUT", endpoint, desired))
    plan += security_plan(call, root, metadata, config, report)
    for name, environment in config["environments"].items():
        plan += environment_plan(call, root, name, environment["deployment_branches"])
    desired_setup = config["code_scanning_default_setup"]
    if default_setup(call, root) != desired_setup:
        plan.append(("PATCH", root + "/code-scanning/default-setup", {"state": desired_setup}))
    changes = {k: v for k, v in config["repository"].items() if metadata.get(k) != v}
    # Protection first, then the general settings.
    if changes:
        plan.append(("PATCH", root, changes))
    return metadata, plan


def missing(call: Call, endpoint: str) -> tuple[Json, bool]:
    """GET an endpoint that answers 404 when its feature is off; True when it does."""
    try:
        return call("GET", endpoint), False
    except SetupError as error:
        if "(HTTP 404)" in str(error):
            return None, True
        raise


def security_fixes(call: Call, root: str) -> tuple[bool, bool]:
    """Whether Dependabot security updates are enabled, and whether GitHub has paused them."""
    value, absent = missing(call, root + "/automated-security-fixes")
    if absent:
        return False, False
    value = value or {}
    return bool(value.get("enabled")), bool(value.get("paused"))


def security_state(call: Call, root: str, metadata: dict[str, Json], feature: str) -> bool:
    """Whether a feature is in effect: paused security updates are enabled, not in effect."""
    if feature == "private_vulnerability_reporting":
        return bool((call("GET", root + "/private-vulnerability-reporting") or {}).get("enabled"))
    if feature == "vulnerability_alerts":
        return not missing(call, root + "/vulnerability-alerts")[1]
    if feature == "automated_security_fixes":
        enabled, paused = security_fixes(call, root)
        return enabled and not paused
    analysis = metadata.get("security_and_analysis") or {}
    return bool((analysis.get(feature) or {}).get("status") == "enabled")


def matches(call: Call, root: str, metadata: dict[str, Json], feature: str, desired: bool) -> bool:
    """Wanted on: the feature must be in effect. Wanted off: security updates must not even be
    enabled, since paused ones resume on their own."""
    if feature == "automated_security_fixes" and not desired:
        return not security_fixes(call, root)[0]
    return security_state(call, root, metadata, feature) == desired


def security_plan(call: Call, root: str, metadata: dict[str, Json], config: dict[str, Json],
                  report: bool) -> Plan:
    """Plan the security features; public-only ones are skipped, with a note, elsewhere."""
    public = metadata.get("visibility") == "public"
    skipped = [name for name, only in SECURITY_PUBLIC_ONLY.items()
               if only and not public and config["security"][name]]
    if skipped and report:
        print(f"NOTE: {', '.join(skipped)} apply only to public repositories here; they are not "
              f"applied while the repository is {metadata.get('visibility')}.")
    plan: Plan = []
    analysis: dict[str, dict[str, str]] = {}
    # Alerts before security updates, which depend on them.
    for feature in ("private_vulnerability_reporting", "vulnerability_alerts",
                    "automated_security_fixes", *ANALYSIS_KEYS):
        desired = config["security"][feature]
        if SECURITY_PUBLIC_ONLY[feature] and not public:
            continue
        if matches(call, root, metadata, feature, desired):
            continue
        if feature in ANALYSIS_KEYS:
            analysis[feature] = {"status": "enabled" if desired else "disabled"}
        else:
            plan.append(("PUT" if desired else "DELETE", root + "/" + feature.replace("_", "-"), None))
    if analysis:
        plan.append(("PATCH", root, {"security_and_analysis": analysis}))
    return plan


def environment_state(call: Call, root: str, name: str) -> tuple[Json, list[Json] | None]:
    """The environment (None when missing) and its custom branch policies (None when the
    environment does not restrict deployments to custom branch policies)."""
    endpoint = f"{root}/environments/{quote(name, safe='')}"
    environment, absent = missing(call, endpoint)
    if absent:
        return None, None
    if (environment.get("deployment_branch_policy") or {}).get("custom_branch_policies") is not True:
        return environment, None
    return environment, pages(call, endpoint + "/deployment-branch-policies", "branch_policies")


def carried_protection(environment: dict[str, Json]) -> dict[str, Json]:
    """The wait timer and required reviewers a PUT would otherwise clear."""
    kept: dict[str, Json] = {}
    for rule in environment.get("protection_rules") or []:
        if rule.get("type") == "wait_timer":
            kept["wait_timer"] = rule.get("wait_timer", 0)
        elif rule.get("type") == "required_reviewers":
            kept["prevent_self_review"] = bool(rule.get("prevent_self_review"))
            kept["reviewers"] = [{"type": r["type"], "id": r["reviewer"]["id"]}
                                 for r in rule.get("reviewers") or []]
    return kept


def environment_plan(call: Call, root: str, name: str, branches: list[str]) -> Plan:
    """Restrict an environment's deployments to exactly the baseline branches."""
    endpoint = f"{root}/environments/{quote(name, safe='')}"
    environment, policies = environment_state(call, root, name)
    plan: Plan = []
    if environment is None or environment.get("deployment_branch_policy") != CUSTOM_BRANCHES:
        payload = {"deployment_branch_policy": dict(CUSTOM_BRANCHES)}
        if environment is not None:
            payload.update(carried_protection(environment))
        plan.append(("PUT", endpoint, payload))
    current = policies or []
    kept = {p["name"] for p in current
            if p.get("type", "branch") == "branch" and p["name"] in branches}
    plan += [("POST", endpoint + "/deployment-branch-policies", {"name": branch, "type": "branch"})
             for branch in branches if branch not in kept]
    plan += [("DELETE", f"{endpoint}/deployment-branch-policies/{p['id']}", None) for p in current
             if p.get("type", "branch") != "branch" or p["name"] not in branches]
    return plan


def describe_environment(call: Call, root: str, name: str) -> str:
    environment, policies = environment_state(call, root, name)
    if environment is None:
        return "missing"
    if policies is None:
        return "protected branches" if environment.get("deployment_branch_policy") else "all branches"
    return ", ".join(sorted(p["name"] if p.get("type", "branch") == "branch"
                            else f"{p['name']} ({p['type']})" for p in policies)) or "none"


def default_setup(call: Call, root: str) -> str:
    return str((call("GET", root + "/code-scanning/default-setup") or {}).get("state"))


def report_unmanaged(call: Call, root: str, current: list[Json], config: dict[str, Json]) -> None:
    """Name the repository rulesets outside the baseline; they stay active alongside it."""
    managed = {rule["name"] for rule in config["rulesets"]}
    for ruleset in current:
        if ruleset.get("source_type", "Repository") != "Repository" or ruleset["name"] in managed:
            continue
        detail = call("GET", f"{root}/rulesets/{ruleset['id']}")
        refs = ", ".join(((detail.get("conditions") or {}).get("ref_name") or {}).get("include") or [])
        print(f'NOTE: unmanaged ruleset "{ruleset["name"]}" '
              f'({detail.get("target", "unknown target")}: {refs or "no ref conditions"}; '
              f'{detail.get("enforcement", "unknown enforcement")}) stays active. This script '
              "never changes or deletes it; delete it in the repository settings if a managed "
              "ruleset replaces it.")


def cell(value: Json) -> str:
    return "on" if value is True else "off" if value is False else str(value)


def evidence(call: Call, repo: str, config: dict[str, Json], metadata: dict[str, Json],
             plan: Plan, mode: str, now: datetime) -> str:
    """A dated Markdown record of the managed settings as GitHub reports them."""
    root = f"repos/{repo}"
    public = metadata.get("visibility") == "public"
    differs = {endpoint for method, endpoint, _ in plan if method == "PUT" and "/rulesets/" in endpoint}
    created = {payload["name"] for method, endpoint, payload in plan
               if method == "POST" and endpoint.endswith("/rulesets")}
    current = pages(call, root + "/rulesets?includes_parents=false")
    inherited = [r for r in pages(call, root + "/rulesets?includes_parents=true")
                 if r.get("source_type") != "Repository"]
    managed = {rule["name"] for rule in config["rulesets"]}
    lines = ["# Repository settings evidence", "",
             f"- Checked: **{now.strftime('%Y-%m-%dT%H:%M:%SZ')}** with "
             f"`tools/setup_repository.py` ({mode})",
             f"- Repository: `{metadata['full_name']}`; visibility: {metadata['visibility']}; "
             f"default branch: `{metadata['default_branch']}`",
             f"- Result: **{'VERIFIED' if not plan else 'DRIFT'}**", "",
             "## General settings", "", "| Setting | Baseline | Observed |", "| --- | --- | --- |"]
    lines += [f"| `{key}` | {cell(value)} | {cell(metadata.get(key))} |"
              for key, value in config["repository"].items()]
    lines += ["", "## Security features", "", "| Feature | Baseline | Observed |", "| --- | --- | --- |"]
    for feature, value in config["security"].items():
        if SECURITY_PUBLIC_ONLY[feature] and not public:
            observed = "not applied (public repositories only)"
        elif feature == "automated_security_fixes" and security_fixes(call, root) == (True, True):
            observed = "on (paused)"
        else:
            observed = cell(security_state(call, root, metadata, feature))
        lines.append(f"| `{feature}` | {cell(value)} | {observed} |")
    lines += ["", "## Environments and code scanning", "", "| Setting | Baseline | Observed |",
              "| --- | --- | --- |"]
    lines += [f"| `{name}` deployment branches | {', '.join(environment['deployment_branches'])} | "
              f"{describe_environment(call, root, name)} |"
              for name, environment in config["environments"].items()]
    lines.append(f"| `code_scanning_default_setup` | {config['code_scanning_default_setup']} | "
                 f"{default_setup(call, root)} |")
    lines += ["", "## Managed rulesets", "", "| Ruleset | ID | Status |", "| --- | --- | --- |"]
    for rule in config["rulesets"]:
        match = next((r for r in current if r["name"] == rule["name"]), None)
        if rule["name"] in created or match is None:
            lines.append(f"| {rule['name']} | — | missing |")
        else:
            status = "differs" if f"{root}/rulesets/{match['id']}" in differs else "matches"
            lines.append(f"| {rule['name']} | {match['id']} | {status} |")
    unmanaged = [r for r in current if r.get("source_type", "Repository") == "Repository"
                 and r["name"] not in managed]
    lines += ["", "## Other rules", "",
              f"- Unmanaged repository rulesets: {', '.join(r['name'] for r in unmanaged) or 'none'}",
              f"- Inherited rulesets: {', '.join(r['name'] for r in inherited) or 'none'}",
              "", "Settings outside the baseline are not recorded here."]
    return "\n".join(lines) + "\n"


def evidence_path(report: Path | None, mode: str, now: datetime) -> Path:
    return report or EVIDENCE_DIR / f"repository-settings-{mode}-{now.strftime('%Y%m%dT%H%M%SZ')}.md"


def write_evidence(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8", newline="\n")
    print(f"Evidence written to {path}")


def check_prerequisites(call: Call, repo: str, metadata: dict[str, Json],
                        config: dict[str, Json]) -> None:
    root = f"repos/{repo}"
    branch = call("GET", root + "/branches/" + quote(metadata["default_branch"], safe=""))
    checks = pages(call, root + "/commits/" + branch["commit"]["sha"] + "/check-runs?filter=latest",
                   "check_runs")
    required = [check for r in config["rulesets"] for rule in r["rules"]
                if rule["type"] == "required_status_checks"
                for check in rule["parameters"]["required_status_checks"]]
    for needed in required:
        if not any(c["name"] == needed["context"]
                   and c.get("app", {}).get("id") == needed["integration_id"]
                   and c.get("conclusion") == "success" and c.get("status") == "completed"
                   for c in checks):
            raise SetupError(f"The required check {needed['context']!r} has not succeeded on the "
                             "default branch's head; run it before applying protection.")
    if not metadata.get("allow_merge_commit"):
        raise SetupError("Merge commits must be enabled before applying these rules.")
    if not metadata.get("allow_squash_merge"):
        raise SetupError("Squash merging must be enabled for the release-branch rules.")


def print_plan(plan: Plan) -> None:
    for method, endpoint, payload in plan:
        print(f"{method} {endpoint}")
        if payload is not None:
            print(json.dumps(payload, indent=2))


def run(args: argparse.Namespace, call: Call = api, now: datetime | None = None) -> int:
    if not re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9-]*/[A-Za-z0-9_.-]+", args.repo) \
            or args.repo.split("/")[1] in (".", ".."):
        raise SetupError("The repository must be an explicit owner/name on github.com.")
    config = load_config(args.config)
    metadata, plan = inspect(call, args.repo, config)
    mode = "apply" if args.apply else "check" if args.check else "preview"
    stamp = now or datetime.now(timezone.utc)
    path = evidence_path(args.report, mode, stamp)
    if not args.apply or not plan:
        write_evidence(path, evidence(call, args.repo, config, metadata, plan, mode, stamp))
    if not plan:
        print("VERIFIED: every setting in the baseline matches.")
        return 0
    print_plan(plan)
    if args.check:
        print("DRIFT: the saved settings differ from the baseline.")
        return 1
    if not args.apply:
        print("PREVIEW ONLY: nothing was changed. Use --apply to make these changes.")
        return 0
    if any(method == "PUT" and "/rulesets/" in endpoint for method, endpoint, _ in plan) \
            and not args.update_existing:
        raise SetupError("A named ruleset differs. Review the preview, then add --update-existing "
                         "to replace that whole ruleset; any rule not in the baseline is removed.")
    check_prerequisites(call, args.repo, metadata, config)
    for method, endpoint, payload in plan:
        call(method, endpoint, payload)
        print(f"Applied {method} {endpoint}")
    metadata, remaining = inspect(call, args.repo, config, report=False)
    write_evidence(path, evidence(call, args.repo, config, metadata, remaining, mode, stamp))
    if remaining:
        raise SetupError("Read-back failed; run the preview to inspect the remaining differences.")
    print("VERIFIED: every setting in the baseline matches after apply.")
    return 0


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=(__doc__ or "").splitlines()[0])
    parser.add_argument("--repo", required=True, help="explicit owner/name")
    parser.add_argument("--config", type=Path, default=CONFIG)
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument("--apply", action="store_true", help="make the changes, then read them back")
    mode.add_argument("--check", action="store_true", help="verify read-only; exit 1 on drift")
    parser.add_argument("--update-existing", action="store_true",
                        help="with --apply, replace a named ruleset that differs")
    parser.add_argument("--report", type=Path,
                        help="evidence file (default: test-results/repository-settings-<mode>-<time>.md)")
    args = parser.parse_args(argv)
    if args.update_existing and not args.apply:
        parser.error("--update-existing requires --apply")
    return args


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    try:
        return run(args)
    except (SetupError, OSError, ValueError, KeyError, TypeError) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main())
