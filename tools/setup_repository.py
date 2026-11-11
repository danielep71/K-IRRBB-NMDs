#!/usr/bin/env python3
"""Compare, apply or check the repository settings baseline through the GitHub CLI.

The baseline is docs/repository-settings.json. Without a mode the script previews
the difference with its plan digest and never writes; --apply, given that digest,
writes the same plan and reads it back; --check verifies read-only and exits 1
on drift. Every run writes a dated
evidence file. It needs an authenticated gh; the preview and --check record what
they cannot read as unknown, and --apply needs admin rights. It changes only the
rulesets and environments the baseline names; it never touches environment
secrets or variables, and it reads the Actions settings without changing them.
"""
from __future__ import annotations

import argparse
import copy
import hashlib
import json
import os
import re
import subprocess
import sys
from collections.abc import Callable
from dataclasses import dataclass, field
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
CONFIG_KEYS = {"schema_version", "repository", "security", "actions", "environments",
               "code_scanning_default_setup", "rulesets"}
# Each Actions setting: the endpoint under the repository and the response field that hold it.
# The script reads them and never changes them.
ACTIONS = {
    "sha_pinning_required": ("actions/permissions", "sha_pinning_required"),
    "default_workflow_permissions": ("actions/permissions/workflow", "default_workflow_permissions"),
    "can_approve_pull_request_reviews": ("actions/permissions/workflow",
                                         "can_approve_pull_request_reviews"),
    "fork_pr_approval_policy": ("actions/permissions/fork-pr-contributor-approval", "approval_policy"),
    "retention_days": ("actions/permissions/artifact-and-log-retention", "days"),
}
FORK_APPROVAL_POLICIES = ("first_time_contributors_new_to_github", "first_time_contributors",
                          "all_external_contributors")
# Only main may deploy: custom branch policies, not "protected branches".
CUSTOM_BRANCHES = {"protected_branches": False, "custom_branch_policies": True}

Json = Any
Call = Callable[..., Json]
Plan = list[tuple[str, str, Json]]


class SetupError(Exception):
    pass


class Unknown:
    """A managed value that could not be read. `==` is never true for it, itself included, so
    an unread value is never taken as matching the baseline."""

    def __init__(self, reason: str) -> None:
        self.reason = reason

    def __eq__(self, other: object) -> bool:
        return False

    __hash__ = object.__hash__

    def __str__(self) -> str:
        return f"unknown ({self.reason})"


@dataclass
class Reading:
    """The writes that would reach the baseline, the managed values that could not be read, and
    the observed Actions settings that differ from it, which the script never changes."""
    metadata: dict[str, Json]
    plan: Plan
    unknown: dict[str, Unknown] = field(default_factory=dict)
    manual: dict[str, Json] = field(default_factory=dict)

    def verified(self) -> bool:
        return not self.plan and not self.unknown and not self.manual


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


class OneReading:
    """A call that makes each GET once and replays its answer, a failure included, so the
    evidence records exactly the values the plan was computed from. Once frozen, a GET that was
    not made before is an error: the evidence never reads anything new."""

    def __init__(self, call: Call) -> None:
        self.call = call
        self.answers: dict[str, Json] = {}
        self.frozen = False

    def __call__(self, method: str, endpoint: str, payload: Json = None) -> Json:
        if method != "GET":
            raise SetupError(f"Internal error: {method} {endpoint} during a reading.")
        if endpoint not in self.answers:
            if self.frozen:
                raise SetupError(f"Internal error: the evidence read {endpoint}, which the "
                                 "inspection did not.")
            try:
                self.answers[endpoint] = self.call(method, endpoint)
            except SetupError as error:
                self.answers[endpoint] = error
        answer = self.answers[endpoint]
        if isinstance(answer, SetupError):
            raise answer
        return copy.deepcopy(answer)


def attempt(read: Callable[[], Json]) -> Json:
    """Run a read; a 403 or 404 makes its value unknown, naming the endpoint and status."""
    try:
        return read()
    except SetupError as error:
        found = re.match(r"(GET \S+) failed \(HTTP (403|404)\)", str(error))
        if found:
            return Unknown(f"{found.group(1)}: HTTP {found.group(2)}")
        raise


def is_admin(metadata: dict[str, Json]) -> bool:
    return bool((metadata.get("permissions") or {}).get("admin"))


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
    # Ignore response metadata, but keep unknown rules and parameters in the comparison. A field
    # the response omits is unknown: GitHub returns bypass_actors only to callers who may see it.
    return canonical({k: rule[k] if k in rule else Unknown(f"{k} not returned") for k in RULE_KEYS})


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


def check_actions(actions: Json) -> None:
    if not isinstance(actions, dict) or set(actions) != set(ACTIONS):
        raise SetupError("The configuration must hold exactly the five Actions settings.")
    valid = {
        "sha_pinning_required": type(actions["sha_pinning_required"]) is bool,
        "default_workflow_permissions": actions["default_workflow_permissions"] in ("read", "write"),
        "can_approve_pull_request_reviews": type(actions["can_approve_pull_request_reviews"]) is bool,
        "fork_pr_approval_policy": actions["fork_pr_approval_policy"] in FORK_APPROVAL_POLICIES,
        "retention_days": type(actions["retention_days"]) is int and 1 <= actions["retention_days"] <= 400,
    }
    for key, ok in valid.items():
        if not ok:
            raise SetupError(f"Invalid baseline Actions setting: {key}.")


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
    check_actions(value["actions"])
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


def ruleset_plan(call: Call, root: str, current: Json, config: dict[str, Json],
                 unknown: dict[str, Unknown]) -> Plan:
    """Create a missing named ruleset and replace one that differs; never touch another."""
    plan: Plan = []
    for desired in config["rulesets"]:
        label = f'ruleset "{desired["name"]}"'
        if isinstance(current, Unknown):
            unknown[label] = current
            continue
        matches = [r for r in current if r["name"] == desired["name"]]
        if len(matches) > 1:
            raise SetupError(f"Ambiguous duplicate ruleset name: {desired['name']}")
        if not matches:
            plan.append(("POST", root + "/rulesets", desired))
        else:
            endpoint = root + "/rulesets/" + str(matches[0]["id"])
            response = attempt(lambda: call("GET", endpoint))
            if isinstance(response, Unknown):
                unknown[label] = response
                continue
            live = rule_spec(response)
            omitted = [k for k in RULE_KEYS if isinstance(live[k], Unknown)]
            if omitted:
                unknown[label] = Unknown(f"{', '.join(omitted)} not returned by GET {endpoint}")
            elif live != rule_spec(desired):
                plan.append(("PUT", endpoint, desired))
    return plan


def inspect(call: Call, repo: str, config: dict[str, Json], report: bool = True) -> Reading:
    root = f"repos/{repo}"
    metadata = call("GET", root)
    if metadata.get("archived") or metadata.get("disabled"):
        raise SetupError("The repository is archived or disabled.")
    current = attempt(lambda: pages(call, root + "/rulesets?includes_parents=false"))
    inherited = attempt(lambda: pages(call, root + "/rulesets?includes_parents=true"))
    print(f"Repository: {metadata['full_name']}; visibility: {metadata['visibility']}; "
          f"default branch: {metadata['default_branch']}")
    if report and not isinstance(inherited, Unknown) \
            and any(r.get("source_type") != "Repository" for r in inherited):
        print("NOTE: inherited rulesets also apply; this script neither changes nor checks them.")
    if report and not isinstance(current, Unknown):
        report_unmanaged(call, root, current, config)
    unknown: dict[str, Unknown] = {}
    plan = ruleset_plan(call, root, current, config, unknown)
    plan += security_plan(call, root, metadata, config, report, unknown)
    for name, environment in config["environments"].items():
        plan += environment_plan(call, root, name, environment["deployment_branches"], unknown)
    desired_setup = config["code_scanning_default_setup"]
    observed_setup = default_setup(call, root, metadata)
    if isinstance(observed_setup, Unknown):
        unknown["`code_scanning_default_setup`"] = observed_setup
    elif observed_setup is None:
        if report:
            print("NOTE: code scanning is not available on this private repository; the CodeQL "
                  "default setup cannot be on and is not checked.")
    elif observed_setup != desired_setup:
        plan.append(("PATCH", root + "/code-scanning/default-setup", {"state": desired_setup}))
    manual: dict[str, Json] = {}
    for key, observed in actions_state(call, root).items():
        if isinstance(observed, Unknown):
            unknown[f"`{key}`"] = observed
        elif observed != config["actions"][key]:
            manual[key] = observed
    unknown.update({f"`{k}`": Unknown(f"{k} not returned by GET {root}")
                    for k in config["repository"] if k not in metadata})
    changes = {k: v for k, v in config["repository"].items() if k in metadata and metadata[k] != v}
    # Protection first, then the general settings.
    if changes:
        plan.append(("PATCH", root, changes))
    return Reading(metadata, plan, unknown, manual)


def actions_state(call: Call, root: str) -> dict[str, Json]:
    """The Actions settings as GitHub reports them, each Unknown when unreadable; read only."""
    responses: dict[str, Json] = {}
    state: dict[str, Json] = {}
    for key, (path, name) in ACTIONS.items():
        endpoint = f"{root}/{path}"
        if endpoint not in responses:
            responses[endpoint] = attempt(lambda: call("GET", endpoint))
        response = responses[endpoint]
        if isinstance(response, Unknown):
            state[key] = response
        elif not isinstance(response, dict) or name not in response:
            state[key] = Unknown(f"{name} not returned by GET {endpoint}")
        else:
            state[key] = response[name]
    return state


def missing(call: Call, endpoint: str, conclusive: bool = True) -> tuple[Json, bool]:
    """GET an endpoint that answers 404 when its feature is off; True when it does. A 404 means
    off only when conclusive: endpoints that answer admins alone also send it to anyone else.
    Otherwise a 403 or 404 leaves the value unknown."""
    value = attempt(lambda: call("GET", endpoint))
    if conclusive and isinstance(value, Unknown) and value.reason.endswith("HTTP 404"):
        return None, True
    return value, False


def security_fixes(call: Call, root: str, admin: bool) -> tuple[bool, bool] | Unknown:
    """Whether Dependabot security updates are enabled, and whether GitHub has paused them."""
    value, absent = missing(call, root + "/automated-security-fixes", admin)
    if absent:
        return False, False
    if isinstance(value, Unknown):
        return value
    value = value or {}
    return bool(value.get("enabled")), bool(value.get("paused"))


def security_state(call: Call, root: str, metadata: dict[str, Json], feature: str) -> bool | Unknown:
    """Whether a feature is in effect: paused security updates are enabled, not in effect."""
    admin = is_admin(metadata)
    if feature == "private_vulnerability_reporting":
        value = attempt(lambda: call("GET", root + "/private-vulnerability-reporting"))
        return value if isinstance(value, Unknown) else bool((value or {}).get("enabled"))
    if feature == "vulnerability_alerts":
        value, absent = missing(call, root + "/vulnerability-alerts", admin)
        return value if isinstance(value, Unknown) else not absent
    if feature == "automated_security_fixes":
        fixes = security_fixes(call, root, admin)
        return fixes if isinstance(fixes, Unknown) else fixes[0] and not fixes[1]
    analysis = metadata.get("security_and_analysis") or {}
    # GitHub returns security_and_analysis to admins only.
    if not admin and feature not in analysis:
        return Unknown(f"security_and_analysis.{feature} not returned by GET {root}")
    return bool((analysis.get(feature) or {}).get("status") == "enabled")


def matches(call: Call, root: str, metadata: dict[str, Json], feature: str,
            desired: bool) -> bool | Unknown:
    """Wanted on: the feature must be in effect. Wanted off: security updates must not even be
    enabled, since paused ones resume on their own."""
    if feature == "automated_security_fixes" and not desired:
        fixes = security_fixes(call, root, is_admin(metadata))
        return fixes if isinstance(fixes, Unknown) else not fixes[0]
    state = security_state(call, root, metadata, feature)
    return state if isinstance(state, Unknown) else state == desired


def security_plan(call: Call, root: str, metadata: dict[str, Json], config: dict[str, Json],
                  report: bool, unknown: dict[str, Unknown]) -> Plan:
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
        matched = matches(call, root, metadata, feature, desired)
        if isinstance(matched, Unknown):
            unknown[f"`{feature}`"] = matched
            continue
        if matched:
            continue
        if feature in ANALYSIS_KEYS:
            analysis[feature] = {"status": "enabled" if desired else "disabled"}
        else:
            plan.append(("PUT" if desired else "DELETE", root + "/" + feature.replace("_", "-"), None))
    if analysis:
        plan.append(("PATCH", root, {"security_and_analysis": analysis}))
    return plan


def environment_state(call: Call, root: str, name: str) -> tuple[Json, Json]:
    """The environment (None when missing) and its custom branch policies (None when the
    environment does not restrict deployments to custom branch policies); either is Unknown
    when unreadable. Anyone who can read the repository can read its environments."""
    endpoint = f"{root}/environments/{quote(name, safe='')}"
    environment, absent = missing(call, endpoint)
    if absent:
        return None, None
    if isinstance(environment, Unknown) or (environment.get("deployment_branch_policy") or {}
                                            ).get("custom_branch_policies") is not True:
        return environment, None
    return environment, attempt(lambda: pages(call, endpoint + "/deployment-branch-policies",
                                              "branch_policies"))


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


def environment_plan(call: Call, root: str, name: str, branches: list[str],
                     unknown: dict[str, Unknown]) -> Plan:
    """Restrict an environment's deployments to exactly the baseline branches."""
    endpoint = f"{root}/environments/{quote(name, safe='')}"
    environment, policies = environment_state(call, root, name)
    for value in (environment, policies):
        if isinstance(value, Unknown):
            unknown[f"`{name}` deployment branches"] = value
            return []
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


def describe_environment(call: Call, root: str, name: str) -> str | Unknown:
    environment, policies = environment_state(call, root, name)
    for value in (environment, policies):
        if isinstance(value, Unknown):
            return value
    if environment is None:
        return "missing"
    if policies is None:
        return "protected branches" if environment.get("deployment_branch_policy") else "all branches"
    return ", ".join(sorted(p["name"] if p.get("type", "branch") == "branch"
                            else f"{p['name']} ({p['type']})" for p in policies)) or "none"


def default_setup(call: Call, root: str, metadata: dict[str, Json]) -> str | None | Unknown:
    """The CodeQL default-setup state; None when code scanning is unavailable, which GitHub
    answers to an admin with 403 or 404 on a private repository without Code Security. On a
    public repository, or to anyone else, those statuses leave the state unknown."""
    value = attempt(lambda: call("GET", root + "/code-scanning/default-setup"))
    if not isinstance(value, Unknown):
        return str((value or {}).get("state"))
    if metadata.get("visibility") != "public" and is_admin(metadata):
        return None
    return value


def report_unmanaged(call: Call, root: str, current: list[Json], config: dict[str, Json]) -> None:
    """Name the repository rulesets outside the baseline; they stay active alongside it."""
    managed = {rule["name"] for rule in config["rulesets"]}
    for ruleset in current:
        if ruleset.get("source_type", "Repository") != "Repository" or ruleset["name"] in managed:
            continue
        detail = attempt(lambda: call("GET", f"{root}/rulesets/{ruleset['id']}"))
        if isinstance(detail, Unknown):
            detail = {}
        refs = ", ".join(((detail.get("conditions") or {}).get("ref_name") or {}).get("include") or [])
        print(f'NOTE: unmanaged ruleset "{ruleset["name"]}" '
              f'({detail.get("target", "unknown target")}: {refs or "no ref conditions"}; '
              f'{detail.get("enforcement", "unknown enforcement")}) stays active. This script '
              "never changes or deletes it; delete it in the repository settings if a managed "
              "ruleset replaces it.")


def cell(value: Json) -> str:
    return "on" if value is True else "off" if value is False else str(value)


def evidence(call: Call, repo: str, config: dict[str, Json], reading: Reading, mode: str,
             now: datetime) -> str:
    """A dated Markdown record of the managed settings as GitHub reports them. Given the
    OneReading that inspection used, it records the answers the plan was computed from."""
    root = f"repos/{repo}"

    metadata, plan = reading.metadata, reading.plan
    public = metadata.get("visibility") == "public"
    differs = {endpoint for method, endpoint, _ in plan if method == "PUT" and "/rulesets/" in endpoint}
    created = {payload["name"] for method, endpoint, payload in plan
               if method == "POST" and endpoint.endswith("/rulesets")}
    current = attempt(lambda: pages(call, root + "/rulesets?includes_parents=false"))
    inherited = attempt(lambda: pages(call, root + "/rulesets?includes_parents=true"))
    managed = {rule["name"] for rule in config["rulesets"]}
    head = ["# Repository settings evidence", "",
            f"- Checked: **{now.strftime('%Y-%m-%dT%H:%M:%SZ')}** with "
            f"`tools/setup_repository.py` ({mode})",
            f"- Repository: `{metadata['full_name']}`; visibility: {metadata['visibility']}; "
            f"default branch: `{metadata['default_branch']}`"]
    lines = ["", "## General settings", "", "| Setting | Baseline | Observed |", "| --- | --- | --- |"]
    lines += [f"| `{key}` | {cell(value)} | {cell(reading.unknown.get(f'`{key}`', metadata.get(key)))} |"
              for key, value in config["repository"].items()]
    lines += ["", "## Security features", "", "| Feature | Baseline | Observed |", "| --- | --- | --- |"]
    for feature, value in config["security"].items():
        if SECURITY_PUBLIC_ONLY[feature] and not public:
            observed = "not applied (public repositories only)"
        elif feature == "automated_security_fixes" \
                and security_fixes(call, root, is_admin(metadata)) == (True, True):
            observed = "on (paused)"
        else:
            observed = cell(security_state(call, root, metadata, feature))
        lines.append(f"| `{feature}` | {cell(value)} | {observed} |")
    lines += ["", "## Actions settings", "", "Read only: this script never changes them.", "",
              "| Setting | Baseline | Observed |", "| --- | --- | --- |"]
    observed_actions = actions_state(call, root)
    lines += [f"| `{key}` | {cell(value)} | {cell(observed_actions[key])} |"
              for key, value in config["actions"].items()]
    lines += ["", "## Environments and code scanning", "", "| Setting | Baseline | Observed |",
              "| --- | --- | --- |"]
    lines += [f"| `{name}` deployment branches | {', '.join(environment['deployment_branches'])} | "
              f"{cell(describe_environment(call, root, name))} |"
              for name, environment in config["environments"].items()]
    setup_state = default_setup(call, root, metadata)
    lines.append(f"| `code_scanning_default_setup` | {config['code_scanning_default_setup']} | "
                 f"{'not available (code scanning unavailable)' if setup_state is None else cell(setup_state)} |")
    lines += ["", "## Managed rulesets", "", "| Ruleset | ID | Status |", "| --- | --- | --- |"]
    for rule in config["rulesets"]:
        match = None if isinstance(current, Unknown) else \
            next((r for r in current if r["name"] == rule["name"]), None)
        unread = reading.unknown.get(f'ruleset "{rule["name"]}"')
        if unread is not None:
            lines.append(f"| {rule['name']} | {match['id'] if match else '—'} | {unread} |")
        elif rule["name"] in created or match is None:
            lines.append(f"| {rule['name']} | — | missing |")
        else:
            status = "differs" if f"{root}/rulesets/{match['id']}" in differs else "matches"
            lines.append(f"| {rule['name']} | {match['id']} | {status} |")
    unmanaged = current if isinstance(current, Unknown) else \
        ", ".join(r["name"] for r in current if r.get("source_type", "Repository") == "Repository"
                  and r["name"] not in managed) or "none"
    others = inherited if isinstance(inherited, Unknown) else \
        ", ".join(r["name"] for r in inherited if r.get("source_type") != "Repository") or "none"
    lines += ["", "## Other rules", "",
              f"- Unmanaged repository rulesets: {unmanaged}",
              f"- Inherited rulesets: {others}",
              "", "Settings outside the baseline are not recorded here."]
    head.append(f"- Result: **{'VERIFIED' if reading.verified() else 'DRIFT'}**")
    return "\n".join(head + lines) + "\n"


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


def print_manual(reading: Reading, config: dict[str, Json]) -> None:
    for key, observed in reading.manual.items():
        print(f"MANUAL: Actions setting `{key}` is {cell(observed)}; the baseline wants "
              f"{cell(config['actions'][key])}. Change it in the repository's Actions settings.")


def plan_digest(plan: Plan) -> str:
    """SHA-256 of the ordered plan: each operation's method, endpoint and payload, as canonical
    JSON (sorted keys, no spaces, UTF-8)."""
    text = json.dumps([{"method": method, "endpoint": endpoint, "payload": payload}
                       for method, endpoint, payload in plan],
                      sort_keys=True, separators=(",", ":"), ensure_ascii=False)
    return hashlib.sha256(text.encode("utf-8")).hexdigest()


def unread(unknown: dict[str, Unknown]) -> str:
    return "; ".join(f"{label}: {value.reason}" for label, value in unknown.items())


def authorize(args: argparse.Namespace, reading: Reading) -> None:
    """Refuse --apply, before any write, without admin rights, on an incomplete reading, or
    for a plan other than the one approved."""
    if not is_admin(reading.metadata):
        raise SetupError("Admin rights on the repository are required for --apply; nothing was "
                         "changed. The preview and --check run without them.")
    if reading.unknown:
        # Never change or certify a setting on a reading that misses part of it.
        raise SetupError(f"Nothing was changed: these managed values could not be read: "
                         f"{unread(reading.unknown)}.")
    if reading.plan and args.approve_plan is None:
        raise SetupError("Nothing was changed: --apply needs --approve-plan with the plan digest "
                         "that the preview printed after the plan you reviewed.")
    if reading.plan and args.approve_plan != plan_digest(reading.plan):
        # The live settings changed since the preview: never write a plan no one reviewed.
        raise SetupError("Nothing was changed: the plan differs from the one approved. Run the "
                         "preview again and review the new plan.")


def run(args: argparse.Namespace, call: Call = api, now: datetime | None = None) -> int:
    if not re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9-]*/[A-Za-z0-9_.-]+", args.repo) \
            or args.repo.split("/")[1] in (".", ".."):
        raise SetupError("The repository must be an explicit owner/name on github.com.")
    config = load_config(args.config)
    # One reading feeds the plan, the result and the evidence; writes go through call itself.
    reader = OneReading(call)
    reading = inspect(reader, args.repo, config)
    metadata, plan = reading.metadata, reading.plan
    mode = "apply" if args.apply else "check" if args.check else "preview"
    stamp = now or datetime.now(timezone.utc)
    path = evidence_path(args.report, mode, stamp)
    if args.apply:
        authorize(args, reading)
    if not args.apply or reading.verified():
        reader.frozen = True
        write_evidence(path, evidence(reader, args.repo, config, reading, mode, stamp))
    if reading.verified():
        print("VERIFIED: every setting in the baseline matches.")
        return 0
    print_plan(plan)
    if plan:
        print(f"Plan digest: {plan_digest(plan)}")
    for label, value in reading.unknown.items():
        print(f"UNKNOWN: {label}: {value.reason}")
    print_manual(reading, config)
    if args.check:
        print("DRIFT: the saved settings differ from the baseline or could not be read.")
        return 1
    if not args.apply:
        print("PREVIEW ONLY: nothing was changed. To make these changes, review them and run "
              "--apply --approve-plan <plan digest>.")
        return 0
    if any(method == "PUT" and "/rulesets/" in endpoint for method, endpoint, _ in plan) \
            and not args.update_existing:
        raise SetupError("A named ruleset differs. Review the preview, then add --update-existing "
                         "to replace that whole ruleset; any rule not in the baseline is removed.")
    if plan:
        check_prerequisites(call, args.repo, metadata, config)
    for method, endpoint, payload in plan:
        call(method, endpoint, payload)
        print(f"Applied {method} {endpoint}")
    reader = OneReading(call)
    remaining = inspect(reader, args.repo, config, report=False)
    reader.frozen = True
    write_evidence(path, evidence(reader, args.repo, config, remaining, mode, stamp))
    if remaining.plan or remaining.unknown:
        raise SetupError("Read-back failed; run the preview to inspect the remaining differences.")
    if remaining.manual:
        print_manual(remaining, config)
        print("DRIFT: the Actions settings above differ from the baseline; this script never "
              "changes them.")
        return 1
    print("VERIFIED: every setting in the baseline matches after apply.")
    return 0


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=(__doc__ or "").splitlines()[0])
    parser.add_argument("--repo", required=True, help="explicit owner/name")
    parser.add_argument("--config", type=Path, default=CONFIG)
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument("--apply", action="store_true", help="make the changes, then read them back")
    mode.add_argument("--check", action="store_true", help="verify read-only; exit 1 on drift")
    parser.add_argument("--approve-plan", metavar="DIGEST",
                        help="with --apply, the plan digest the preview printed; a different plan "
                             "is refused before any write")
    parser.add_argument("--update-existing", action="store_true",
                        help="with --apply, replace a named ruleset that differs")
    parser.add_argument("--report", type=Path,
                        help="evidence file (default: test-results/repository-settings-<mode>-<time>.md)")
    args = parser.parse_args(argv)
    if args.update_existing and not args.apply:
        parser.error("--update-existing requires --apply")
    if args.approve_plan is not None:
        if not args.apply:
            parser.error("--approve-plan requires --apply")
        if not re.fullmatch(r"[0-9a-fA-F]{64}", args.approve_plan):
            parser.error("--approve-plan takes the 64-character SHA-256 plan digest")
        args.approve_plan = args.approve_plan.lower()
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
