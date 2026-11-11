#!/usr/bin/env python3
"""Shared rules of the Excel evidence bundle written by tools/Certify-IRRBB.ps1.

A bundle is a folder outside the checkout holding the logs the script wrote,
session.json (what the script observed of the host) and host.json (the record
tools/write_excel_evidence.py writes). Every log is UTF-8 text of KEY=VALUE
lines, the first one CANDIDATE=<full SHA>. This module reads the expectations
from the repository at the candidate commit, never from the working tree, and
judges each stage of docs/EXCEL_EVIDENCE.md from the logs alone. It never runs
Excel: what it checks is that the record is complete, consistent and bound to
the candidate, not what a person did.
"""
from __future__ import annotations

import functools
import hashlib
import json
import re
import subprocess
import sys
import types
from collections.abc import Callable
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

REPOSITORY = "danielep71/K-IRRBB-NMDs"
SCHEMA_VERSION = 1
SHA = re.compile(r"^[0-9a-f]{40}$")
LOGS = ("template.log", "build.log", "harness.log", "failure.log", "import.log",
        "sync.log", "sync-failure.log", "roundtrip.log")
STAGES = ("template", "build", "references", "compile", "smoke", "failure_path", "import",
          "cleanup", "synchronize", "synchronize_failure", "round_trip")
TEMPLATE = "src/workbook/IRRBB_Template.xlsx"
FIXTURES = "tests/fixtures/data_contract"
EXPECTED = "tests/expected/data_contract"
REFERENCES = (re.compile(r"^Visual Basic [Ff]or Applications$"),
              re.compile(r"^Microsoft Excel \d+\.\d+ Object Library$"),
              re.compile(r"^OLE Automation$"),
              re.compile(r"^Microsoft Office \d+\.\d+ Object Library$"))
COMPLETE = "COMPLETE - ALL EXECUTED CASES PASSED"
ACCEPTED = "COMPLETE - IMPORT ACCEPTED"
FAILURE_RUNS = ("TEST_Harness.TEST_RunTestsWithInjectedFailures", "TEST_Harness.RunTests",
                "TEST_Harness.TEST_RunTestsWithWriteFault", "TEST_Harness.RunTests")
INJECTED = {"HARN-OUT-01": "FAIL", "DATA-CODES-01": "ERROR"}
WRITE_FAULT = "-2147220503"
SCENARIOS = ("accepted", "control-match", "control-differ", "pasted-window", "pasted-attestation", "rerun")
FAILED_BECAUSE = {"pasted-window": "data window", "pasted-attestation": "not attested"}
# Checks_Info rows B4:B11 and DataSummary_Info rows B4:B12, 0-based.
CHECKS_COMMIT, CHECKS_COUNTS, CHECKS_ASSERTIONS, CHECKS_PRIMARY, CHECKS_CLEANUP = 2, 4, 5, 6, 7
IMPORT_PRIMARY, IMPORT_CLEANUP = 7, 8


class EvidenceError(Exception):
    """A bundle or candidate that cannot be read at all."""


# ----------------------------------------------------------------------------
# The candidate commit
# ----------------------------------------------------------------------------
def blob(root: Path, sha: str, path: str) -> bytes:
    """The bytes of a file as stored in Git at the candidate commit (LF for text)."""
    completed = subprocess.run(["git", "-C", str(root), "cat-file", "blob", f"{sha}:{path}"],
                               stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=False)
    if completed.returncode:
        raise EvidenceError(f"{path} is not in commit {sha}")
    return completed.stdout


def sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


@dataclass
class Expected:
    """What a passing bundle shows for one candidate, read from that commit."""
    registry: list[tuple[str, bool]]
    sheets: list[tuple[str, str]]
    names: dict[str, str]
    validated: list[str]
    totals: int
    differ_errors: list[tuple[int, str]]
    control_sha: str
    sources: dict[str, str] = field(default_factory=dict)

    @property
    def counts_line(self) -> str:
        built = sum(1 for _, b in self.registry if b)
        return (f"{built} PASS, 0 FAIL, 0 ERROR, {len(self.registry) - built} NOT RUN "
                f"(of {len(self.registry)} registered)")


def _template_module(root: Path, sha: str) -> types.ModuleType:
    """tools/workbook/build_template.py as committed at the candidate, for its SHEETS, NAMES
    and VALIDATIONS; its dataclasses need the module registered while it loads."""
    name = f"_candidate_build_template_{sha}"
    module = types.ModuleType(name)
    module.__dict__["__file__"] = str(root / "tools" / "workbook" / "build_template.py")
    sys.modules[name] = module
    try:
        exec(compile(blob(root, sha, "tools/workbook/build_template.py"), "build_template.py", "exec"),
             module.__dict__)
    finally:
        del sys.modules[name]
    return module


@functools.lru_cache(maxsize=8)
def expected(root: Path, sha: str) -> Expected:
    harness = blob(root, sha, "tests/modules/TEST_Harness.bas").decode("cp1252")
    registry = [(case, built == "True") for case, built in
                re.findall(r'^\s*c\.Add Array\("([A-Z0-9-]+)", (True|False),', harness, re.MULTILINE)]
    if not registry:
        raise EvidenceError("the harness registry could not be read")
    template = _template_module(root, sha)
    sheets = [(s.name, s.code_name) for s in template.SHEETS]
    names = {name: f"={sheet}!{ref}" for name, (sheet, ref) in template.NAMES.items()}
    validated = [f"{sheet}!{cells.split(':')[0]}" for sheet, cells, *_ in template.VALIDATIONS]
    summary = json.loads(blob(root, sha, f"{EXPECTED}/accounts_summary.json"))
    findings = json.loads(blob(root, sha, f"{EXPECTED}/control_totals_findings.json"))
    differ = next(c for c in findings["cases"] if c["fixture"].endswith("control_totals_balance.csv"))
    projects = json.loads(blob(root, sha, "tools/vba-projects.json"))
    paths = list(projects["document_modules"]) + list(projects["configurations"]["development"]["modules"])
    sources = {path: sha256(blob(root, sha, path)) for path in sorted(paths + [TEMPLATE])}
    return Expected(registry=registry, sheets=sheets, names=names, validated=validated,
                    totals=len(summary["totals"]),
                    differ_errors=[(e["line"], e["code"]) for e in differ["errors"]],
                    control_sha=sha256(blob(root, sha, f"{FIXTURES}/control_totals.csv")),
                    sources=sources)


# ----------------------------------------------------------------------------
# Logs
# ----------------------------------------------------------------------------
Lines = list[tuple[str, str]]


def read_log(bundle: Path, name: str, sha: str) -> Lines:
    """The KEY=VALUE lines of a log after its CANDIDATE= line, which must name sha."""
    path = bundle / name
    if not path.is_file():
        raise EvidenceError(f"{name} is missing")
    lines: Lines = []
    for raw in path.read_text(encoding="utf-8-sig").splitlines():
        if raw.strip():
            key, _, value = raw.partition("=")
            lines.append((key.strip(), value))
    if not lines or lines[0] != ("CANDIDATE", sha):
        raise EvidenceError(f"{name} does not start with CANDIDATE={sha}")
    return lines[1:]


def blocks(lines: Lines, key: str) -> list[Lines]:
    """The lines split at each `key` line; each block starts with that line."""
    out: list[Lines] = []
    for item in lines:
        if item[0] == key:
            out.append([item])
        elif out:
            out[-1].append(item)
    return out


def pairs(items: list[str]) -> dict[str, str]:
    """name|value entries as a dict."""
    out: dict[str, str] = {}
    for item in items:
        name, sep, value = item.partition("|")
        if sep:
            out[name] = value
    return out


def values(lines: Lines, key: str) -> list[str]:
    return [value for k, value in lines if k == key]


def one(lines: Lines, key: str) -> str | None:
    found = values(lines, key)
    return found[0] if len(found) == 1 else None


# ----------------------------------------------------------------------------
# Stage judgements: each returns the reasons a stage fails, empty for a pass
# ----------------------------------------------------------------------------
def _state(block: Lines) -> list[str]:
    before, after = one(block, "STATE_BEFORE"), one(block, "STATE_AFTER")
    if before is None or after is None:
        return ["Excel state before and after the run is not recorded"]
    return [] if before == after else [f"Excel state changed: {before} -> {after}"]


def _cases(block: Lines, exp: Expected, injected: dict[str, str] | None = None) -> list[str]:
    rows = [value.split("|") for value in values(block, "CASE")]
    if [r[0] for r in rows] != [case for case, _ in exp.registry]:
        return ["the case table does not list the registry in order"]
    reasons = []
    for row, (case, built) in zip(rows, exp.registry, strict=True):
        if len(row) < 5:
            reasons.append(f"{case}: incomplete row")
            continue
        outcome, asserts, passed, failed = row[1], row[2], row[3], row[4]
        want = (injected or {}).get(case, "PASS" if built else "NOT RUN")
        if outcome != want:
            reasons.append(f"{case}: {outcome}, expected {want}")
        elif want == "PASS" and not (asserts.isdigit() and int(asserts) > 0 and asserts == passed
                                     and failed == "0"):
            reasons.append(f"{case}: PASS with assertions {asserts}, passed {passed}, failed {failed}")
    return reasons


def _clean_run(block: Lines, exp: Expected, sha: str) -> list[str]:
    reasons = _state(block) + _cases(block, exp)
    info = values(block, "INFO")
    if one(block, "STATUS") != COMPLETE:
        reasons.append(f"status {one(block, 'STATUS')!r}, expected {COMPLETE!r}")
    if len(info) != 8:
        return reasons + ["Checks_Info is not recorded in full"]
    if info[CHECKS_COMMIT] != sha:
        reasons.append(f"the workbook names source commit {info[CHECKS_COMMIT]!r}")
    if info[CHECKS_COUNTS] != exp.counts_line:
        reasons.append(f"counts {info[CHECKS_COUNTS]!r}, expected {exp.counts_line!r}")
    total = sum(int(r.split("|")[2]) for r in values(block, "CASE") if r.split("|")[2].isdigit())
    if info[CHECKS_ASSERTIONS] != f"{total} of {total} assertions passed" or total == 0:
        reasons.append(f"assertions {info[CHECKS_ASSERTIONS]!r} for {total} recorded")
    if info[CHECKS_PRIMARY] or info[CHECKS_CLEANUP]:
        reasons.append("an error is reported in Checks!B10:B11")
    return reasons + [f"run error: {e}" for e in values(block, "ERROR")]


def judge_template(lines: Lines, exp: Expected) -> list[str]:
    reasons = []
    if one(lines, "OPENED") != "True":
        return ["the template did not open: " + "; ".join(values(lines, "ERROR"))]
    caption = one(lines, "CAPTION") or ""
    if "repair" in caption.lower():
        reasons.append(f"Excel repaired the template: {caption}")
    if [tuple(v.split("|", 1)) for v in values(lines, "SHEET")] != exp.sheets:
        reasons.append("sheets or CodeNames differ from the workbook contract")
    names = pairs(values(lines, "NAME"))
    if names != exp.names:
        differing = sorted(set(names) ^ set(exp.names) | {n for n in names if names[n] != exp.names.get(n)})
        reasons.append("defined names differ from the workbook contract: " + ", ".join(differing[:5]))
    probes = [v.split("|") for v in values(lines, "VALIDATION")]
    for cell in exp.validated:
        outcomes = {p[2] for p in probes if len(p) == 3 and p[0] == cell}
        if outcomes != {"refused", "accepted"}:
            reasons.append(f"validation of {cell}: a refused and an accepted value are both needed")
    fonts = {v.split("|", 1)[1] for v in values(lines, "FONT") if "|" in v}
    if fonts != {"Aptos"}:
        reasons.append(f"title fonts {sorted(fonts)}, expected Aptos")
    return reasons


def judge_exit(lines: Lines, what: str) -> list[str]:
    code = one(lines, "EXIT")
    return [] if code == "0" else [f"{what} exited with code {code}"]


def judge_failure_path(lines: Lines, exp: Expected, sha: str) -> list[str]:
    runs = blocks(lines, "RUN")
    if [one(b, "RUN") for b in runs] != list(FAILURE_RUNS):
        return ["the four runs of stage 4 are not recorded in order"]
    injected, rerun, fault, final = runs
    reasons = _state(injected) + _cases(injected, exp, INJECTED)
    if one(injected, "STATUS") != "COMPLETE - FAILURES":
        reasons.append(f"injected run: status {one(injected, 'STATUS')!r}")
    reasons += [f"rerun: {r}" for r in _clean_run(rerun, exp, sha)]
    info = values(fault, "INFO")
    reasons += [f"write fault: {r}" for r in _state(fault)]
    if one(fault, "STATUS") != "FAILED" or len(info) != 8:
        reasons.append(f"write fault: status {one(fault, 'STATUS')!r}")
    else:
        if not info[CHECKS_PRIMARY].startswith(f"Primary error {WRITE_FAULT}"):
            reasons.append("write fault: the primary error is not kept")
        if not info[CHECKS_CLEANUP].startswith(f"Cleanup error {WRITE_FAULT}"):
            reasons.append("write fault: the cleanup error is not kept")
    if values(fault, "CASE") or one(fault, "ASSERTION_ROWS") != "0":
        reasons.append("write fault: the result tables are not empty")
    return reasons + [f"final run: {r}" for r in _clean_run(final, exp, sha)]


def _accepted(block: Lines, exp: Expected) -> list[str]:
    reasons = []
    if one(block, "STATUS") != ACCEPTED:
        reasons.append(f"status {one(block, 'STATUS')!r}")
    if one(block, "TOTAL_ROWS") != str(exp.totals):
        reasons.append(f"{one(block, 'TOTAL_ROWS')} total rows, expected {exp.totals}")
    if values(block, "FINDING"):
        reasons.append("findings listed for an accepted import")
    return reasons


def _withdrawn(block: Lines) -> list[str]:
    rows = (one(block, "TOTAL_ROWS"), one(block, "OUTCOME_ROWS"))
    return [] if rows == ("0", "0") else ["earlier totals or outcomes are still shown"]


def judge_import(lines: Lines, exp: Expected) -> list[str]:
    runs = blocks(lines, "SCENARIO")
    if [one(b, "SCENARIO") for b in runs] != list(SCENARIOS):
        return ["the import scenarios are not recorded in order"]
    reasons: list[str] = []
    for block in runs:
        name = one(block, "SCENARIO") or ""
        reasons += [f"{name}: {r}" for r in _state(block) + _scenario(name, block, exp)]
    return reasons


def _scenario(name: str, block: Lines, exp: Expected) -> list[str]:
    lineage = pairs(values(block, "LINEAGE"))
    if name in ("accepted", "rerun"):
        found = _accepted(block, exp)
        if lineage.get("Control-totals file") != "none":
            found.append("the lineage names a control-totals file")
        return found
    if name == "control-match":
        found = _accepted(block, exp)
        if lineage.get("Control-totals file") != "control_totals.csv" \
                or lineage.get("Control-totals SHA-256") != exp.control_sha:
            found.append("the lineage does not record control_totals.csv and its SHA-256")
        return found
    if name == "control-differ":
        found = []
        want = f"FAILED - IMPORT REJECTED ({len(exp.differ_errors)} errors)"
        if one(block, "STATUS") != want:
            found.append(f"status {one(block, 'STATUS')!r}, expected {want!r}")
        rows = [f.split("|") for f in values(block, "FINDING")]
        got = sorted((int(r[1]), r[2]) for r in rows if len(r) >= 3 and r[1].isdigit())
        if got != sorted(exp.differ_errors):
            found.append(f"findings {got}, expected {sorted(exp.differ_errors)}")
        return found + _withdrawn(block)
    info = values(block, "INFO")
    primary = info[IMPORT_PRIMARY] if len(info) == 9 else ""
    found = []
    if one(block, "STATUS") != "FAILED" or FAILED_BECAUSE[name] not in primary.lower():
        found.append(f"status {one(block, 'STATUS')!r} with primary error {primary!r}")
    return found + _withdrawn(block)


def judge_synchronize(lines: Lines, exp: Expected, sha: str) -> list[str]:
    reasons = judge_exit(lines, "Sync-IRRBB-VBA.ps1")
    if not any("SUCCESS" in v for v in values(lines, "OUTPUT")):
        reasons.append("the synchronizer did not report SUCCESS")
    if one(lines, "COMPILE") != "PASS":
        reasons.append("the synchronized workbook did not compile")
    runs = blocks(lines, "RUN")
    if len(runs) != 1:
        return reasons + ["the harness run on the synchronized workbook is not recorded"]
    return reasons + [f"harness: {r}" for r in _clean_run(runs[0], exp, sha)]


def judge_synchronize_failure(lines: Lines) -> list[str]:
    reasons = []
    if one(lines, "EXIT") in (None, "0"):
        reasons.append("the synchronizer did not fail on a workbook opened elsewhere")
    if not any("read-only" in v.lower() for v in values(lines, "OUTPUT")):
        reasons.append("the synchronizer did not report the read-only workbook")
    before, after = one(lines, "HASH_BEFORE"), one(lines, "HASH_AFTER")
    if not before or before != after:
        reasons.append("the workbook changed")
    for setting in ("MACRO_POLICY", "VBA_PROJECT_ACCESS"):
        if one(lines, f"{setting}_BEFORE") != one(lines, f"{setting}_AFTER"):
            reasons.append(f"{setting.lower().replace('_', ' ')} changed")
    return reasons


def judge_round_trip(lines: Lines) -> list[str]:
    exits = values(lines, "EXIT")
    reasons = [] if exits == ["0", "0"] else [f"the export and the comparison exited with codes {exits}"]
    results = [v for v in values(lines, "OUTPUT") if v.startswith("RESULT=")]
    if not results or not results[-1].startswith("RESULT=PASS") or "binding=VERIFIED" not in results[-1]:
        reasons.append(f"comparison {results[-1] if results else 'missing'}, expected RESULT=PASS "
                       "with binding=VERIFIED")
    return reasons


def judge_session(session: dict[str, Any]) -> dict[str, list[str]]:
    refs = list(session.get("references") or [])
    ref_reasons = [] if len(refs) == len(REFERENCES) and all(p.match(r) for p, r in zip(REFERENCES, refs)) \
        else [f"references {refs}, expected the four defaults"]
    compile_status = (session.get("compile") or {}).get("status")
    return {
        "references": ref_reasons,
        "compile": [] if compile_status == "PASS" else [f"compile {compile_status}"],
        "cleanup": [] if session.get("workbook_closed") is True else ["the test workbook was not closed"],
    }


# ----------------------------------------------------------------------------
# The whole bundle
# ----------------------------------------------------------------------------
def evaluate(root: Path, sha: str, bundle: Path, session: dict[str, Any]) -> dict[str, dict[str, Any]]:
    """Each stage's outcome and reasons, judged from the logs and the session."""
    exp = expected(root, sha)

    def smoke(lines: Lines) -> list[str]:
        runs = blocks(lines, "RUN")
        return _clean_run(runs[0], exp, sha) if len(runs) == 1 else ["no single harness run"]

    judges: dict[str, tuple[str, Callable[[Lines], list[str]]]] = {
        "template": ("template.log", lambda ls: judge_template(ls, exp)),
        "build": ("build.log", lambda ls: judge_exit(ls, "build_dev_workbook.py")),
        "smoke": ("harness.log", smoke),
        "failure_path": ("failure.log", lambda ls: judge_failure_path(ls, exp, sha)),
        "import": ("import.log", lambda ls: judge_import(ls, exp)),
        "synchronize": ("sync.log", lambda ls: judge_synchronize(ls, exp, sha)),
        "synchronize_failure": ("sync-failure.log", judge_synchronize_failure),
        "round_trip": ("roundtrip.log", judge_round_trip),
    }
    reasons: dict[str, list[str]] = judge_session(session)
    for stage, (log, judge) in judges.items():
        try:
            reasons[stage] = judge(read_log(bundle, log, sha))
        except EvidenceError as error:
            reasons[stage] = [str(error)]
    return {stage: {"outcome": "PASS" if not reasons[stage] else "FAIL", "reasons": reasons[stage]}
            for stage in STAGES}


def record(root: Path, sha: str, bundle: Path) -> dict[str, Any]:
    """The host.json record of a bundle: environment, digests, stage outcomes."""
    if not SHA.match(sha):
        raise EvidenceError("the candidate must be a full 40-character lowercase SHA")
    session_path = bundle / "session.json"
    if not session_path.is_file():
        raise EvidenceError("session.json is missing")
    session = json.loads(session_path.read_text(encoding="utf-8-sig"))
    stages = evaluate(root, sha, bundle, session)
    logs = {name: sha256((bundle / name).read_bytes())
            for name in LOGS + ("session.log", "session.json") if (bundle / name).is_file()}
    environment = {key: session.get(key, "unknown") for key in
                   ("excel_version", "excel_build", "office_bitness", "os", "os_architecture",
                    "references", "macro_policy", "vba_project_access")}
    environment["trust_changes"] = False
    passed = all(s["outcome"] == "PASS" for s in stages.values())
    return {
        "schema_version": SCHEMA_VERSION,
        "repository": REPOSITORY,
        "candidate": sha,
        "operator": session.get("operator", ""),
        "started_at": session.get("started_at", ""),
        "finished_at": session.get("finished_at", ""),
        "environment": environment,
        "sources": expected(root, sha).sources,
        "logs": logs,
        "stages": stages,
        "result": ("PASS" if certifying(session) else "BEST_EFFORT") if passed else "FAIL",
    }


def certifying(session: dict[str, Any]) -> bool:
    """Only 64-bit Excel certifies; a pass on any other host is best-effort evidence
    (docs/EXCEL_EVIDENCE.md#bitness)."""
    return session.get("office_bitness") == "64-bit"


def validate(root: Path, bundle: Path, sha: str | None = None) -> tuple[str, list[str]]:
    """Check host.json against a fresh judgement of the bundle; (result, findings)."""
    host_path = bundle / "host.json"
    if not host_path.is_file():
        return "FAIL", ["host.json is missing"]
    host = json.loads(host_path.read_text(encoding="utf-8"))
    candidate = host.get("candidate", "")
    if sha is not None and candidate != sha:
        return "FAIL", [f"host.json certifies {candidate}, not {sha}"]
    try:
        fresh = record(root, candidate, bundle)
    except EvidenceError as error:
        return "FAIL", [str(error)]
    # Every field is rebuilt from session.json, the logs and the candidate, so an edited
    # host, operator, time, digest or outcome cannot pass.
    findings = [f"host.json {key} does not match the bundle and the candidate"
                for key in sorted(set(fresh) | set(host)) if host.get(key) != fresh.get(key)]
    for stage, outcome in fresh["stages"].items():
        findings += [f"{stage}: {reason}" for reason in outcome["reasons"]]
    if fresh["result"] == "BEST_EFFORT":
        bitness = fresh["environment"]["office_bitness"]
        findings.append(f"Office is {bitness}: only 64-bit Excel certifies, so this pass is best-effort evidence")
        return ("BEST_EFFORT" if len(findings) == 1 else "FAIL"), findings
    result = "PASS" if not findings and fresh["result"] == "PASS" else "FAIL"
    return result, findings


def report(host: dict[str, Any]) -> list[str]:
    """One line per stage and a final RESULT line."""
    lines = []
    for stage in STAGES:
        outcome = host["stages"][stage]
        detail = "" if outcome["outcome"] == "PASS" else " - " + "; ".join(outcome["reasons"][:3])
        lines.append(f"{stage:20} {outcome['outcome']}{detail}")
    lines.append(f"RESULT={host['result']}; candidate={host['candidate']}")
    return lines
