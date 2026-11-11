#!/usr/bin/env python3
"""Tests for the Excel evidence bundle: tools/excel_evidence.py, its writer and
validator, and the static shape of tools/Certify-IRRBB.ps1.

A synthetic bundle is built from the expectations of the current commit, as a
passing run of the script would leave it; each test then breaks one thing and
expects the validator to report it. The PowerShell script is parsed where
PowerShell 7 (pwsh) is installed, as on the CI runner; nothing here runs Excel.
"""
from __future__ import annotations

import io
import json
import re
import shutil
import subprocess
import tempfile
import unittest
from collections.abc import Callable
from contextlib import redirect_stderr, redirect_stdout
from pathlib import Path
from typing import Any

import check_excel_evidence
import excel_evidence as ev
import write_excel_evidence

ROOT = Path(__file__).resolve().parents[1]
SCRIPT = ROOT / "tools" / "Certify-IRRBB.ps1"
LAUNCHER = ROOT / "tools" / "Certify-IRRBB.cmd"
PWSH = shutil.which("pwsh")
STATE = "-4105|True|True|True"
Lines = list[tuple[str, str]]


def head_sha() -> str:
    return subprocess.run(["git", "-C", str(ROOT), "rev-parse", "HEAD"], check=True,
                          capture_output=True, text=True).stdout.strip()


def harness_run(exp: ev.Expected, sha: str, entry: str, injected: dict[str, str] | None = None) -> Lines:
    lines: Lines = [("RUN", entry), ("STATE_BEFORE", STATE), ("STATE_AFTER", STATE)]
    cases, total = [], 0
    for case, built in exp.registry:
        outcome = (injected or {}).get(case, "PASS" if built else "NOT RUN")
        asserts, passed, failed = {"PASS": (3, 3, 0), "FAIL": (4, 3, 1)}.get(outcome, (0, 0, 0))
        total += asserts
        cases.append(("CASE", f"{case}|{outcome}|{asserts}|{passed}|{failed}|"))
    status = "COMPLETE - FAILURES" if injected else ev.COMPLETE
    info = ["46337.4", "46337.5", sha, "Microsoft Excel 16.0", exp.counts_line,
            f"{total} of {total} assertions passed", "", ""]
    lines += [("STATUS", status)] + [("INFO", value) for value in info] + cases
    return lines + [("ASSERTION_ROWS", str(total))]


def fault_run(sha: str) -> Lines:
    info = ["46337.4", "46337.5", sha, "Microsoft Excel 16.0", "", "",
            "Primary error -2147220503: Injected results write fault [TEST_Harness.WriteResults]",
            "Cleanup error -2147220503: Injected results write fault [TEST_Harness.WriteResults]"]
    return ([("RUN", "TEST_Harness.TEST_RunTestsWithWriteFault"), ("STATE_BEFORE", STATE),
             ("STATE_AFTER", STATE), ("STATUS", "FAILED")] + [("INFO", v) for v in info]
            + [("ASSERTION_ROWS", "0")])


def scenario(name: str, exp: ev.Expected) -> Lines:
    lines: Lines = [("SCENARIO", name), ("STATE_BEFORE", STATE), ("STATE_AFTER", STATE)]
    primary = {"pasted-window": "Primary error 5: The data window must be two month ends in order, "
                                "without a time of day. [IRRBB_Import.IRRBB_ImportData]",
               "pasted-attestation": "Primary error 5: Timely closure reporting is not attested; the "
                                     "import does not run. [IRRBB_Import.IRRBB_ImportData]"}.get(name, "")
    info = [("INFO", v) for v in ["46337.4", "46337.5", "accounts.csv", "market_rates.csv",
                                  "2025-01-31 to 2025-06-30", "EUR_1M", "Yes", primary, ""]]
    if name in ("accepted", "rerun", "control-match"):
        control = ("control_totals.csv", exp.control_sha) if name == "control-match" else ("none", "none")
        return lines + [("STATUS", ev.ACCEPTED)] + info + [
            ("TOTAL_ROWS", str(exp.totals)), ("OUTCOME_ROWS", "40"),
            ("LINEAGE", f"Control-totals file|{control[0]}"), ("LINEAGE", f"Control-totals SHA-256|{control[1]}")]
    if name == "control-differ":
        status = f"FAILED - IMPORT REJECTED ({len(exp.differ_errors)} errors)"
        findings = [("FINDING", f"control_totals_balance.csv|{line}|{code}|") for line, code in exp.differ_errors]
        return lines + [("STATUS", status)] + info + [("TOTAL_ROWS", "0"), ("OUTCOME_ROWS", "0")] + findings
    return lines + [("STATUS", "FAILED")] + info + [("TOTAL_ROWS", "0"), ("OUTCOME_ROWS", "0")]


def passing_logs(exp: ev.Expected, sha: str) -> dict[str, Lines]:
    template: Lines = [("FILE", ev.TEMPLATE), ("OPENED", "True"), ("CAPTION", "IRRBB_Template.xlsx  [Read-Only]")]
    template += [("SHEET", f"{name}|{code}") for name, code in exp.sheets]
    template += [("FONT", f"{name}|Aptos") for name, _ in exp.sheets]
    template += [("NAME", f"{name}|{ref}") for name, ref in exp.names.items()]
    for cell in exp.validated:
        template += [("VALIDATION", f"{cell}|x|refused"), ("VALIDATION", f"{cell}|Yes|accepted")]
    return {
        "template.log": template,
        "build.log": [("COMMAND", "python tools/devbuild/build_dev_workbook.py IRRBB_Dev.xlsm"), ("EXIT", "0")],
        "harness.log": harness_run(exp, sha, "TEST_Harness.RunTests"),
        "failure.log": (harness_run(exp, sha, ev.FAILURE_RUNS[0], ev.INJECTED)
                        + harness_run(exp, sha, ev.FAILURE_RUNS[1]) + fault_run(sha)
                        + harness_run(exp, sha, ev.FAILURE_RUNS[3])),
        "import.log": [line for name in ev.SCENARIOS for line in scenario(name, exp)],
        "sync.log": [("COMMAND", "Sync-IRRBB-VBA.ps1"),
                     ("OUTPUT", "SUCCESS: VBA synchronized. Workbook sheets were not removed or recreated."),
                     ("EXIT", "0"), ("COMPILE", "PASS")] + harness_run(exp, sha, "TEST_Harness.RunTests"),
        "sync-failure.log": [("HASH_BEFORE", "ab" * 32), ("MACRO_POLICY_BEFORE", "Disable VBA macros with notification"),
                             ("VBA_PROJECT_ACCESS_BEFORE", "On"), ("COMMAND", "Sync-IRRBB-VBA.ps1"),
                             ("OUTPUT", "Sync-IRRBB-VBA.ps1 : Workbook opened read-only. Close any other Excel "
                                        "instance using it."),
                             ("EXIT", "1"), ("HASH_AFTER", "ab" * 32),
                             ("MACRO_POLICY_AFTER", "Disable VBA macros with notification"),
                             ("VBA_PROJECT_ACCESS_AFTER", "On")],
        "roundtrip.log": [("COMMAND", "Export-IRRBB-VBA.ps1"), ("EXIT", "0"), ("COMMAND", "compare_vba_export.py"),
                          ("OUTPUT", "RESULT=PASS; components=31; differences=0; binding=VERIFIED"), ("EXIT", "0")],
    }


SESSION = {
    "operator": "Repository owner, local Windows workstation", "started_at": "2026-11-12T09:00:00+01:00",
    "finished_at": "2026-11-12T09:20:00+01:00", "excel_version": "16.0",
    "excel_build": "Version 2510 (build 19328)", "office_bitness": "64-bit",
    "os": "Windows 11 Pro (version 10.0.26100, build 26100)", "os_architecture": "x64",
    "references": ["Visual Basic For Applications", "Microsoft Excel 16.0 Object Library", "OLE Automation",
                   "Microsoft Office 16.0 Object Library"],
    "macro_policy": "Disable VBA macros with notification", "vba_project_access": "On",
    "compile": {"status": "PASS", "detail": "Debug > Compile VBAProject on IRRBB_Dev.xlsm"},
    "workbook_closed": True,
}


class BundleTests(unittest.TestCase):
    sha: str
    exp: ev.Expected

    @classmethod
    def setUpClass(cls) -> None:
        cls.sha = head_sha()
        cls.exp = ev.expected(ROOT, cls.sha)

    def setUp(self) -> None:
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.bundle = Path(temp.name)
        self.logs = passing_logs(self.exp, self.sha)
        self.session = json.loads(json.dumps(SESSION))

    def write(self) -> None:
        for name, lines in self.logs.items():
            text = f"CANDIDATE={self.sha}\r\n" + "".join(f"{k}={v}\r\n" for k, v in lines)
            (self.bundle / name).write_text(text, encoding="utf-8", newline="")
        (self.bundle / "session.log").write_text("09:00:00  started\n", encoding="utf-8")
        (self.bundle / "session.json").write_text(json.dumps(self.session), encoding="utf-8")

    def run_writer(self) -> tuple[int, str]:
        self.write()
        out, err = io.StringIO(), io.StringIO()
        with redirect_stdout(out), redirect_stderr(err):
            code = write_excel_evidence.main(["--root", str(ROOT), "--candidate-sha", self.sha,
                                              "--bundle", str(self.bundle)])
        return code, out.getvalue() + err.getvalue()

    def run_checker(self, sha: str | None = None) -> tuple[int, str]:
        out = io.StringIO()
        args = ["--root", str(ROOT), "--bundle", str(self.bundle)] + (["--candidate-sha", sha] if sha else [])
        with redirect_stdout(out):
            code = check_excel_evidence.main(args)
        return code, out.getvalue()

    def stage(self, name: str) -> dict[str, object]:
        host = json.loads((self.bundle / "host.json").read_text(encoding="utf-8"))
        result: dict[str, object] = host["stages"][name]
        return result

    def breaks(self, stage: str, change: Callable[[], None], reason: str) -> None:
        change()
        code, output = self.run_writer()
        self.assertEqual(code, 1, output)
        self.assertEqual(self.stage(stage)["outcome"], "FAIL")
        self.assertRegex("; ".join(str(r) for r in self.stage(stage)["reasons"]), reason)  # type: ignore[attr-defined]
        self.assertEqual(self.run_checker()[0], 1)

    def replace(self, log: str, key: str, old: str, new: str, count: int = 1) -> None:
        lines = self.logs[log]
        hits = [i for i, (k, v) in enumerate(lines) if k == key and v == old]
        self.assertGreaterEqual(len(hits), count, (log, key, old))
        for i in hits[:count]:
            lines[i] = (key, new)

    def test_a_passing_bundle_validates(self) -> None:
        code, output = self.run_writer()
        self.assertEqual(code, 0, output)
        self.assertIn(f"RESULT=PASS; candidate={self.sha}", output)
        host = json.loads((self.bundle / "host.json").read_text(encoding="utf-8"))
        self.assertEqual(host["result"], "PASS")
        self.assertEqual(set(host["stages"]), set(ev.STAGES))
        self.assertIs(host["environment"]["trust_changes"], False)
        self.assertEqual(set(host["sources"]), set(self.exp.sources))
        self.assertEqual(self.run_checker(self.sha), (0, "RESULT=PASS\n"))

    def test_a_record_is_never_replaced_and_names_its_candidate(self) -> None:
        self.assertEqual(self.run_writer()[0], 0)
        out, err = io.StringIO(), io.StringIO()
        with redirect_stdout(out), redirect_stderr(err):
            code = write_excel_evidence.main(["--root", str(ROOT), "--candidate-sha", self.sha,
                                              "--bundle", str(self.bundle)])
        self.assertEqual(code, 2)
        self.assertEqual(self.run_checker("0" * 40)[0], 1)

    def test_a_log_changed_after_the_record_fails(self) -> None:
        self.assertEqual(self.run_writer()[0], 0)
        path = self.bundle / "harness.log"
        path.write_text(path.read_text(encoding="utf-8") + "NOTE=edited\r\n", encoding="utf-8")
        code, output = self.run_checker()
        self.assertEqual(code, 1)
        self.assertIn("logs does not match", output)

    def test_an_edited_result_fails(self) -> None:
        self.logs["harness.log"] = harness_run(self.exp, self.sha, "TEST_Harness.RunTests", {"HARN-OUT-01": "FAIL"})
        self.assertEqual(self.run_writer()[0], 1)
        path = self.bundle / "host.json"
        host = json.loads(path.read_text(encoding="utf-8"))
        host["result"] = "PASS"
        host["stages"]["smoke"] = {"outcome": "PASS", "reasons": []}
        path.write_text(json.dumps(host), encoding="utf-8")
        self.assertEqual(self.run_checker()[0], 1)

    def test_edited_host_metadata_fails(self) -> None:
        def set_in(section: str, field: str, value: Any) -> Callable[[dict[str, Any]], None]:
            return lambda host: host[section].__setitem__(field, value)

        edits: list[tuple[str, Callable[[dict[str, Any]], None]]] = [
            ("environment", set_in("environment", "office_bitness", "32-bit")),
            ("environment", set_in("environment", "excel_build", "Version 2601")),
            ("environment", set_in("environment", "trust_changes", True)),
            ("operator", lambda host: host.__setitem__("operator", "Someone else")),
            ("started_at", lambda host: host.__setitem__("started_at", "")),
            ("finished_at", lambda host: host.pop("finished_at")),
            ("extra", lambda host: host.__setitem__("extra", "field")),
        ]
        for key, edit in edits:
            with self.subTest(key=key):
                self.setUp()
                self.assertEqual(self.run_writer()[0], 0)
                path = self.bundle / "host.json"
                host = json.loads(path.read_text(encoding="utf-8"))
                edit(host)
                path.write_text(json.dumps(host), encoding="utf-8")
                code, output = self.run_checker()
                self.assertEqual(code, 1)
                self.assertIn(f"host.json {key} does not match", output)

    def test_every_log_is_bound_to_the_candidate(self) -> None:
        self.write()
        path = self.bundle / "import.log"
        path.write_text(path.read_text(encoding="utf-8").replace(self.sha, "1" * 40), encoding="utf-8")
        host = ev.record(ROOT, self.sha, self.bundle)
        self.assertEqual(host["stages"]["import"]["reasons"], [f"import.log does not start with CANDIDATE={self.sha}"])

    def test_a_missing_log_fails_its_stage(self) -> None:
        self.write()
        (self.bundle / "roundtrip.log").unlink()
        host = ev.record(ROOT, self.sha, self.bundle)
        self.assertEqual(host["result"], "FAIL")
        self.assertEqual(host["stages"]["round_trip"]["reasons"], ["roundtrip.log is missing"])

    def test_template_failures(self) -> None:
        cases: list[tuple[Callable[[], None], str]] = [
            (lambda: self.replace("template.log", "CAPTION", "IRRBB_Template.xlsx  [Read-Only]",
                                  "IRRBB_Template.xlsx  [Repaired]"), "repaired"),
            (lambda: self.replace("template.log", "SHEET", "Checks|shChecks", "Checks|Sheet10"), "CodeNames"),
            (lambda: self.replace("template.log", "NAME", "Readme_SourceCommit|=Readme!$B$5",
                                  "Readme_SourceCommit|=Readme!$B$6"), "Readme_SourceCommit"),
            (lambda: self.replace("template.log", "VALIDATION", "Config!B11|x|refused", "Config!B11|x|accepted"),
             "Config!B11"),
            (lambda: self.replace("template.log", "FONT", "Readme|Aptos", "Readme|Calibri"), "Aptos"),
        ]
        for change, reason in cases:
            with self.subTest(reason=reason):
                self.setUp()
                self.breaks("template", change, reason)

    def test_harness_failures(self) -> None:
        cases: list[tuple[str, Callable[[], None], str]] = [
            ("smoke", lambda: self.replace("harness.log", "CASE", "HARN-OUT-01|PASS|3|3|0|",
                                           "HARN-OUT-01|FAIL|4|3|1|"), "HARN-OUT-01: FAIL"),
            ("smoke", lambda: self.replace("harness.log", "STATE_AFTER", STATE, "-4135|True|True|True"),
             "state changed"),
            ("smoke", lambda: self.replace("harness.log", "STATUS", ev.COMPLETE, "COMPLETE - FAILURES"), "status"),
            ("smoke", lambda: self.replace("harness.log", "INFO", self.sha, "0" * 40), "source commit"),
            ("smoke", lambda: self.replace("harness.log", "CASE", "RATE-OLS-01|NOT RUN|0|0|0|",
                                           "RATE-OLS-01|PASS|3|3|0|"), "RATE-OLS-01: PASS, expected NOT RUN"),
            ("failure_path", lambda: self.replace("failure.log", "CASE", "DATA-CODES-01|ERROR|0|0|0|",
                                                  "DATA-CODES-01|PASS|3|3|0|"), "DATA-CODES-01"),
            ("failure_path", lambda: self.replace(
                "failure.log", "INFO",
                "Cleanup error -2147220503: Injected results write fault [TEST_Harness.WriteResults]", ""),
             "cleanup error"),
            ("failure_path", lambda: self.replace("failure.log", "ASSERTION_ROWS", "0", "12"), "not empty"),
            ("failure_path", lambda: self.logs["failure.log"].__delitem__(0), "four runs"),
        ]
        for stage, change, reason in cases:
            with self.subTest(reason=reason):
                self.setUp()
                self.breaks(stage, change, reason)

    def test_import_failures(self) -> None:
        cases: list[tuple[Callable[[], None], str]] = [
            (lambda: self.replace("import.log", "STATUS", "FAILED - IMPORT REJECTED (2 errors)", ev.ACCEPTED),
             "control-differ: status"),
            (lambda: self.replace("import.log", "FINDING", "control_totals_balance.csv|4|E13|",
                                  "control_totals_balance.csv|4|E04|"), "findings"),
            (lambda: self.replace("import.log", "TOTAL_ROWS", "0", str(self.exp.totals)), "still shown"),
            (lambda: self.replace("import.log", "LINEAGE", f"Control-totals SHA-256|{self.exp.control_sha}",
                                  "Control-totals SHA-256|" + "0" * 64), "control_totals.csv"),
            (lambda: self.replace("import.log", "SCENARIO", "pasted-attestation", "pasted-window"), "in order"),
            (lambda: self.replace("import.log", "TOTAL_ROWS", str(self.exp.totals), "22"), "accepted: 22 total rows"),
        ]
        for change, reason in cases:
            with self.subTest(reason=reason):
                self.setUp()
                self.breaks("import", change, reason)
        self.setUp()
        lines = self.logs["import.log"]
        start = lines.index(("SCENARIO", "pasted-window"))
        status = next(i for i in range(start, len(lines)) if lines[i][0] == "STATUS")
        lines[status] = ("STATUS", ev.ACCEPTED)
        self.breaks("import", lambda: None, "pasted-window: status")

    def test_synchronizer_and_round_trip_failures(self) -> None:
        cases: list[tuple[str, Callable[[], None], str]] = [
            ("synchronize", lambda: self.replace("sync.log", "EXIT", "0", "1"), "exited with code 1"),
            ("synchronize", lambda: self.replace("sync.log", "COMPILE", "PASS", "FAIL"), "did not compile"),
            ("synchronize_failure", lambda: self.replace("sync-failure.log", "EXIT", "1", "0"), "did not fail"),
            ("synchronize_failure", lambda: self.replace("sync-failure.log", "HASH_AFTER", "ab" * 32, "cd" * 32),
             "workbook changed"),
            ("synchronize_failure", lambda: self.replace("sync-failure.log", "VBA_PROJECT_ACCESS_AFTER", "On", "Off"),
             "vba project access changed"),
            ("round_trip", lambda: self.replace("roundtrip.log", "OUTPUT",
                                                "RESULT=PASS; components=31; differences=0; binding=VERIFIED",
                                                "RESULT=PASS; components=31; differences=0; binding=UNVERIFIED"),
             "binding=VERIFIED"),
            ("round_trip", lambda: self.replace("roundtrip.log", "EXIT", "0", "1"), "exited with codes"),
        ]
        for stage, change, reason in cases:
            with self.subTest(reason=reason):
                self.setUp()
                self.breaks(stage, change, reason)

    def test_session_failures(self) -> None:
        cases: list[tuple[str, Callable[[], None], str]] = [
            ("references", lambda: self.session["references"].append("Microsoft Scripting Runtime"), "four defaults"),
            ("compile", lambda: self.session.__setitem__("compile", {"status": "FAIL"}), "compile FAIL"),
            ("cleanup", lambda: self.session.__setitem__("workbook_closed", False), "not closed"),
        ]
        for stage, change, reason in cases:
            with self.subTest(reason=reason):
                self.setUp()
                self.breaks(stage, change, reason)

    def test_the_candidate_must_be_a_full_sha(self) -> None:
        self.write()
        with self.assertRaises(ev.EvidenceError):
            ev.record(ROOT, self.sha[:7], self.bundle)


class ScriptShapeTests(unittest.TestCase):
    """What the validator expects is what the script writes."""

    def setUp(self) -> None:
        self.text = SCRIPT.read_text(encoding="ascii")

    def test_the_script_and_launcher_are_ascii_with_crlf(self) -> None:
        for path in (SCRIPT, LAUNCHER):
            raw = path.read_bytes()
            raw.decode("ascii")
            self.assertNotIn(b"\n", raw.replace(b"\r\n", b""), path.name)

    def test_logs_scenarios_and_runs_match_the_validator(self) -> None:
        self.assertEqual(re.findall(r'New-EvidenceLog "([a-z-]+\.log)"', self.text),
                         ["template.log", "build.log", "harness.log", "failure.log", "import.log",
                          "sync.log", "roundtrip.log", "sync-failure.log"])
        self.assertEqual(set(re.findall(r'New-EvidenceLog "([a-z-]+\.log)"', self.text)), set(ev.LOGS))
        scenarios = re.findall(r'Write-ImportScenario \$excel \$book \$qualified \$log "([a-z-]+)"', self.text)
        self.assertEqual(scenarios, list(ev.SCENARIOS))
        runs = re.search(r'foreach \(\$entry in @\(([^)]*)\)\)', self.text)
        assert runs is not None
        self.assertEqual(re.findall(r'"([\w.]+)"', runs.group(1)), list(ev.FAILURE_RUNS))

    def test_the_script_never_changes_security_settings(self) -> None:
        for forbidden in ("Set-ItemProperty", "New-ItemProperty", "Remove-ItemProperty", "reg add", "reg.exe"):
            self.assertNotIn(forbidden, self.text)
        self.assertIn("This script never changes it.", self.text)

    def test_a_range_is_returned_whole(self) -> None:
        # Returned bare from a function, an enumerable Range is unrolled into its cells.
        body = re.search(r"function Get-Area \{(.*?)\n\}", self.text, re.DOTALL)
        assert body is not None
        self.assertIn("return ,(", body.group(1))

    def test_cell_arrays_are_read_with_their_true_bounds(self) -> None:
        # Excel's Value2 arrays start at 1; PowerShell's own indexer counts from 0.
        self.assertNotRegex(self.text, r"\$data\[")
        self.assertIn("GetValue(", self.text)

    def test_the_script_never_waits_on_a_printer(self) -> None:
        # Exporting or printing makes Excel query the default printer, which can stall a run.
        for call in ("ExportAsFixedFormat", "PrintOut", "PrintPreview", "PageSetup", "ActivePrinter"):
            self.assertNotIn(call, self.text)

    @unittest.skipUnless(PWSH, "PowerShell 7 (pwsh) is not installed")
    def test_the_script_parses(self) -> None:
        command = ("$errors = $null; [void][System.Management.Automation.Language.Parser]::ParseFile("
                   f"'{SCRIPT}', [ref]$null, [ref]$errors); "
                   "if ($errors.Count) { $errors | ForEach-Object { $_.Message }; exit 1 }")
        completed = subprocess.run([str(PWSH), "-NoProfile", "-NonInteractive", "-Command", command],
                                   capture_output=True, text=True, timeout=120, check=False)
        self.assertEqual(completed.returncode, 0, completed.stdout + completed.stderr)


if __name__ == "__main__":
    unittest.main()
