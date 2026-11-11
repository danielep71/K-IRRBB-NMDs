<div align="center">

# 🪟 Excel Evidence

### What a real Excel run must record before anyone can rely on it

[![Execution](https://img.shields.io/badge/Execution-manual-1d76db?style=flat-square)](#procedure)
[![Binding](https://img.shields.io/badge/Binding-exact_SHA-217346?style=flat-square)](#what-a-record-binds)
[![Hosts](https://img.shields.io/badge/Hosts-64--bit_target-6f42c1?style=flat-square)](../INSTALLATION.md#supported-hosts)

<br>

**One commit · One host · Observed values only · NOT_RUN is never a pass**

</div>

---

This document is authoritative for **the Excel evidence record: what a compile
and smoke run must capture, how it is bound to one commit, and which outcomes
count**. How to build the workbook is in
[`INSTALLATION.md`](../INSTALLATION.md#importing-vba-into-excel); which hosts
need evidence is in [supported hosts](../INSTALLATION.md#supported-hosts); when
evidence is required is in [`CONTRIBUTING.md`](../CONTRIBUTING.md#validation-and-evidence)
and [`RELEASING.md`](../RELEASING.md#release-invariants).

> [!IMPORTANT]
> `python tools/check.py` and CI never compile VBA or run Excel. Only a record
> made by a person running Excel, bound to an exact commit, is Excel evidence.
> No record exists yet for the VBA source of v0.2.0.
> [`tools/Certify-IRRBB.ps1`](../tools/Certify-IRRBB.ps1) runs every stage
> below in one command and writes the record;
> [`tools/check_excel_evidence.py`](../tools/check_excel_evidence.py) validates
> it offline. A record made by hand is checked by review against this document.

<a id="outcomes"></a>

## 🚦 Outcomes

Every stage of a record has exactly one outcome.

| Outcome | Meaning | Counts as a pass? |
| --- | --- | :---: |
| `PASS` | The stage ran on the stated host and met every expectation | ✅ |
| `FAIL` | The stage ran and at least one expectation was not met | ❌ |
| `NOT_RUN` | The stage was not executed, with the reason stated | ❌ |
| `UNAVAILABLE` | No suitable host was available, with the reason stated | ❌ |

A skipped, interrupted or cleanup-failed run is not a pass. A later passing
run never replaces an earlier failing one in the record; both are kept.

<a id="what-a-record-binds"></a>

## 🔗 What a record binds

| Item | Recorded |
| --- | --- |
| Candidate | Full 40-character commit SHA; the checkout was clean |
| Sources | Every imported component and the workbook template, all from that one commit |
| Host | Excel version and build, Office bitness, Windows edition, version and architecture |
| References | The entries in **Tools → References**, which must be exactly the [four defaults](../INSTALLATION.md#vba-references) |
| Security context | Macro setting and *Trust access to the VBA project object model*, as found and **not changed**, except Trust access enabled for stage 6 and turned off after it, when so recorded |
| Operator and time | Who ran it, start and finish time with UTC offset |
| Outcomes | One outcome per stage below, with the copied output that supports it |

Any change to the source after the record invalidates it for the affected
components. Evidence is never carried from one commit to another.

<a id="one-command-session"></a>

## ⚡ One-command session

`tools/Certify-IRRBB.ps1` runs the [required stages](#required-stages) for the
checked-out commit in one run, on Windows desktop Excel. It needs Python 3.11
or later and *Trust access to the VBA project object model* already on: the
compile through the VBE, the synchronizer and the export need it for the whole
session, and the script never turns it on or off. Turn it on before the
session and off after it, if that is your practice; the record shows it as
found. Check out the candidate in a clean tree and double-click
`tools\Certify-IRRBB.cmd`, or run:

```powershell
git fetch origin
git switch --detach FULL_SHA
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\Certify-IRRBB.ps1
```

The script refuses a working tree with changes. It creates the bundle folder
`irrbb-evidence-SHORTSHA` next to the checkout (`-BundleDir` names another; a
folder that is not empty is refused), starts `session.log` and records the host
and its security settings without changing them. Each stage writes its own log,
first line `CANDIDATE=` and the SHA:

| Log | What the script does and records |
| --- | --- |
| `template.log` | Opens the macro-free template read-only: no repair, its sheets, CodeNames, defined names and title font, and every validated cell refusing an invalid value and accepting a valid one; exports the print view to `template.pdf` |
| `build.log` | Builds `IRRBB_Dev.xlsm` in the bundle with the [development builder](../INSTALLATION.md#development-workbook) (stage 1) |
| `harness.log` | Opens it in a visible Excel that follows the macro setting, compiles it (stage 2) and runs the harness (stage 3) |
| `failure.log` | The four runs of [stage 4](#4-failure-path) |
| `import.log` | **Import data** with the synthetic fixtures: accepted; control totals that match and that differ; a window date and an attestation pasted past the validation; a rerun. Each later result must withdraw the earlier one |
| `sync.log` | The template saved as `IRRBB_Sync.xlsm`, synchronized with `Sync-IRRBB-VBA.ps1 -Mode Development`, compiled and run through the harness |
| `roundtrip.log` | [Stage 6](#6-source-round-trip) on `IRRBB_Sync.xlsm` |
| `sync-failure.log` | The synchronizer run again while another Excel holds the workbook open: it must fail, and the workbook and the security settings must be unchanged |

Excel's calculation, events, screen updating and alerts are recorded before
and after every run. The development workbook is closed without saving
(stage 5). The script then writes `session.json` and runs
`tools/write_excel_evidence.py`, which judges every stage from the logs,
writes `host.json` and validates the bundle; the exit code is 0 only for
`PASS`. The operator clicks **Enable Content** when Excel asks, once for each
workbook, and otherwise leaves Excel alone until the run ends: a click in a
sheet during a run can change the state it records.

`host.json` binds the record: the candidate, the environment with
`trust_changes: false`, the SHA-256 of every source the workbook is built from
(the bytes stored in Git), of every log and of `session.json`, and each
stage's outcome with its reasons. The validator never runs Excel. It judges
the stages again from the logs and the repository at the candidate, with the
expected values read from that commit: the harness registry, the template
generator and the data-contract fixtures. It then compares everything with
`host.json`:

```sh
python tools/check_excel_evidence.py --bundle ../irrbb-evidence-SHORTSHA --candidate-sha FULL_SHA
```

A log changed after the record, a missing log, a stage that did not pass or a
record that names another commit is `RESULT=FAIL`. Keep a failing bundle as it
is: a later passing run never replaces it. Attach the bundle's `host.json` and
the validator's output to the issue or pull request.

<a id="required-stages"></a>

## 🧪 Required stages

### 1. Clean build

Build the workbook from the template at the candidate commit, outside the
checkout or in a location Git ignores (the development builder writes
`build/IRRBB_Dev.xlsm`), exactly as in
[Importing VBA into Excel](../INSTALLATION.md#importing-vba-into-excel).
Record the files imported, in order, and the reference list.

`PASS` requires: every component imported from the one commit, no extra or
missing component, and only the four default references.

### 2. Compile

Run **Debug → Compile VBAProject**.

`PASS` requires: compilation completes with no error dialog. Record the exact
error text and the highlighted component if it fails.

### 3. Smoke run

The smoke run shows that the built workbook works end to end on synthetic
data, not that the models are correct. It consists of:

1. opening the saved workbook afresh with macros enabled through the existing
   trust mechanism;
2. running the harness entry point `TEST_Harness.RunTests` (the **Run tests**
   button on `Checks`) and copying its status block (`Checks!A3:B11`) and case
   table (`Checks!A13:F48`) into the log; and
3. checking that `Calculation`, `EnableEvents`, `ScreenUpdating` and
   `DisplayAlerts` are the same after the run as before it.

`PASS` requires: the run completes, its own result line reports a pass with
the expected number of cases, and Excel state is restored.

### 4. Failure path

The harness must be able to show a failure, so that a passing smoke run is
known to be able to fail. Run, from the Immediate window, in this order:

1. `TEST_Harness.TEST_RunTestsWithInjectedFailures`: `PASS` requires status
   `COMPLETE - FAILURES`, `HARN-OUT-01` `FAIL` and `DATA-CODES-01` `ERROR`
   (both noted as injected), every other executed case `PASS`;
2. `TEST_Harness.RunTests`: `PASS` requires status
   `COMPLETE - ALL EXECUTED CASES PASSED` and no `FAIL` or `ERROR` row left;
3. `TEST_Harness.TEST_RunTestsWithWriteFault`: `PASS` requires status `FAILED`
   with both the primary and the cleanup error in `Checks!B10:B11` and empty
   result tables; and
4. `TEST_Harness.RunTests` again, to leave current results.

Excel state must be restored after each run, as in stage 3.

### 5. Cleanup

Close the test workbook without saving it into the checkout, and record that
it was closed. A built workbook is never committed.

### 6. Source round trip

Export the saved workbook with `tools/Export-IRRBB-VBA.ps1` into a new folder
and compare it with `python tools/compare_vba_export.py --export <folder>
--mode development --candidate-sha <sha> --workbook <workbook>`
([source round trip](../tools/README.md#source-round-trip)). Copy its output
into the log.

`PASS` requires `RESULT=PASS` with `binding=VERIFIED`: the workbook holds
exactly the components of its mode at the candidate commit. The export needs
*Trust access to the VBA project object model*; record whether it was already
enabled, and if it was enabled for this stage, that it was turned off again.

<a id="bitness"></a>

## 🖥️ Bitness

| Host | Stages required | When no host is available |
| --- | --- | --- |
| Windows, 64-bit | All six | The candidate cannot be released; record `UNAVAILABLE` and stop |
| Windows, 32-bit (best effort) | Stages 1 and 2 | Record `NOT_RUN` with the reason; the change may proceed |

A 32-bit `PASS` is reported as best-effort compile evidence, never as
certification. `Declare` statements stay `PtrSafe` with `LongPtr` where
required, so that both bitnesses keep compiling; `tools/check.py` checks the
conditional-compilation branches statically.

<a id="procedure"></a>

## 🧑‍💻 Procedure

When the script cannot be used, follow these steps by hand.

Execution is **manual**: a person runs Excel interactively. The procedure never
changes macro security or Trust Center settings, with one exception: stage 6
needs *Trust access to the VBA project object model*. If it is off, enable it
for stage 6 only, turn it off afterwards, and record both; a stage 6 `PASS` run
that way is valid.

1. Check out the candidate in a clean tree: `git switch --detach FULL_SHA`,
   then `git status`.
2. Create a record folder **outside** the checkout and start a plain-text
   `session.log` in UTF-8 with the SHA and the start time.
3. Record the host and security context without changing them (stage 6 apart,
   as above).
4. Run the [required stages](#required-stages) in order, copying each output
   into the log as it appears.
5. Write the outcome of every stage, with `NOT_RUN` or `UNAVAILABLE` and a
   reason where a stage was not executed.
6. Attach the log, or a summary that quotes its outcome lines, to the issue or
   pull request. Remove any path, user name or workbook content that is not
   needed.

Do not fill in a host, stage or result that was not observed.

## 🔐 Data in evidence

Evidence uses synthetic data only. Do not attach screenshots, workbooks or
logs that contain client data, internal paths beyond what is needed, or
proprietary content ([`SECURITY.md`](../SECURITY.md#data-and-secrets)).

## 📚 Related documents

- [`INSTALLATION.md`](../INSTALLATION.md) — supported hosts, references, build and import
- [`CONTRIBUTING.md`](../CONTRIBUTING.md#validation-and-evidence) — when evidence is required for a change
- [`RELEASING.md`](../RELEASING.md) — certification for a release
- [`methodology/VALIDATION_PLAN.md`](methodology/VALIDATION_PLAN.md) — numerical validation, separate from host evidence

---

**Evidence principle:** record what was observed on one host for one commit,
and call everything else by its name.
