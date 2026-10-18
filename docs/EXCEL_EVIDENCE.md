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
> At v0.1.0 there is no VBA source, so no record exists yet. Until a validator
> is added, records are checked by review against this document.

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
| Security context | Macro setting and *Trust access to the VBA project object model*, as found and **not changed** |
| Operator and time | Who ran it, start and finish time with UTC offset |
| Outcomes | One outcome per stage below, with the copied output that supports it |

Any change to the source after the record invalidates it for the affected
components. Evidence is never carried from one commit to another.

<a id="required-stages"></a>

## 🧪 Required stages

### 1. Clean build

Build the workbook from the template at the candidate commit, outside the
checkout, exactly as in
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
   table (`Checks!A13:F46`) into the log; and
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

<a id="bitness"></a>

## 🖥️ Bitness

| Host | Stages required | When no host is available |
| --- | --- | --- |
| Windows, 64-bit | All five | The change cannot be merged or released; record `UNAVAILABLE` and stop |
| Windows, 32-bit (best effort) | Stages 1 and 2 | Record `NOT_RUN` with the reason; the change may proceed |

A 32-bit `PASS` is reported as best-effort compile evidence, never as
certification. `Declare` statements stay `PtrSafe` with `LongPtr` where
required, so that both bitnesses keep compiling; `tools/check.py` checks the
conditional-compilation branches statically.

<a id="procedure"></a>

## 🧑‍💻 Procedure

Execution is **manual**: a person runs Excel interactively. The procedure never
changes macro security or Trust Center settings.

1. Check out the candidate in a clean tree: `git switch --detach FULL_SHA`,
   then `git status`.
2. Create a record folder **outside** the checkout and start a plain-text
   `session.log` in UTF-8 with the SHA and the start time.
3. Record the host and security context without changing them.
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
