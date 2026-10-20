<div align="center">

# 📦 Installation and Developer Setup

### Set up, validate, and later import K-IRRBB into Excel

[![Status](https://img.shields.io/badge/Status-In_development-d97706?style=flat-square)](#current-status)
[![Checks](https://img.shields.io/badge/Checks-python_tools%2Fcheck.py-0969da?style=flat-square)](#run-the-checks)
[![Python](https://img.shields.io/badge/Python-3.10%2B-3776ab?style=flat-square&logo=python&logoColor=white)](#prerequisites)
[![Security](https://img.shields.io/badge/Security-Private_policy-d73a49?style=flat-square)](SECURITY.md)

<br>

**One identifiable commit · Clean import · Compile · Validate · Preserve caller state**

</div>

---

This document is authoritative for **developer setup, import, upgrade, recovery
and removal**. Contribution workflow is owned by
[`CONTRIBUTING.md`](CONTRIBUTING.md), vulnerability handling by
[`SECURITY.md`](SECURITY.md), and publication by [`RELEASING.md`](RELEASING.md).

> [!IMPORTANT]
> VBA executes with the user's Office permissions. Review the exact source,
> follow organizational macro policy, and never enable macros in an untrusted
> workbook.

<a id="current-status"></a>

## 🧭 Current status

Milestone **v0.2.0 – Workbook, data and Excel tests** is in progress;
**v0.1.0 – Repository foundation** closed on 2026-10-09. There is **no
IRRBB release** and **no validated model**: the repository holds
documentation, issue governance, CI, static source checks, the data import
(core, public facade and workbook adapter), the workbook template, the
synchronizer and source round trip, the regression harness and the draft
decay model under review in #36.

| Topic | Status |
| --- | --- |
| VBA source layout | Defined in [`docs/REPOSITORY_STRUCTURE.md`](docs/REPOSITORY_STRUCTURE.md) |
| VBA conventions | Defined in [`docs/VBA_HOUSE_STYLE.md`](docs/VBA_HOUSE_STYLE.md) |
| Static checks | Available: `python tools/check.py` |
| VBA source | `src/core/`: CORE_Parse (dates, decimals, rates), CORE_Codes (segments, currency allowlist, identifiers), CORE_Sha256 (data fingerprint) and CORE_Import (validation, summary, outcomes and point-in-time view of the data files); `src/modules/IRRBB_Import` (public import from files) and `src/workbook/WB_Import` (runs the import from `Config` into `DataSummary`), #71; the draft decay model CORE_Math, CORE_Decay, CORE_Overlay and `src/modules/IRRBB_Decay` (#36, not validated) |
| Workbook template | [`src/workbook/IRRBB_Template.xlsx`](src/workbook/IRRBB_Template.xlsx), macro-free, generated from source; layout in [`docs/WORKBOOK_CONTRACT.md`](docs/WORKBOOK_CONTRACT.md) (#5); not yet opened in Excel |
| Regression harness | `tests/modules/` (#8): runs from a [development workbook](#development-workbook); no Excel run recorded yet |
| Supported hosts and references | Decided in issue #3; see [supported hosts](#supported-hosts) |
| Excel evidence | Requirements defined in [`docs/EXCEL_EVIDENCE.md`](docs/EXCEL_EVIDENCE.md); no record yet |
| Released versions | None |

Passing static checks is **not** evidence of correct IRRBB calculations.

<a id="prerequisites"></a>

## 🧰 Prerequisites

| Tool | Needed for |
| --- | --- |
| Git | Cloning and all repository work |
| Python 3.10 or later | `python tools/check.py`; standard library only, no packages |
| Node.js 20 or later | Optional: local label-catalog checks |
| Microsoft Excel for Windows | VBA import, compile, the regression harness and the data import, and later the models; see [supported hosts](#supported-hosts); not needed for static checks |

<a id="supported-hosts"></a>

## 🖥️ Supported hosts

Decided on 2026-10-08; see the
[acceptance record in #3](https://github.com/danielep71/K-IRRBB_Behavioral_Models/issues/3#issuecomment-6068524873).

| Host | Support level | Evidence required |
| --- | --- | --- |
| Microsoft 365 or Excel 2016 and later, **Windows, 64-bit** | Supported target | Compile, regression and smoke run for every VBA change and release |
| Same versions, Windows, 32-bit | Best effort: kept compiling, not certified | Compile when a 32-bit host is available; otherwise recorded as `NOT_RUN` |
| Excel for Mac, Excel for the web, Excel 2013 and earlier | Not supported | None |

<a id="vba-references"></a>

**VBA references:** only the four defaults of a new workbook:

1. *Visual Basic For Applications*
2. *Microsoft Excel 16.0 Object Library*
3. *OLE Automation*
4. *Microsoft Office 16.0 Object Library*

Adding any other reference, add-in (including Solver) or late-bound
`CreateObject` dependency needs an issue and an update to this section first.

<a id="runtime-dependencies"></a>

**Runtime dependencies:** none beyond Excel and VBA. The workbook never calls
Python, R, a database, a web service or any other external engine. Python or R
may be used **outside** the workbook only to produce independent reference
values for tests; such scripts are not part of the deliverable, and their
results are committed as reviewed expected values with their provenance
([`docs/methodology/VALIDATION_PLAN.md`](docs/methodology/VALIDATION_PLAN.md)).
Python is also used by the static checks, which never run with the workbook.

<a id="commitments-and-evidence"></a>

**Commitments are not evidence.** The tables above state what the project
commits to support. Whether a given commit actually compiles and runs on a
host is shown only by an Excel evidence record bound to that commit, as defined
in [`docs/EXCEL_EVIDENCE.md`](docs/EXCEL_EVIDENCE.md). No host is listed as
tested until such a record exists; none does yet.

## 📥 Get the source

Use a **Git clone**:

```sh
git clone https://github.com/danielep71/K-IRRBB_Behavioral_Models.git
cd K-IRRBB_Behavioral_Models
git switch release/v0.2.0
```

`release/v0.2.0` is the active development branch; `main` receives it only on
the owner's request.

Do not use **Code → Download ZIP**. `.gitattributes` excludes repository
plumbing such as `.gitattributes`, `.gitignore`, `.editorconfig` and `.github/`
from GitHub archives, so a ZIP is not a complete checkout and the checks will
not behave as documented.

A clone checks VBA files out with the CRLF line endings the VBE expects. Do not
download individual modules through GitHub's **Raw** view: it serves the LF
form stored in Git.

<a id="run-the-checks"></a>

## ✅ Run the checks

From the repository root:

```sh
python tools/check.py
```

All gates must pass. Before pushing a committed candidate, the stricter form
also requires a clean working tree:

```sh
python tools/check.py --ci
python tools/check.py --ci --base <base-sha>   # for a pull request's full range
```

Optional label-catalog checks:

```sh
node .github/scripts/labels-sync.mjs --policy .github/label-policy.json --self-test
node .github/scripts/labels-drift.mjs --policy .github/label-policy.json --self-test
```

The gates are described in [`tools/README.md`](tools/README.md). They are static
checks: they never compile VBA, run Excel or validate a model.

<a id="importing-vba-into-excel"></a>

## 📂 Importing VBA into Excel

<a id="development-workbook"></a>

### Development workbook

For compile and regression evidence, build the workbook from the checkout
instead of importing by hand:

```shell
python tools/devbuild/build_dev_workbook.py   # writes build/IRRBB_Dev.xlsm
```

The workbook is the [template](docs/WORKBOOK_CONTRACT.md) with the VBA
project added. The builder embeds every `.bas` under `src/core/`, `src/modules/`,
`src/workbook/`, `tests/modules/` and `examples/modules/` unchanged and stamps
the commit on the `Readme` sheet; a dirty checkout is stamped as not
evidence-grade. Open the file, enable macros, run **Debug → Compile
VBAProject**, then **Run tests** on the `Checks` sheet. To import data, fill
the settings on `Config` (or use **Choose account file** and **Choose rate
file**), attest timely closure reporting, and click **Import data**; the result
appears on `DataSummary`. Details and the
LibreOffice early-warning run are in
[`tools/devbuild/README.md`](tools/devbuild/README.md).

<a id="updating-an-existing-workbook"></a>

### Updating an existing workbook

To bring a workbook saved from the template (or an earlier build) up to the
source of the checkout, close it and drag it onto `tools/Sync-IRRBB-VBA.cmd`.
The synchronizer asks whether it is a production or a development workbook and
replaces the VBA components from `src/` (and, for development, `tests/` and
`examples/`) without touching sheets or data. To try modules downloaded for a
review without copying them into the checkout, use
`tools/Sync-IRRBB-VBA-Downloads.cmd` instead. Both need Windows Excel with
**Trust access to the VBA project object model** enabled; see
[`tools/README.md`](tools/README.md#local-vba-synchronization). Afterwards run
**Debug → Compile VBAProject** and enter the commit in `Readme!B5` as in step 6
below.

### Manual import

Use these steps when neither the builder nor the synchronizer can be used. Components at present:
`src/core/CORE_Codes.bas`, `src/core/CORE_Parse.bas`,
`src/core/CORE_Import.bas`, `src/core/CORE_Sha256.bas`, `src/modules/IRRBB_Import.bas`,
`src/workbook/WB_Import.bas`; development only: `tests/modules/TEST_Harness.bas`,
`tests/modules/TEST_CoreCases.bas`, `tests/modules/TEST_WorkbookCases.bas`,
`tests/modules/TEST_DataCases.bas`, `tests/modules/TEST_Fixtures.bas`,
`tests/modules/TEST_ImportCases.bas`, `tests/modules/TEST_HashCases.bas`.

1. Start from **one exact commit** in a Git checkout, so `.bas`, `.cls` and
   `.frm` files have CRLF line endings. Never mix components from different
   commits.
2. Open the macro-free template `src/workbook/IRRBB_Template.xlsx` and save
   it as a macro-enabled workbook (its sheets already carry the CodeNames of
   [`docs/WORKBOOK_CONTRACT.md`](docs/WORKBOOK_CONTRACT.md))
   **outside** the checkout or in an ignored location. Built workbooks are never
   committed.
3. Open the VBE (**Alt+F11**) and check that **Tools → References** lists only
   the [four default references](#vba-references).
4. Import in dependency order: `src/core/`, then `src/classes/` (none exist
   yet), then `src/modules/`, then the standard modules in `src/workbook/`,
   then `src/forms/` (none yet; each `.frm` with its `.frx`). For a development
   workbook only, add the `tests/` and `examples/` modules.
5. Paste the code of the document modules (`src/workbook/ThisWorkbook.cls` and
   the sheet modules `src/workbook/sh*.cls`) into the existing module instead
   of importing it, which would create a duplicate class.
6. Run **Debug → Compile VBAProject**; it must complete with no error.
   Enter the 40-digit commit of the imported source in `Readme` cell `B5`
   (`Readme_SourceCommit`); the data import refuses to run without it, since
   the lineage must name the code version.
7. Run the synthetic regression harness and the specific scenario under test,
   then record the result as described under
   [validation record](#validation-record).

Do not paste source into arbitrarily named modules: the component name
(`VB_Name`) is part of the contract, and `tools/check_source.py` requires it to
match the file name.

### Exporting changes back

To check a workbook against a commit without changing the checkout, export it
with `tools/Export-IRRBB-VBA.ps1` and compare with `tools/compare_vba_export.py`
([source round trip](tools/README.md#source-round-trip)). To bring an edit made
in the VBE back into the repository:

1. In the VBE, export each changed component over its file in the checkout, in
   the same folder.
2. Review `git diff`. The VBE silently changes the capitalization of an
   identifier everywhere it appears when one declaration changes case; revert
   case-only changes you did not intend.
3. Run `python tools/check.py`.
4. Commit. Git stores the files with LF; you do not convert anything by hand.

<a id="validation-record"></a>

## 🧪 Validation record

A successful import is not certification. Record at least:

```text
Commit (full SHA):
Files imported:
Excel / Office version and build:
Office bitness:
Operating system:
Compile:
Regression harness:
Reference cases, tolerances and results:
Specific scenario:
Cleanup:
Skipped or unverified:
```

A skipped, incomplete or cleanup-failed run is not a pass. Static checks cannot
substitute for Excel execution. Commit-bound evidence follows
[`docs/EXCEL_EVIDENCE.md`](docs/EXCEL_EVIDENCE.md).

## ⬆️ Upgrade

Before moving to a newer commit:

1. read the changelog between the two commits;
2. back up the host workbook, parameter sets and any user configuration;
3. identify the complete component set at the target commit; and
4. review compatibility notes and known limitations.

Replace the whole component set, compile again and repeat the validation. Do not
infer backward compatibility from successful compilation alone. A change in
estimated parameters between commits must be explained by the changelog.

Treat a locally modified copy as a fork: diff it against the old and new source,
reapply modifications deliberately and retest them.

## 🧯 Troubleshooting

| Symptom | Check first |
| --- | --- |
| `tools/check.py` fails on a clean clone | Python 3.10+ is on `PATH`, and the checkout is a Git clone, not a ZIP. |
| `tools/check.py --ci` fails locally | The working tree is clean and every change is committed. |
| `tools/check.py` reports CRLF in Git | Renormalize with `git add --renormalize .` and commit the result separately. |
| Compile error or missing procedure | All components come from one commit, and required references are set. |
| Ambiguous name | Remove duplicate or older copies of a component. |
| Different result from a reference | Confirm the commit, inputs, data window, settings and the independent reference used. |
| Security warning | Verify the source origin and organizational macro settings. |

Suspected security problems follow [`SECURITY.md`](SECURITY.md), not issues.

## 🗑️ Removal

1. Run any project-owned shutdown or cleanup procedure.
2. Remove the components and integrations the project owns.
3. Compile the remaining VBA project.
4. Close and reopen, and verify nothing still depends on removed components.

Removing modules does not remove formulas, names, links, connections or other
host integrations. Remove only state the project owns.

## 📚 Related documents

- [`README.md`](README.md) — overview and status
- [`CONTRIBUTING.md`](CONTRIBUTING.md) — change and review workflow
- [`tools/README.md`](tools/README.md) — static checks and CI
- [`RELEASING.md`](RELEASING.md) — release sequence
- [`SECURITY.md`](SECURITY.md) — vulnerability and data-handling policy

---

**Installation principle:** use one identifiable commit, compile it, exercise
its real behavior in Excel, and record what was and was not validated.
