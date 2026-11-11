# Repository tooling

The tools implement source-first conventions and portable source gates; they do not perform IRRBB calculations. Adapted files and their sources are listed in [`docs/licenses/THIRD_PARTY_NOTICES.txt`](../docs/licenses/THIRD_PARTY_NOTICES.txt).

```shell
python tools/check.py
python tools/check.py --ci
python tools/check.py --ci --base <base-sha>
```

The gate runner checks:
- Python-tool regression self-tests;
- committed / working-tree whitespace;
- source placement, LF storage and VBA export headers; every VBA component lies in a directory the synchronizer's `$sourceRoots` list reads, and the document modules in `src/workbook/` match the template's sheets by CodeName (`test_sync_checks.py`);
- the synchronizer's folder selection, document-module decoding, production/development answer and component sets, run in PowerShell 7 without Excel (`test_sync_script.py`; skipped where `pwsh` is not installed, as on the CI runner);
- the export comparison of the source round trip (`test_compare_vba_export.py`);
- data-contract fixtures in `tests/` agree with their hand-computed expectations and with `docs/methodology/DATA_CONTRACT.md` (`test_data_contract.py`);
- the VBA harness: every literal expectation in the `TEST_*Cases` modules agrees with the contract, the harness rules or the template generator, `TEST_Fixtures.bas` equals the output of `devbuild/fixture_module.py`, and the harness registry equals the register in `docs/methodology/TEST_CASES.md` (`test_vba_harness.py`; VBA is read, not run);
- the development-workbook builder: `vbaProject.bin` read back with an independent [MS-OVBA]/[MS-CFB] decoder (`test_devbuild.py`);
- the workbook template: `src/workbook/IRRBB_Template.xlsx` equals the output of `workbook/build_template.py`, agrees with `docs/WORKBOOK_CONTRACT.md` and the harness areas, and holds no formulas, VBA or document properties (`test_workbook_template.py`);
- `src/core/` host independence: no Excel object-model or UI identifiers outside comments and strings;
- VBA jump targets and conditional compilation across supported Windows bitness contexts;
- explicit public API declarations compared with `docs/PUBLIC_API.txt`;
- changelog/version consistency;
- Markdown links (`check_links.py`): every relative link names an existing
  file or directory in the repository, every fragment names a heading or
  anchor of its target, and every external link uses HTTPS, so a plain
  `http://` link is an error (`test_check_links.py`);
- the issue forms and label catalog (`check_issues.py --check-config`): each
  form assigns `danielep71`, sets its provisional priority and requires a
  start date, a target end date and the dependencies, and `blocked` is the only
  dependency-status label (`test_check_issues.py`).

CI runs these checks on GitHub's hosted Linux runner. **VBA compile, Excel workbook execution, econometric calibration and numerical result validation are not performed by these gates.** Record Excel host evidence separately under `docs/EXCEL_EVIDENCE.md`.

`workbook/build_template.py` generates the macro-free template (standard library only). `devbuild/` builds the macro-enabled development workbook used for Excel compile and harness evidence, without Excel automation and without Trust access to the VBA project object model; see [`devbuild/README.md`](devbuild/README.md).

`traffic_history.py` is not a gate: the daily traffic export workflow runs it
to merge each traffic collection into the `traffic-history` branch and to
prepare the traffic alert. Its offline tests (`test_traffic_history.py`) run
with the other self-tests; see [`docs/TRAFFIC.md`](../docs/TRAFFIC.md).

## Local VBA synchronization

`Sync-IRRBB-VBA.ps1` synchronizes the repository VBA source into an existing
local `.xlsm` workbook, such as one saved from the template, without
intentionally changing worksheets, cells, formulas, names, tables or
formatting.

For normal use on Windows, close the workbook and drag its `.xlsm` file onto
`tools/Sync-IRRBB-VBA.cmd`, or run the launcher with the workbook's full path
as its argument:

```bat
tools\Sync-IRRBB-VBA.cmd "C:\path\to\IRRBB.xlsm"
```

The launcher runs the PowerShell synchronizer on that workbook, shows the
result and pauses before closing. Started without a workbook, for example by
double-clicking it, it prints how to use it, pauses and exits with code 1
without starting PowerShell. Run the script directly through
`powershell.exe -ExecutionPolicy Bypass -File`, as the launcher does: where the
execution policy disables scripts, `.\tools\Sync-IRRBB-VBA.ps1` alone is
refused. The bypass applies to that one process and changes no setting.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\Sync-IRRBB-VBA.ps1 -WorkbookPath "C:\path\to\IRRBB.xlsm"
```

A full synchronization first asks which workbook it is building, unless
`-Mode Production` or `-Mode Development` gives the answer:

| Mode | Components | Use |
| --- | --- | --- |
| Production | The `production` configuration of `vba-projects.json` (the standard modules of `src/`) and the document modules | A workbook used with real data |
| Development | The `development` configuration: the production modules and every test and example module | A workbook that runs the regression harness, which writes to `Checks` and imports synthetic fixtures |

Any other answer cancels with exit code 2 and the workbook is not changed.
`vba-projects.json` is the single list of both sets; `test_devbuild.py` keeps
its development configuration equal to what the development-workbook builder
embeds. A component in the source folders that it does not list stops the
synchronization before Excel starts. A production synchronization never
removes modules: if the workbook already holds `TEST_` or `EX_` modules, it
says so and leaves them; build a production workbook again from the template.
`-Mode` cannot be combined with `-SourceDir`.

Before it opens the workbook, the script lists the worksheet and
`ThisWorkbook` modules in the source and asks whether to replace their code as
well. Answer N to keep the sheet code of the workbook; the standard modules are
imported either way. `-IncludeSheets` or `-SkipSheets` gives the answer in
advance. The repository's sheet modules hold only `Option Explicit`, so
skipping them changes no behaviour.

To test modules received outside the repository, for example downloaded during
a review, without copying them into the working tree, drag the workbook onto
`tools/Sync-IRRBB-VBA-Downloads.cmd`. It asks for the folder (Enter for
`%USERPROFILE%\Downloads`) and runs the synchronizer with `-SourceDir`:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\Sync-IRRBB-VBA.ps1 -WorkbookPath "C:\path\to\IRRBB.xlsm" -SourceDir "$env:USERPROFILE\Downloads"
```

In that mode only the modules found in the folder are imported; every other
module in the workbook is left as it is. A file is identified only by the
`Attribute VB_Name` it declares, never by its file name: a file with no such
line, a malformed one, several ones or a name that is not a valid VBA name is
listed as ignored with the reason. A file is taken when its `VB_Name` is a
component of the repository source with the same extension, or when it is a
standard module whose name starts with `CORE_`, `IRRBB_`, `WB_`, `TEST_` or
`EX_`, such as a test module added by the pull request under test; that one is
listed apart as a new module. When several files carry the same `VB_Name`,
such as `CORE_Import.bas` and `CORE_Import (1).bas`, the newest is taken. The
script lists the chosen and the ignored files and asks for confirmation
(`-Force` skips it; answering no exits with code 2 and leaves the workbook
unchanged). The chosen files are copied to a temporary folder with CRLF line
endings, imported from there, and the copy is deleted. The repository files are
never changed. A later run of `Sync-IRRBB-VBA.cmd` brings the workbook back to
the repository source.

Standard modules are replaced from source; `ThisWorkbook` and the worksheet
document modules keep their workbook objects and only their code text is
replaced, read as cp1252. The script disables Excel events, macro execution,
external-link updates and automatic calculation while synchronizing, saves the
workbook with the calculation mode and before-save setting it was opened with,
checks that the sheet names and CodeNames are unchanged, and makes a timestamped
safety copy of the workbook unless `-NoBackup` is given. It deletes that copy
after a successful save and keeps it on failure; it is not a rollback archive.

`check_source.py` keeps the synchronizer complete: it fails when the
synchronizer or the template is not tracked, when a VBA component lies outside
the directories in `$sourceRoots`, when one of those directories is missing,
and when the document modules in `src/workbook/` and the template's sheets (by
CodeName) do not match one to one. A new source directory must be added to
`$sourceRoots`; a new worksheet needs both its sheet in the template and its
document module.

Because it changes code only, a change to a sheet layout does not reach an
existing workbook; such a workbook is rebuilt from the current template.

It requires Windows desktop Excel and Excel's **Trust access to the VBA project
object model** setting (see [`SECURITY.md`](../SECURITY.md)). The workbook's VBA
project must not be password-locked. The development-workbook builder above
remains the route that needs neither.

## Source round trip

`Export-IRRBB-VBA.ps1` and `compare_vba_export.py` check that a workbook holds
exactly the source of one commit. The script opens a built workbook read-only
in Excel, with macros disabled, and exports every VBA component to a new
folder, with `export-manifest.json` recording the workbook's SHA-256 and every
exported file's. The comparison checks that folder against the source of one
commit for a production or a development workbook, prints the template and
workbook digests, and with `--workbook` checks the workbook against the
manifest (`VERIFIED`, `FAIL`, or `UNVERIFIED` without a manifest):

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\Export-IRRBB-VBA.ps1 -WorkbookPath "C:\path\to\IRRBB.xlsm" -OutDir "C:\path\to\export"
python tools/compare_vba_export.py --export "C:\path\to\export" --mode development --candidate-sha <sha> --workbook "C:\path\to\IRRBB.xlsm"
```

It ignores line endings, trailing blank lines, the letter case of code outside
strings and comments (the VBE re-cases identifiers on import; such lines are
counted) and the attributes of document modules, which Excel writes. Everything
else must match. The export needs the same Trust access as the synchronizer;
the comparison never runs Excel.

Label reconciliation scripts and workflows apply the label policy and run from the default branch `main`; see [`docs/LABELS.md`](../docs/LABELS.md).

## Code scanning, links and pull-request milestones

- **CodeQL** (`.github/workflows/codeql.yml`) scans the workflows (`actions`),
  the label scripts (`javascript-typescript`) and the Python tooling with the `security-extended` queries on pushes and
  pull requests to `main` and `release/**`, weekly and on demand. VBA is not a
  CodeQL language. Pushes and scheduled runs upload results to code scanning;
  a pull request is analysed without uploading, and its SARIF files are kept
  as a run artifact for 30 days. This workflow replaces GitHub's default
  CodeQL setup, which must stay off: the two cannot run together. On a private
  repository the scan needs GitHub Code Security and the repository variable
  `CODE_SCANNING` set to `enabled`; without them it is skipped, not failed.
- **External links** (`.github/workflows/link-check.yml`) runs
  `python tools/check_links.py --external` every Monday and on demand. It
  requests each external https link once, over HTTPS only; links into this
  repository are skipped, because the offline check covers them as relative
  links. Rate limits and network errors are reported as warnings; a 404, and a
  redirect to a URL that is not HTTPS, as errors.
- **Pull-request milestones** (`.github/workflows/pr-milestones.yml` with
  `.github/scripts/pr-milestones.py`) removes any milestone set on a pull
  request, when one is opened, reopened, edited or given a milestone, and
  daily across all pull requests. It reads each item back and refuses to
  change anything that is not a pull request, so issue milestones are never
  touched (`test_pr_milestones.py`).

<a id="repository-settings"></a>

## Repository settings

`docs/repository-settings.json` is the baseline of the general settings, the
security features, the deployment branches of the `analytics` environment
(`main` only, [`docs/TRAFFIC.md`](../docs/TRAFFIC.md#credentials-and-access)),
the CodeQL default setup (`not-configured`, see above) and the rulesets for
`main`, `release/**` (one for pull requests and checks, one that forbids
deletion), the release tags `v*` and `traffic-history`; none can be bypassed,
by admins included.
`setup_repository.py` compares the saved settings with it through an authenticated GitHub CLI and refuses to run without admin
rights on the repository:

```shell
python tools/setup_repository.py --repo danielep71/K-IRRBB-NMDs            # preview; never writes
python tools/setup_repository.py --repo danielep71/K-IRRBB-NMDs --apply    # write, then read back
python tools/setup_repository.py --repo danielep71/K-IRRBB-NMDs --check    # read-only; exit 1 on drift
```

Every run writes a dated evidence file,
`test-results/repository-settings-<mode>-<time>.md` unless `--report` names
another. `--apply` stops before writing when a named ruleset differs, until
`--update-existing` allows it to replace that whole ruleset. A ruleset the
baseline does not name is reported and never changed or deleted. For an
environment, `--apply` sets the deployment branch policy, keeping its wait
timer and required reviewers, then adds or deletes branch policies until only
the baseline branches remain; it never reads or changes environment secrets or
variables. Dependabot
security updates stay off: their pull requests target `main`, which takes work
only through the release branch, and the weekly version updates on the
release branch cover the pinned actions and tools. `test_setup_repository.py`
runs the script against a fake API.

## Issue triage

`check_issues.py` audits every open issue against the rules in
[`docs/GOVERNANCE.md`](../docs/GOVERNANCE.md#issue-triage) and never changes
one. `.github/workflows/issue-triage.yml` runs it on issue and milestone events
and on demand with read-only permissions, writes the findings to the job
summary and fails when there are any. Locally, export the issues as REST JSON
and run:

```shell
python tools/check_issues.py --input issues.json --repo-name danielep71/K-IRRBB-NMDs
```

## Dependency updates

`.github/dependabot.yml` asks Dependabot for weekly version updates of the
GitHub Actions pinned in the workflows and of the hash-locked quality tools in
`tools/requirements-quality-ci.txt` (Ruff, mypy and their dependencies, in one
grouped pull request). Every `github/codeql-action` step also moves in one
grouped pull request, because CodeQL refuses to run when its `init` and
`analyze` steps differ in version. They are opened against the active release branch
(`target-branch`) and labelled `ci` and `P3`. Dependabot reads this file from
the default branch, so it takes effect once the file reaches `main`. Update
every `target-branch` whenever a new release branch is opened
([`RELEASING.md`](../RELEASING.md)). An update of the quality tools may add
hashes for other platforms; the file must keep a hash for CPython on Linux
x86_64, which CI installs.

Review each update pull request like any other:

1. Check the new version's release notes and the commit it pins. The pin must
   stay a full 40-character SHA with the version as a trailing comment.
2. Let **Repository integrity** run and confirm it is green.
3. Squash-merge it. Nothing is merged or approved automatically, and
   Dependabot alerts are never dismissed automatically.

## Quality gates

CI runs one exact Python version, 3.13.18, in every workflow, and the minimum
supported version, 3.11, in the non-required **Minimum Python (3.11)** job of
Static repository checks, which runs `tools/check.py --ci`.
`tools/test_workflow_pins.py` fails when a version is not exact or differs
between workflows, or when a job runs Python without setting it up.

After the portable gates above, the **Repository integrity** job installs Ruff
and mypy from the hash-locked `tools/requirements-quality-ci.txt` (CPython 3.13
on Linux x86_64, wheels only) and runs:

- `ruff check tools` with the settings in `pyproject.toml`, plus a probe that
  proves the complexity ceiling of 15 is active;
- `mypy --strict` with `warn_unreachable` on every module in `tools/`, tests
  included, plus probes that prove strict mode and unreachable-code warnings
  are active;
- actionlint 1.7.12 on every workflow, downloaded and checked against its
  SHA-256.

The quality steps run even when an earlier step failed, so one run reports
every problem. Dependabot proposes updates of Ruff, mypy and their
dependencies; actionlint's version and checksum in the workflow `env` are
changed by hand in one reviewed pull request. Locally, with Python 3.13:

```shell
python -m pip install --only-binary=:all: --require-hashes -r tools/requirements-quality-ci.txt
ruff check tools
mypy
```
