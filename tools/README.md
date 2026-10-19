# Repository tooling

The tools are adapted from the shared source-first conventions and portable source gates of VBA-SACCR-Toolkit; they do not perform IRRBB calculations.

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
- changelog/version consistency.

CI runs these checks on GitHub's hosted Linux runner. **VBA compile, Excel workbook execution, econometric calibration and numerical result validation are not performed by these gates.** Record Excel host evidence separately under `docs/EXCEL_EVIDENCE.md`.

`workbook/build_template.py` generates the macro-free template (standard library only). `devbuild/` builds the macro-enabled development workbook used for Excel compile and harness evidence, without Excel automation and without Trust access to the VBA project object model; see [`devbuild/README.md`](devbuild/README.md).

## Local VBA synchronization

`Sync-IRRBB-VBA.ps1` synchronizes the repository VBA source into an existing
local `.xlsm` workbook, such as one saved from the template, without
intentionally changing worksheets, cells, formulas, names, tables or
formatting. It follows the synchronizer of VBA-SACCR-Toolkit.

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

Label reconciliation scripts and workflows mirror the SACCR label policy; merging them into the trusted default branch is needed before automatic label synchronization begins.

Ruff and mypy settings in `pyproject.toml` are optional development targets, not
CI checks or verified compliance. The required CI command remains the portable
standard-library gate runner above.
