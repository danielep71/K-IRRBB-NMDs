# Development workbook build

Assembles a macro-enabled **development** workbook from one checked-out commit, so the source can be compiled and the regression harness run in Excel and recorded per [`docs/EXCEL_EVIDENCE.md`](../../docs/EXCEL_EVIDENCE.md). This builder is the import route chosen for #6: it writes the VBA project from source, deterministically, without enabling *Trust access to the VBA project object model*. It is built from the workbook template of #5 but is not the template, and never the production deliverable.

| Script | Purpose |
| --- | --- |
| `build_dev_workbook.py` | Writes `build/IRRBB_Dev.xlsm` (ignored by Git): the [workbook template](../../docs/WORKBOOK_CONTRACT.md) with a build stamp and evidence steps on `Readme` a **Run tests** button on `Checks`, and **Choose account file**, **Choose rate file** and **Import data** buttons on `Config`; embeds every `.bas` in `src/core`, `src/modules`, `src/workbook`, `tests/modules` and `examples/modules` unchanged (the decay example `EX_DecayExample` compiles but needs a `DecayExample` sheet the template does not have yet, #36), with document modules for `ThisWorkbook` and each sheet. Standard library only. A tracked change, or an untracked file in those folders, stamps the build as not evidence-grade. |
| `vba_project.py` | Writes `vbaProject.bin` from the [MS-OVBA] and [MS-CFB] specifications: compressed module source, `dir` stream, `PROJECT` streams, no p-code cache, so Excel compiles from source on open. Deterministic for a given commit. Read back by an independent decoder in [`tools/test_devbuild.py`](../test_devbuild.py). |
| `fixture_module.py` | Writes `tests/modules/TEST_Fixtures.bas`: the data-contract fixtures and their hand-computed expectations as VBA string literals, so the data cases need no file paths. `--check` compares; `tools/test_vba_harness.py` fails when the committed module is stale. |
| `make_large_panel.py` | Writes a large synthetic panel and market-rate file into `build/large_panel/` (ignored by Git) for timing the import: 50,000 accounts × 16 month ends, about 47 MB, by default. Contract-valid, from a fixed seed. |
| `lo_smoke.py` | Optional early warning: opens the workbook in headless LibreOffice with VBA compatibility, runs `TEST_Harness.RunTests`, plants stale rows, runs again to check they are withdrawn, runs `TEST_RunTestsWithInjectedFailures` (a failing assertion and an unexpected error must show as `FAIL` and `ERROR`, and a normal rerun must withdraw them) and `TEST_RunTestsWithWriteFault`, and prints a JSON summary. LibreOffice is not a supported host; this is never Excel evidence. Needs LibreOffice and `python3-uno`. |

```shell
python tools/devbuild/build_dev_workbook.py
python tools/devbuild/lo_smoke.py            # optional
```

**Timing an import in Excel.** Run `python tools/devbuild/make_large_panel.py`
(it prints the two file paths and the data window). In the development
workbook, enter both paths and the window on `Config`, attest closure
reporting, click **Import data**, and record from `DataSummary` the started and
finished times and the account panel bytes from the lineage, with the Excel
version, bitness and machine. The run reads the file once and computes its
SHA-256 in the same pass; the result decides whether the fingerprint needs the
host's cryptography (#71).

LibreOffice is stricter than Excel in a few places. Valid VBA that it does not run correctly, and that the source therefore avoids: a `vbObjectError + n` constant expression; `CDec`; functions returning a typed array (`As Long()`); arrays of user-defined types; a collection item read through the default member of a module-level object (`c(key)`, use `c.Item(key)`); a single-line `If ... Then ReDim Preserve`.

Changes made in the VBA editor are not read back: export the changed modules by hand to their source folders (LF line endings, `Attribute VB_Name` header kept) and rebuild.

If Excel rejects or strips the embedded project, fall back to the manual import order in [`INSTALLATION.md`](../../INSTALLATION.md#importing-vba-into-excel) on the template `src/workbook/IRRBB_Template.xlsx`, and report the failure in an issue.
