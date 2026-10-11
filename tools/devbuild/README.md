# Development workbook build

Assembles a macro-enabled **development** workbook from one checked-out commit, so the source can be compiled and the regression harness run in Excel and recorded per [`docs/EXCEL_EVIDENCE.md`](../../docs/EXCEL_EVIDENCE.md). This builder is the import route chosen for #6: it writes the VBA project from source, deterministically, without enabling *Trust access to the VBA project object model*. It is built from the workbook template of #5 but is not the template, and never the production deliverable.

| Script | Purpose |
| --- | --- |
| `build_dev_workbook.py` | Writes `build/IRRBB_Dev.xlsm` (ignored by Git): the [workbook template](../../docs/WORKBOOK_CONTRACT.md) with a build stamp and evidence steps on `Readme` and a **Run tests** button on `Checks`; embeds every `.bas` in `src/core`, `src/modules`, `src/workbook`, `tests/modules` and `examples/modules` unchanged, with document modules for `ThisWorkbook` and each sheet. Standard library only. A tracked change, or an untracked file in those folders, stamps the build as not evidence-grade. |
| `vba_project.py` | Writes `vbaProject.bin` from the [MS-OVBA] and [MS-CFB] specifications: compressed module source, `dir` stream, `PROJECT` streams, no p-code cache, so Excel compiles from source on open. Deterministic for a given commit. Read back by an independent decoder in [`tools/test_devbuild.py`](../test_devbuild.py). |
| `lo_smoke.py` | Optional early warning: opens the workbook in headless LibreOffice with VBA compatibility, runs `TEST_Harness.RunTests`, plants stale rows, runs again to check they are withdrawn, and prints a JSON summary. LibreOffice is not a supported host; this is never Excel evidence. Needs LibreOffice and `python3-uno`. |

```shell
python tools/devbuild/build_dev_workbook.py
python tools/devbuild/lo_smoke.py            # optional
```

Changes made in the VBA editor are not read back: export the changed modules by hand to their source folders (LF line endings, `Attribute VB_Name` header kept) and rebuild.

If Excel rejects or strips the embedded project, fall back to the manual import order in [`INSTALLATION.md`](../../INSTALLATION.md#importing-vba-into-excel) on the template `src/workbook/IRRBB_Template.xlsx`, and report the failure in an issue.
