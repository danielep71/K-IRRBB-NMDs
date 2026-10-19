<div align="center">

# 📗 Workbook Contract

### Sheets, CodeNames, named areas and result status of the IRRBB workbook

[![Template](https://img.shields.io/badge/Template-macro--free_xlsx-217346?style=flat-square)](#template)
[![Source](https://img.shields.io/badge/Built-from_source-1d76db?style=flat-square)](#build)
[![Status](https://img.shields.io/badge/Excel-not_yet_verified-d97706?style=flat-square)](#evidence)

<br>

**Names, not addresses · Values, not formulas · Only COMPLETE is current**

</div>

---

This document is authoritative for **the layout of the workbook: which sheets
exist, who writes each area, how VBA finds it, and when a result counts as
current**. The template is
[`src/workbook/IRRBB_Template.xlsx`](../src/workbook/IRRBB_Template.xlsx)
(issue #5). It is generated from
[`tools/workbook/build_template.py`](../tools/workbook/build_template.py), which
holds this contract as data; `tools/test_workbook_template.py` fails when the
tables below, the generator or the committed template disagree.

<a id="sheets"></a>

## 📑 Sheets

A sheet name is not evidence that a function exists. **Status** says what is
implemented; a placeholder reserves a sheet for the milestone shown and holds
no function.

| Sheet | CodeName | Role | Owner | Status |
| --- | --- | --- | --- | --- |
| `Readme` | `shReadme` | Documentation | Builder | Implemented: sheet map and build stamp |
| `Config` | `shConfig` | Input | User | Implemented: run settings read by the data importer (#71) |
| `Segments` | `shSegments` | Input | User | Implemented: segment selection read by the data importer (#71) |
| `DataSummary` | `shDataSummary` | Output | VBA | Implemented: import totals, outcomes, findings, warnings and lineage (#71) |
| `ModelParams` | `shModelParams` | Output | VBA | Placeholder: parameter sets (v0.3.0 to v0.5.0) |
| `Calibration` | `shCalibration` | Output | VBA | Placeholder: estimation diagnostics (v0.3.0 to v0.5.0) |
| `Backtesting` | `shBacktesting` | Output | VBA | Placeholder: out-of-sample backtests (v0.6.0) |
| `Scenarios` | `shScenarios` | Output | VBA | Placeholder: scenario runs (v0.7.0) |
| `Results` | `shResults` | Output | VBA | Placeholder: run results and export view (v0.8.0) |
| `Checks` | `shChecks` | Output | VBA | Implemented: regression harness output in the development workbook (#8) |

- **Order** is part of the contract; VBA refers to a sheet by its CodeName,
  never by its tab name or position.
- **Owner** is the only writer. *User* cells are typed by the user (or set by a
  picker VBA opens); *VBA* areas are written only by VBA; *Builder* is the
  build stamp written when the workbook is built from source.
- Every sheet starts with the `K-IRRBB` brand strip in `A1` and its title in
  `A2`. The purpose is in `A3` on input/documentation sheets and `D3` on output
  sheets, leaving the output status in `B3`. Presentation merges never cover
  input or runtime output areas. All existing workbook-scoped names and their
  addresses remain unchanged.

<a id="names"></a>

## 🏷️ Named areas

VBA reads and writes through these workbook-scoped names, never through literal
addresses, so a layout change is one change here and in the generator.

| Name | Refers to | Owner | Content |
| --- | --- | --- | --- |
| `Readme_SourceCommit` | `Readme!$B$5` | Builder | Commit of the source the workbook was built from; recorded in the import lineage |
| `Config_AccountsFile` | `Config!$B$6` | User | Account panel CSV path |
| `Config_RatesFile` | `Config!$B$7` | User | Market rate CSV path |
| `Config_WindowStart` | `Config!$B$8` | User | First month end of the data window |
| `Config_WindowEnd` | `Config!$B$9` | User | Last month end of the data window |
| `Config_RateSeries` | `Config!$B$10` | User | `rate_id` of the market rate series used by the rate model |
| `Config_ClosureAttested` | `Config!$B$11` | User | `Yes` when closures are reported by the month end in which they occur |
| `Segments_Codes` | `Segments!$A$6:$A$8` | Contract | The segment codes of the data contract |
| `Segments_Include` | `Segments!$C$6:$C$8` | User | `Yes` or `No` per segment: the run's segment list (lineage) and the segments model samples use; every segment is still imported |
| `DataSummary_Info` | `DataSummary!$B$4:$B$12` | VBA | Run information: start, finish, file names, window, rate series, attestation, primary and cleanup error |
| `DataSummary_Totals` | `DataSummary!$A$14:$E$20013` | VBA | Accounts and balance per date, segment and currency; only for an accepted import |
| `DataSummary_Findings` | `DataSummary!$G$14:$J$20013` | VBA | Errors of the last import: file, line, code, month end (`E12`) |
| `DataSummary_Warnings` | `DataSummary!$L$14:$N$20013` | VBA | Warnings `W01`–`W06`: code, account, month end |
| `DataSummary_Lineage` | `DataSummary!$P$14:$Q$63` | VBA | Lineage items and values |
| `DataSummary_Outcomes` | `DataSummary!$S$14:$AB$20013` | VBA | Per-period lineage: transition outcomes, the last balance of censored disappearances and the masked close dates per month end, segment and currency; only for an accepted import |
| `Checks_Cases` | `Checks!$A$14:$F$514` | VBA | Harness case table |
| `Checks_Assertions` | `Checks!$H$14:$N$20014` | VBA | Harness assertion table |
| `<Sheet>_Status` | `<Sheet>!$B$3` | VBA | Run status of each output sheet: `DataSummary`, `ModelParams`, `Calibration`, `Backtesting`, `Scenarios`, `Results`, `Checks` |

The settings follow [`DATA_CONTRACT.md`](methodology/DATA_CONTRACT.md): the
data window, the segment list and the rate series are part of the
[lineage](methodology/DATA_CONTRACT.md#lineage); the attestation is the
[timely closure reporting](methodology/DATA_CONTRACT.md#look-ahead) the import
cannot check. A blank or `No` attestation stops the import. File paths are
read-only inputs ([`SECURITY.md`](../SECURITY.md#security-scope)).

<a id="status"></a>

## 🚦 Result status

Every output sheet has a status cell (`<Sheet>_Status`, `B3`). **Only a status
beginning `COMPLETE` marks that sheet's results as current.**

| Status | Meaning | Shown |
| --- | --- | --- |
| `NOT RUN - not implemented` | Placeholder; no function writes the sheet | Amber |
| `NOT RUN` | Implemented, not run in this workbook | Amber |
| `RUNNING - previous results withdrawn` | A run started; earlier results were cleared first | Amber |
| `COMPLETE - ...` | The run finished; the results are current | Green |
| `COMPLETE - FAILURES` | The run finished and found failures (harness) | Red |
| `FAILED` | The run stopped; partial results, if any, are not current | Red |
| `FAILED - IMPORT REJECTED` | The data files break the contract; findings are listed, no totals or outcomes | Red |

A run that writes a sheet:

1. sets the status to `RUNNING - previous results withdrawn` and clears the
   sheet's result areas **before** it computes anything;
2. writes results as values;
3. sets `COMPLETE - ...` last, or `FAILED` with the primary error and any
   cleanup error, never one in place of the other.

Invalid input therefore leaves no earlier result looking current. The harness
follows this protocol on `Checks` (#8) and the import on `DataSummary`
(`WB_ImportData`, #71); each later writer adopts it with the function that
writes its sheet. A table longer than its named area is cut at the area's
last row and the cut is stated in the status block.

<a id="rules"></a>

## 📏 Template rules

- **No formulas.** Results are written as values, so recalculation cannot
  change a result after its status was set, and the workbook has no volatile
  or external references. Conditional formatting and data validation are the
  only expressions.
- **Validation is a convenience.** Excel's data validation (dates from 1900 to
  9999 in `Config_WindowStart` and `Config_WindowEnd`; `Yes`/`No` lists in
  `Config_ClosureAttested` and `Segments_Include`) is bypassed by pasting; VBA
  validates every input it reads, including that window dates are month ends.
- **No protection.** Sheet protection is not a security boundary and would
  block VBA writes; ownership is enforced by the code that writes.
- **Sanitized.** No VBA project, no document properties (author, company,
  timestamps), no external links, no personal data, no real data. Enforced by
  `tools/check_source.py` and `tools/test_workbook_template.py`.
- **Result layout.** An output sheet reserves rows 1–12 for its title, purpose
  and status block; result tables start at row 13 (headers) and are defined
  when the function that writes them is built.

### Visual identity

The workbook follows the maintainer's SACCR template visual identity: Aptos,
a black brand strip with gold `K-IRRBB` lettering, a blue title band and
consistent spacing. The reference is `VBA-SACCR-Toolkit`'s
`src/workbook/SACCR_Template.xlsx` at commit
`cf31de00e283f44e7a8b990dc53d7f231a4bf652`; only presentation conventions are
adapted, not its model content or workbook layout.

| Element | Style |
| --- | --- |
| Brand strip | Black `#000000`, gold text `#FFC000`, 14 pt |
| Sheet title | Blue `#265A99`, white bold text, 18 pt |
| Input table headers | Navy `#1F3864`, white bold text, 11 pt |
| Output table headers | Blue `#2E75B6`, white bold text, 11 pt |
| Body / notes | Aptos 10 pt / 9 pt; secondary text `#595959` |
| Editable inputs | Yellow `#FFF2CC`, thin gray `#BFBFBF` borders |
| Tabs | Documentation navy; inputs gold `#BF9000`; outputs blue; regression tests green `#548235` |
| Amounts / counts | `#,##0.00` / `#,##0`; display precision does not change stored values |
| Diagnostic differences / tolerances | Scientific notation `0.00E+00`, retaining visibility of small nonzero values |
| Dates / timestamps | `yyyy-mm-dd` / `yyyy-mm-dd hh:mm:ss`; imported ISO text remains text |

Gridlines are hidden. Output statuses keep the existing amber/green/red rules.
Readme includes a short workflow and the sheet-readiness map. DataSummary
labels totals, findings, warnings, lineage and outcomes; Checks labels cases
and assertions. Placeholder sheets remain explicitly unavailable.

Print setup uses A4 landscape and repeating title/header rows. Readme and
input sheets have bounded print areas; output sheets have no fixed print area
so later VBA results are not omitted. Wide output tables print across pages
at readable widths instead of shrinking all columns onto one page.
Column number formats provide defaults for future values without allocating
thousands of blank result rows. VBA can still assign text formats where the
data contract requires text.

<a id="template"></a>
<a id="build"></a>

## 🛠️ Build

```shell
python tools/workbook/build_template.py           # rewrite the template
python tools/workbook/build_template.py --check   # compare, write nothing
```

The generator uses the standard library only and writes the same package
members on every run (fixed member order and timestamps). Change the layout in
the generator and this document together, regenerate the template, and commit
all three. Never edit the template in Excel and commit the saved file: Excel
adds document properties and rewrites the package.

The development workbook (`tools/devbuild/build_dev_workbook.py`) is this
template plus a build stamp on `Readme`, a **Run tests** button on `Checks` and
the VBA project assembled from source. Development buttons use the same navy
and blue palette and Aptos lettering. The Checks button sits outside the status
and purpose blocks. Development evidence instructions start at Readme row 34,
after the workflow, and never overwrite the sheet map or build-stamp cells.

<a id="evidence"></a>

## 🧾 Evidence

Checked without Excel: the template matches its generator, the tables above
match the generator, the package is sanitized, and the `Checks` names match the
harness constants. LibreOffice, as an early warning only, opens the template
with every sheet, CodeName, name and validation, and runs `WB-LAYOUT-01`.

**Not yet verified in Excel:** that Excel opens the template without a repair
prompt, and that `WB-LAYOUT-01` passes in the development workbook on a
supported host ([`EXCEL_EVIDENCE.md`](EXCEL_EVIDENCE.md)).
