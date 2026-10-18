#!/usr/bin/env python3
"""Generate the macro-free workbook template src/workbook/IRRBB_Template.xlsx (issue #5).

Usage: python tools/workbook/build_template.py [--check]

The sheet contract in docs/WORKBOOK_CONTRACT.md is written here as data (SHEETS,
NAMES, VALIDATIONS) and turned into SpreadsheetML with the standard library
only, so the template is reproducible from source: every package member is the
same on every run, and tools/test_workbook_template.py fails if the committed
file differs from what this script produces. --check compares without writing.

The package holds no VBA project, no document properties (author, company,
timestamps), no external links and no cell formulas. The development-workbook
builder (tools/devbuild/) calls package() with a build stamp and adds the VBA
project to the same sheets.
"""
from __future__ import annotations

import argparse
import io
import sys
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from xml.sax.saxutils import escape

ROOT = Path(__file__).resolve().parents[2]
TEMPLATE = ROOT / "src" / "workbook" / "IRRBB_Template.xlsx"

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
ZIP_TIME = (1980, 1, 1, 0, 0, 0)

# Status written to every output sheet's status cell (B3) before any run. Only a
# status beginning "COMPLETE" marks the sheet's results as current.
NOT_IMPLEMENTED = "NOT RUN - not implemented"

# Cell styles (index into cellXfs below).
TITLE, PURPOSE, LABEL, HEADER, INPUT, INPUT_DATE, STATUS, STATUS_TIME, TEXT = range(1, 10)


@dataclass
class Sheet:
    name: str
    code_name: str
    role: str
    owner: str
    status: str
    title: str
    purpose: str
    cells: dict[str, tuple[str, int]] = field(default_factory=dict)
    widths: dict[str, float] = field(default_factory=dict)
    freeze_row: int = 0


SHEETS = [
    Sheet("Readme", "shReadme", "Documentation", "Builder", "Implemented: sheet map and build stamp",
          "IRRBB Toolkit", "Workbook for IRRBB behavioral models. This sheet maps every sheet and records the build."),
    Sheet("Config", "shConfig", "Input", "User", "Implemented: run settings read by the data importer (#71)",
          "Run settings", "Data files, data window, rate series and attestation for one run (DATA_CONTRACT.md)."),
    Sheet("Segments", "shSegments", "Input", "User",
          "Implemented: segment selection read by the data importer (#71)",
          "Segments", "Segments of the data contract; choose which ones a run includes."),
    Sheet("DataSummary", "shDataSummary", "Output", "VBA",
          "Implemented: import totals, outcomes, findings, warnings and lineage (#71)",
          "Data summary", "Import counts, reconciliation, warnings and lineage of the loaded data."),
    Sheet("ModelParams", "shModelParams", "Output", "VBA", "Placeholder: parameter sets (v0.3.0 to v0.5.0)",
          "Model parameters", "Versioned parameter sets of the decay, rate and stable-amount models."),
    Sheet("Calibration", "shCalibration", "Output", "VBA", "Placeholder: estimation diagnostics (v0.3.0 to v0.5.0)",
          "Calibration", "Estimation results and diagnostics of each model."),
    Sheet("Backtesting", "shBacktesting", "Output", "VBA", "Placeholder: out-of-sample backtests (v0.6.0)",
          "Backtesting", "Frozen-coefficient forecasts and their scores."),
    Sheet("Scenarios", "shScenarios", "Output", "VBA", "Placeholder: scenario runs (v0.7.0)",
          "Scenarios", "Model responses under rate scenarios on a frozen parameter set."),
    Sheet("Results", "shResults", "Output", "VBA", "Placeholder: run results and export view (v0.8.0)",
          "Results", "Results of the latest run for review and export."),
    Sheet("Checks", "shChecks", "Output", "VBA",
          "Implemented: regression harness output in the development workbook (#8)",
          "Checks", "Regression harness (TEST_Harness.RunTests) output: run status, cases and assertions."),
]

# Workbook-scoped defined names: VBA addresses areas by name, never by literal address.
NAMES = {
    "Readme_SourceCommit": ("Readme", "$B$5"),
    "Config_AccountsFile": ("Config", "$B$6"),
    "Config_RatesFile": ("Config", "$B$7"),
    "Config_WindowStart": ("Config", "$B$8"),
    "Config_WindowEnd": ("Config", "$B$9"),
    "Config_RateSeries": ("Config", "$B$10"),
    "Config_ClosureAttested": ("Config", "$B$11"),
    "Segments_Codes": ("Segments", "$A$6:$A$8"),
    "Segments_Include": ("Segments", "$C$6:$C$8"),
    "DataSummary_Info": ("DataSummary", "$B$4:$B$12"),
    "DataSummary_Totals": ("DataSummary", "$A$14:$E$20013"),
    "DataSummary_Findings": ("DataSummary", "$G$14:$J$20013"),
    "DataSummary_Warnings": ("DataSummary", "$L$14:$N$20013"),
    "DataSummary_Lineage": ("DataSummary", "$P$14:$Q$63"),
    "DataSummary_Outcomes": ("DataSummary", "$S$14:$AB$20013"),
    "Checks_Cases": ("Checks", "$A$14:$F$514"),
    "Checks_Assertions": ("Checks", "$H$14:$N$20014"),
    **{f"{s.name}_Status": (s.name, "$B$3") for s in SHEETS if s.role == "Output"},
}

SEGMENTS = [("RET_TX", "Retail, transactional"), ("RET_NTX", "Retail, non-transactional"),
            ("WHS_NFC", "Wholesale, non-financial corporate")]
YES_NO = "Yes,No"
FIRST_SERIAL, LAST_SERIAL = 1, 2958465  # 1900-01-01 and 9999-12-31 in the 1900 date system

# (sheet, range, kind, list or bounds, error message). Excel's validation helps
# the user; it is bypassed by pasting, so VBA validates every input it reads.
VALIDATIONS = [
    ("Config", "B8:B9", "date", (FIRST_SERIAL, LAST_SERIAL), "Enter a month-end date from 1900 to 9999."),
    ("Config", "B11", "list", YES_NO, "Choose Yes or No."),
    ("Segments", "C6:C8", "list", YES_NO, "Choose Yes or No."),
]

BUILD_LABELS = ["Source commit", "Built (UTC)", "Builder", "Components", "References"]
TEMPLATE_STAMP = ["None: template without a VBA project", "Not built from source", "tools/workbook/build_template.py",
                  "None", "Not applicable"]


def _layout() -> None:
    sheets = {s.name: s for s in SHEETS}
    for s in SHEETS:
        s.cells["A1"] = (s.title, TITLE)
        s.cells["A2"] = (s.purpose, PURPOSE)
        if s.role == "Output":
            s.cells["A3"] = ("Status", LABEL)
            s.cells["B3"] = (NOT_IMPLEMENTED, STATUS)
            s.widths = {"A": 18, "B": 40}
        if s.status.startswith("Placeholder"):
            s.cells["A5"] = ("Reserved: no function writes this sheet yet. The result layout from row 13 is defined "
                             "when the function is built (docs/WORKBOOK_CONTRACT.md).", TEXT)

    readme = sheets["Readme"]
    readme.cells["A4"] = ("Build", LABEL)
    readme.cells["A11"] = ("Sheets", LABEL)
    for col, text in zip("ABCDE", ("Sheet", "CodeName", "Role", "Owner", "Status")):
        readme.cells[f"{col}12"] = (text, HEADER)
    for row, s in enumerate(SHEETS, 13):
        for col, text in zip("ABCDE", (s.name, s.code_name, s.role, s.owner, s.status)):
            readme.cells[f"{col}{row}"] = (text, TEXT)
    readme.cells["A24"] = ("A sheet name is not evidence that a function exists: the Status column says what is "
                           "implemented. Contract: docs/WORKBOOK_CONTRACT.md.", TEXT)
    readme.widths = {"A": 18, "B": 18, "C": 15, "D": 10, "E": 72}

    config = sheets["Config"]
    for col, text in zip("ABCD", ("Setting", "Value", "Format", "Notes")):
        config.cells[f"{col}5"] = (text, HEADER)
    settings = [
        ("Account panel file", INPUT, "Path", "Account panel CSV; set by the file picker. Read only."),
        ("Market rates file", INPUT, "Path", "Market rate CSV; set by the file picker. Read only."),
        ("Data window start", INPUT_DATE, "Month end", "First month end of the data window."),
        ("Data window end", INPUT_DATE, "Month end", "Last month end of the data window."),
        ("Rate series", INPUT, "rate_id", "Market rate series used by the rate model."),
        ("Timely closure reporting attested", INPUT, "Yes / No",
         "Closures are reported by the month end in which they occur (DATA_CONTRACT.md). No: the import stops."),
    ]
    for row, (label, style, fmt, note) in enumerate(settings, 6):
        config.cells[f"A{row}"] = (label, LABEL)
        config.cells[f"B{row}"] = ("", style)
        config.cells[f"C{row}"] = (fmt, TEXT)
        config.cells[f"D{row}"] = (note, TEXT)
    config.widths = {"A": 32, "B": 40, "C": 12, "D": 90}

    segments = sheets["Segments"]
    for col, text in zip("ABC", ("Code", "Description", "Include")):
        segments.cells[f"{col}5"] = (text, HEADER)
    for row, (code, description) in enumerate(SEGMENTS, 6):
        segments.cells[f"A{row}"] = (code, TEXT)
        segments.cells[f"B{row}"] = (description, TEXT)
        segments.cells[f"C{row}"] = ("Yes", INPUT)
    segments.widths = {"A": 12, "B": 40, "C": 10}

    summary = sheets["DataSummary"]
    for row, label in enumerate(["Status", "Started", "Finished", "Account panel file", "Market-rate file",
                                 "Data window", "Rate series", "Closures attested", "Primary error",
                                 "Cleanup error"], 3):
        summary.cells[f"A{row}"] = (label, LABEL)
        if row > 3:
            summary.cells[f"B{row}"] = ("", STATUS_TIME if row in (4, 5) else STATUS)
    for col, text in zip("ABCDE", ("Date", "Segment", "Currency", "Accounts", "Balance")):
        summary.cells[f"{col}13"] = (text, HEADER)
    for col, text in zip("GHIJ", ("File", "Line", "Code", "Date")):
        summary.cells[f"{col}13"] = (text, HEADER)
    for col, text in zip("LMN", ("Code", "Account", "Date")):
        summary.cells[f"{col}13"] = (text, HEADER)
    for col, text in zip("PQ", ("Item", "Value")):
        summary.cells[f"{col}13"] = (text, HEADER)
    outcome_columns = ("S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB")
    for col, text in zip(outcome_columns, ("Month end", "Segment", "Currency", "Observed", "Verified closures", "Gaps",
                                           "Censored", "Censored last balance", "End of sample",
                                           "Close dates masked")):
        summary.cells[f"{col}13"] = (text, HEADER)
    summary.widths = dict(zip("ABCDEFGHIJKLMNOPQR", (26, 40, 10, 10, 16, 3, 22, 8, 7, 12, 3, 7, 14, 12, 3, 30, 22,
                                                     3)))
    summary.widths.update(zip(outcome_columns, (12, 10, 9, 10, 16, 7, 9, 20, 13, 18)))
    summary.freeze_row = 13

    checks = sheets["Checks"]
    checks.cells["B3"] = ("NOT RUN", STATUS)
    for row, label in enumerate(["Status", "Started", "Finished", "Source commit", "Host", "Cases", "Assertions",
                                 "Primary error", "Cleanup error"], 3):
        checks.cells[f"A{row}"] = (label, LABEL)
        if row > 3:
            checks.cells[f"B{row}"] = ("", STATUS_TIME if row in (4, 5) else STATUS)
    for col, text in zip("ABCDEF", ("Case", "Outcome", "Assertions", "Passed", "Failed", "Note")):
        checks.cells[f"{col}13"] = (text, HEADER)
    for col, text in zip("HIJKLMN", ("Case", "Assertion", "Expected", "Actual", "AbsDiff", "Tolerance", "Outcome")):
        checks.cells[f"{col}13"] = (text, HEADER)
    checks.widths = dict(zip("ABCDEFGHIJKLMN", (16, 40, 11, 9, 9, 60, 3, 14, 34, 18, 18, 10, 14, 9)))
    checks.freeze_row = 13


_layout()


# ============================================================================ SpreadsheetML
def _cell_key(ref: str) -> tuple[int, int]:
    letters = ref.rstrip("0123456789")
    col = 0
    for ch in letters:
        col = col * 26 + ord(ch) - 64
    return int(ref[len(letters):]), col


def _column_letters(col: int) -> str:
    letters = ""
    while col:
        col, rest = divmod(col - 1, 26)
        letters = chr(65 + rest) + letters
    return letters


def _xml(body: str) -> bytes:
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + body).encode("utf-8")


def _status_rules(priority: int) -> tuple[str, int]:
    rules = (f'<conditionalFormatting sqref="B3">'
             f'<cfRule type="expression" dxfId="1" priority="{priority}"><formula>'
             f'OR(LEFT($B$3,6)="FAILED",$B$3="COMPLETE - FAILURES")</formula></cfRule>'
             f'<cfRule type="expression" dxfId="0" priority="{priority + 1}"><formula>'
             f'AND(LEFT($B$3,8)="COMPLETE",$B$3&lt;&gt;"COMPLETE - FAILURES")</formula></cfRule>'
             f'<cfRule type="expression" dxfId="2" priority="{priority + 2}"><formula>'
             f'AND(LEFT($B$3,6)&lt;&gt;"FAILED",LEFT($B$3,8)&lt;&gt;"COMPLETE")</formula></cfRule>'
             f'</conditionalFormatting>')
    return rules, priority + 3


def _sheet_xml(sheet: Sheet, strings: dict[str, int], selected: bool) -> bytes:
    rows: dict[int, list[str]] = {}
    for ref in sorted(sheet.cells, key=_cell_key):
        text, style = sheet.cells[ref]
        row = _cell_key(ref)[0]
        if text:
            index = strings.setdefault(text, len(strings))
            rows.setdefault(row, []).append(f'<c r="{ref}" s="{style}" t="s"><v>{index}</v></c>')
        else:
            rows.setdefault(row, []).append(f'<c r="{ref}" s="{style}"/>')
    last_row = max(_cell_key(ref)[0] for ref in sheet.cells)
    last_col = max(_cell_key(ref)[1] for ref in sheet.cells)
    dimension = f"A1:{_column_letters(last_col)}{last_row}"

    tab = ' tabSelected="1"' if selected else ""
    view = f'<sheetView showGridLines="0"{tab} workbookViewId="0">'
    if sheet.freeze_row:
        top = f"A{sheet.freeze_row + 1}"
        view += (f'<pane ySplit="{sheet.freeze_row}" topLeftCell="{top}" activePane="bottomLeft" state="frozen"/>'
                 f'<selection pane="bottomLeft" activeCell="{top}" sqref="{top}"/>')
    view += "</sheetView>"
    cols = "".join(f'<col min="{_cell_key(c + "1")[1]}" max="{_cell_key(c + "1")[1]}" width="{w}" customWidth="1"/>'
                   for c, w in sheet.widths.items())
    data = "".join(f'<row r="{r}">{"".join(cells)}</row>' for r, cells in sorted(rows.items()))

    formatting, priority = "", 1
    if sheet.role == "Output":
        formatting, priority = _status_rules(priority)
    if sheet.name == "Checks":
        for sqref, value, dxf in (("B14:B514", "PASS", 0), ("B14:B514", "FAIL", 1), ("B14:B514", "ERROR", 1),
                                  ("N14:N20014", "FAIL", 1)):
            formatting += (f'<conditionalFormatting sqref="{sqref}"><cfRule type="cellIs" dxfId="{dxf}" '
                           f'priority="{priority}" operator="equal"><formula>"{value}"</formula></cfRule>'
                           f'</conditionalFormatting>')
            priority += 1

    checks = [v for v in VALIDATIONS if v[0] == sheet.name]
    validations = ""
    if checks:
        validations = f'<dataValidations count="{len(checks)}">'
        for _, sqref, kind, rule, message in checks:
            if kind == "list":
                body = f'type="list" allowBlank="1" showErrorMessage="1" error="{escape(message)}" sqref="{sqref}">'
                body += f'<formula1>"{rule}"</formula1>'
            else:
                body = (f'type="date" operator="between" allowBlank="1" showErrorMessage="1" '
                        f'error="{escape(message)}" sqref="{sqref}"><formula1>{rule[0]}</formula1>'
                        f'<formula2>{rule[1]}</formula2>')
            validations += f"<dataValidation {body}</dataValidation>"
        validations += "</dataValidations>"

    return _xml(f'<worksheet xmlns="{MAIN_NS}" xmlns:r="{REL_NS}"><sheetPr codeName="{sheet.code_name}"/>'
                f'<dimension ref="{dimension}"/><sheetViews>{view}</sheetViews>'
                f'<sheetFormatPr defaultRowHeight="15"/><cols>{cols}</cols><sheetData>{data}</sheetData>'
                f'{formatting}{validations}'
                '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
                '</worksheet>')


STYLES = _xml(
    f'<styleSheet xmlns="{MAIN_NS}">'
    '<numFmts count="2"><numFmt numFmtId="164" formatCode="yyyy-mm-dd"/>'
    '<numFmt numFmtId="165" formatCode="yyyy-mm-dd hh:mm:ss"/></numFmts>'
    '<fonts count="5">'
    '<font><sz val="11"/><name val="Calibri"/><family val="2"/></font>'
    '<font><b/><sz val="11"/><name val="Calibri"/><family val="2"/></font>'
    '<font><b/><sz val="14"/><color rgb="FF1F3A5F"/><name val="Calibri"/><family val="2"/></font>'
    '<font><i/><sz val="11"/><color rgb="FF5B6573"/><name val="Calibri"/><family val="2"/></font>'
    '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/><family val="2"/></font>'
    '</fonts>'
    '<fills count="5"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>'
    '<fill><patternFill patternType="solid"><fgColor rgb="FF1F3A5F"/><bgColor indexed="64"/></patternFill></fill>'
    '<fill><patternFill patternType="solid"><fgColor rgb="FFF2F4F7"/><bgColor indexed="64"/></patternFill></fill>'
    '<fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/><bgColor indexed="64"/></patternFill></fill>'
    '</fills>'
    '<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>'
    '<border><left style="thin"><color rgb="FFC8CED6"/></left><right style="thin"><color rgb="FFC8CED6"/></right>'
    '<top style="thin"><color rgb="FFC8CED6"/></top><bottom style="thin"><color rgb="FFC8CED6"/></bottom>'
    '<diagonal/></border></borders>'
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
    '<cellXfs count="10">'
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
    '<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
    '<xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
    '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
    '<xf numFmtId="0" fontId="4" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" '
    'applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>'
    '<xf numFmtId="49" fontId="0" fillId="4" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" '
    'applyBorder="1"/>'
    '<xf numFmtId="164" fontId="0" fillId="4" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" '
    'applyBorder="1"/>'
    '<xf numFmtId="0" fontId="0" fillId="3" borderId="1" xfId="0" applyFill="1" applyBorder="1"/>'
    '<xf numFmtId="165" fontId="0" fillId="3" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" '
    'applyBorder="1"/>'
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1">'
    '<alignment vertical="top"/></xf>'
    '</cellXfs>'
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
    '<dxfs count="3">'
    '<dxf><fill><patternFill><bgColor rgb="FFC6EFCE"/></patternFill></fill></dxf>'
    '<dxf><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf>'
    '<dxf><fill><patternFill><bgColor rgb="FFFFEB9C"/></patternFill></fill></dxf>'
    '</dxfs></styleSheet>')


def package(stamp: list[str] | None = None, notes: list[str] | None = None) -> bytes:
    """The template package; a build stamp (BUILD_LABELS values) and notes give the development variant."""
    readme = SHEETS[0]
    saved = dict(readme.cells)
    for row, (label, value) in enumerate(zip(BUILD_LABELS, stamp or TEMPLATE_STAMP), 5):
        readme.cells[f"A{row}"] = (label, LABEL)
        readme.cells[f"B{row}"] = (value, TEXT)
    for row, text in enumerate(notes or [], 26):
        readme.cells[f"A{row}"] = (text, LABEL if row == 26 else TEXT)

    strings: dict[str, int] = {}
    members: list[tuple[str, bytes]] = []
    try:
        sheets_xml = [_sheet_xml(s, strings, i == 0) for i, s in enumerate(SHEETS)]
    finally:
        readme.cells = saved

    overrides = "".join(f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.'
                        'openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
                        for i in range(1, len(SHEETS) + 1))
    members.append(("[Content_Types].xml", _xml(
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.'
        'spreadsheetml.sheet.main+xml"/>' + overrides +
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.'
        'spreadsheetml.styles+xml"/>'
        '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.'
        'spreadsheetml.sharedStrings+xml"/></Types>')))
    members.append(("_rels/.rels", _xml(
        f'<Relationships xmlns="{PKG_REL_NS}"><Relationship Id="rId1" Type="{REL_NS}/officeDocument" '
        'Target="xl/workbook.xml"/></Relationships>')))

    sheet_entries = "".join(f'<sheet name="{s.name}" sheetId="{i}" r:id="rId{i}"/>' for i, s in enumerate(SHEETS, 1))
    defined = "".join(f'<definedName name="{name}">{sheet}!{ref}</definedName>'
                      for name, (sheet, ref) in sorted(NAMES.items(), key=lambda item: item[0].upper()))
    members.append(("xl/workbook.xml", _xml(
        f'<workbook xmlns="{MAIN_NS}" xmlns:r="{REL_NS}"><workbookPr codeName="ThisWorkbook"/>'
        f'<bookViews><workbookView activeTab="0"/></bookViews><sheets>{sheet_entries}</sheets>'
        f'<definedNames>{defined}</definedNames><calcPr calcId="191029"/></workbook>')))
    n = len(SHEETS)
    rels = "".join(f'<Relationship Id="rId{i}" Type="{REL_NS}/worksheet" Target="worksheets/sheet{i}.xml"/>'
                   for i in range(1, n + 1))
    rels += (f'<Relationship Id="rId{n + 1}" Type="{REL_NS}/styles" Target="styles.xml"/>'
             f'<Relationship Id="rId{n + 2}" Type="{REL_NS}/sharedStrings" Target="sharedStrings.xml"/>')
    members.append(("xl/_rels/workbook.xml.rels", _xml(f'<Relationships xmlns="{PKG_REL_NS}">{rels}</Relationships>')))
    members.append(("xl/styles.xml", STYLES))
    items = "".join(f'<si><t xml:space="preserve">{escape(text)}</t></si>' for text in strings)
    references = sum(xml.count(b't="s"') for xml in sheets_xml)  # count: every reference; uniqueCount: entries
    members.append(("xl/sharedStrings.xml", _xml(
        f'<sst xmlns="{MAIN_NS}" count="{references}" uniqueCount="{len(strings)}">{items}</sst>')))
    for i, xml in enumerate(sheets_xml, 1):
        members.append((f"xl/worksheets/sheet{i}.xml", xml))

    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, data in members:
            info = zipfile.ZipInfo(name, ZIP_TIME)
            info.compress_type = zipfile.ZIP_DEFLATED
            info.external_attr = 0o644 << 16
            archive.writestr(info, data)
    return out.getvalue()


def members(blob: bytes) -> list[tuple[str, bytes]]:
    """Package members in order, decompressed: what must match, independent of the zlib build."""
    with zipfile.ZipFile(io.BytesIO(blob)) as archive:
        return [(name, archive.read(name)) for name in archive.namelist()]


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--check", action="store_true", help="compare with the committed template; write nothing")
    args = parser.parse_args()
    blob = package()
    if args.check:
        same = TEMPLATE.is_file() and members(TEMPLATE.read_bytes()) == members(blob)
        print("template matches its source" if same else "template differs from its source: rebuild it")
        return 0 if same else 1
    TEMPLATE.parent.mkdir(parents=True, exist_ok=True)
    TEMPLATE.write_bytes(blob)
    print("wrote", TEMPLATE.relative_to(ROOT))
    return 0


if __name__ == "__main__":
    sys.exit(main())
