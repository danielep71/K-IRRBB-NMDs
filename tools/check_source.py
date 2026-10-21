#!/usr/bin/env python3
"""Check VBA source, Git storage, the workbook template, synchronizer roots,
CHANGELOG and VERSION without Office."""
from __future__ import annotations

import lzma
import re
import sys
import zipfile
import zlib
from datetime import date
from pathlib import Path
from typing import Any
from xml.etree import ElementTree

from _gatelib import git_bytes, parse_report_args, run_gate, tracked_files
from _vba_lex import physical_lines

VBA_SUFFIXES = {".bas", ".cls", ".frm"}
# The VBE exports in the Windows code page; IRRBB targets Western-European hosts.
VBA_ENCODING = "cp1252"
VB_NAME = re.compile(r'^Attribute VB_Name = "([^"]+)"\s*$', re.M)
OPTION_EXPLICIT = re.compile(r"^[ \t]*Option[ \t]+Explicit[ \t]*(?:'.*)?$", re.M | re.I)
OPTION_PRIVATE = re.compile(r"^[ \t]*Option[ \t]+Private[ \t]+Module[ \t]*(?:'.*)?$", re.M | re.I)
# Role prefix of standard modules in the core, facade and test homes; see
# docs/VBA_HOUSE_STYLE.md.
ROLE_PREFIXES = (("src/core/", "CORE_"), ("src/modules/", "IRRBB_"), ("tests/", "TEST_"))
# Homes for VBA components; see docs/REPOSITORY_STRUCTURE.md.
VBA_HOMES = ("src/core/", "src/modules/", "src/classes/", "src/workbook/", "src/forms/",
             "tests/", "examples/")
FRX_REFERENCE = re.compile(r'"([^"\r\n]+\.frx)":([0-9A-Fa-f]+)')
SEMVER = r"(?:0|[1-9]\d*)\.(?:0|[1-9]\d*)\.(?:0|[1-9]\d*)"
VERSION_HEADING = re.compile(r"^## \[([^\]]+)\](.*)$", re.M)
RELEASE_SUFFIX = re.compile(r" - (\d{4}-\d{2}-\d{2})")
# Committed workbooks (only the template, see docs/REPOSITORY_STRUCTURE.md) carry
# no VBA project and no document properties: author, title, timestamps (#5).
WORKBOOK_SUFFIXES = {".xlsx", ".xlsm", ".xlsb", ".xltx", ".xltm"}
# Excel object model and UI that host-independent src/core code must not use;
# see docs/REPOSITORY_STRUCTURE.md, "Dependency direction". VBA identifiers are
# case-insensitive, so these names are reserved in core even as variable names.
HOST_IDENTIFIER = re.compile(
    r"\b(Application|Excel|ThisWorkbook|ActiveWorkbook|ActiveSheet|ActiveCell|ActiveChart|"
    r"Workbooks?|Worksheets?|Sheets|Range|Cells|Charts?|Shapes|ListObjects?|PivotTables?|"
    r"WorksheetFunction|Evaluate|MsgBox|InputBox|UserForms?)\b", re.I)
# Core cannot activate or attach to external COM objects, even via dynamic ProgIDs.
AUTOMATION_IDENTIFIER = re.compile(r"\b(CreateObject|GetObject)\b", re.I)
FORBIDDEN_PART = re.compile(r"(^|/)vbaProject\.bin$|^docProps/", re.I)
# The synchronizer imports the components listed in tools/vba-projects.json and
# scans the directories in its $sourceRoots list, requiring every component found
# there to be listed; a component outside those directories would escape that
# check. Document modules in src/workbook/ update the code of
# existing sheets, so each needs its sheet in the template. Both files are
# required: without them the checks below could not run.
SYNC_SCRIPT = "tools/Sync-IRRBB-VBA.ps1"
SYNC_START = re.compile(r"^\s*\$sourceRoots\s*=\s*@\(\s*$")
# The one entry form the list may hold; anything else is reported, so that
# a commented-out or unusual entry is never counted as synchronized.
SYNC_ENTRY = re.compile(r'\(Join-Path\s+\$repoFullPath\s+"([^"]+)"\),?')
TEMPLATE = "src/workbook/IRRBB_Template.xlsx"
DOCUMENT_HOME = "src/workbook/"
CODE_NAME = re.compile(r'<(?:sheetPr|workbookPr)\b[^>]*\bcodeName="([^"]+)"')


def index_eol(root: Path) -> dict[str, tuple[str, str]]:
    """Map each tracked path to its (index EOL, attribute) pair from ``git ls-files --eol``."""
    completed = git_bytes(root, "ls-files", "--eol", "-z")
    if completed.returncode:
        raise RuntimeError(completed.stderr.decode("utf-8", errors="replace").strip())
    result: dict[str, tuple[str, str]] = {}
    for record in completed.stdout.split(b"\0"):
        if not record:
            continue
        info, _, path = record.decode("utf-8", errors="surrogateescape").partition("\t")
        fields = info.split()
        result[path] = (fields[0].removeprefix("i/"), " ".join(fields[2:]).removeprefix("attr/"))
    return result


def check_storage(root: Path) -> list[str]:
    """Text must be stored with LF in Git; exported VBA must also check out as CRLF."""
    findings = []
    for path, (eol, attr) in sorted(index_eol(root).items()):
        if eol in {"crlf", "mixed"}:
            findings.append(f"{path}: stored with {eol.upper()} in Git; renormalize to LF")
        # Git reports lone CR line endings as non-text; that is a defect only for declared text.
        elif eol == "-text" and attr.split()[:1] == ["text"]:
            findings.append(f"{path}: declared text but Git classifies the blob as non-text (lone CR?)")
        if Path(path).suffix.lower() in VBA_SUFFIXES and "eol=crlf" not in attr:
            findings.append(f"{path}: .gitattributes must check VBA source out as CRLF")
    return findings


def code_lines(text: str) -> list[tuple[int, str]]:
    """Return (line number, code) with comments and string literals blanked."""
    return [(number, code) for number, code, _ in
            physical_lines(text.split("\n"), mask_strings=True)]


def check_core_host_independence(path: str, text: str) -> list[str]:
    """Core code must not reach the Excel object model or UI; comments and strings may."""
    findings: list[str] = []
    for number, code in code_lines(text):
        if code.lstrip().startswith("Attribute "):
            continue
        for pattern, kind in ((HOST_IDENTIFIER, "Excel host"),
                              (AUTOMATION_IDENTIFIER, "COM automation")):
            findings.extend(f"{path}:{number}: core must not use {kind} identifier {match.group(1)}"
                            for match in pattern.finditer(code))
    return findings


def check_component(root: Path, path: str, tracked: set[str], names: dict[str, str]) -> list[str]:
    try:
        text = (root / path).read_bytes().decode(VBA_ENCODING)
    except UnicodeDecodeError as error:
        return [f"{path}: not decodable as {VBA_ENCODING}: {error}"]
    # Windows checkouts are CRLF; normalize so line anchors behave on every host.
    text = text.replace("\r\n", "\n")
    findings = []
    matches = VB_NAME.findall(text)
    if matches != [Path(path).stem]:
        findings.append(f"{path}: VB_Name must match the filename exactly")
    elif matches[0].casefold() in names:
        findings.append(f"{path}: duplicate component name (also {names[matches[0].casefold()]})")
    else:
        names[matches[0].casefold()] = path
    if not OPTION_EXPLICIT.search(text):
        findings.append(f"{path}: missing Option Explicit")
    if path.startswith("src/core/") and path.lower().endswith(".bas") and not OPTION_PRIVATE.search(text):
        findings.append(f"{path}: core modules must declare Option Private Module")
    if path.startswith("src/core/"):
        findings.extend(check_core_host_independence(path, text))
    for home, prefix in ROLE_PREFIXES:
        if path.startswith(home) and path.lower().endswith(".bas") and not Path(path).stem.startswith(prefix):
            findings.append(f"{path}: standard modules in {home} must be named {prefix}<subject>")
    if path.lower().endswith(".frm"):
        companions = FRX_REFERENCE.findall(text)
        if not companions:
            findings.append(f"{path}: no form resource reference")
        for filename, offset in companions:
            companion = (Path(path).parent / filename).as_posix()
            if Path(filename).name != filename or companion not in tracked:
                findings.append(f"{path}: missing or unsafe form resource {filename}")
            elif (root / companion).stat().st_size <= int(offset, 16):
                findings.append(f"{path}: resource offset is outside {filename}")
    return findings


def check_workbook(root: Path, path: str) -> list[str]:
    """Check core XML parts, archive integrity and sanitization, not Excel execution."""
    binary = Path(path).suffix.lower() == ".xlsb"
    workbook_part = "xl/workbook.bin" if binary else "xl/workbook.xml"
    workbook_rels = "xl/_rels/workbook.bin.rels" if binary else "xl/_rels/workbook.xml.rels"
    required = {"[Content_Types].xml", "_rels/.rels", workbook_part, workbook_rels}
    try:
        with zipfile.ZipFile(root / path) as package:
            parts = package.namelist()
            if len(parts) != len(set(parts)):
                return [f"{path}: duplicate workbook member names are not allowed"]
            missing = sorted(required - set(parts))
            if missing:
                return [f"{path}: missing required workbook parts: {', '.join(missing)}"]
            bad_part = package.testzip()
            if bad_part is not None:
                return [f"{path}: corrupt workbook member {bad_part}"]
            references: list[str] = []
            for name in parts:
                if name.endswith(".rels") or name == "[Content_Types].xml" or name == "xl/workbook.xml":
                    xml = ElementTree.fromstring(package.read(name))
                    references.extend(value for element in xml.iter() for value in element.attrib.values())

    except (OSError, zipfile.BadZipFile, RuntimeError, NotImplementedError,
            zlib.error, lzma.LZMAError, ElementTree.ParseError, LookupError, ValueError) as error:
        return [f"{path}: not a readable workbook package ({error})"]
    findings = [f"{path}: must not contain {part}" for part in parts if FORBIDDEN_PART.search(part)]
    if any("docprops/" in value.casefold() or "vbaproject" in value.casefold()
           for value in references):
        findings.append(f"{path}: package still references document properties or a VBA project")
    return findings


def valid_date(text: str) -> bool:
    try:
        date.fromisoformat(text)
    except ValueError:
        return False
    return True


def check_changelog(root: Path) -> list[str]:
    path = root / "CHANGELOG.md"
    if not path.is_file():
        return ["CHANGELOG.md is missing"]
    text = path.read_text(encoding="utf-8")
    headings = VERSION_HEADING.findall(text)
    findings = []
    if not headings or headings[0] != ("Unreleased", ""):
        findings.append("CHANGELOG.md: the first version heading must be '## [Unreleased]'")
    seen = set()
    for label, suffix in headings:
        if label in seen:
            findings.append(f"CHANGELOG.md: duplicate heading [{label}]")
        seen.add(label)
        if label == "Unreleased":
            continue
        dated = RELEASE_SUFFIX.fullmatch(suffix)
        if not re.fullmatch(SEMVER, label) or not dated:
            findings.append(f"CHANGELOG.md: release heading must be '## [X.Y.Z] - YYYY-MM-DD': [{label}]{suffix}")
        elif not valid_date(dated.group(1)):
            findings.append(f"CHANGELOG.md: [{label}] date {dated.group(1)} is not a calendar date")
        if not re.search(rf"^\[{re.escape(label)}\]: \S+$", text, re.M):
            findings.append(f"CHANGELOG.md: missing link reference for [{label}]")
    releases = [(tuple(int(part) for part in label.split(".")), dated.group(1))
                for label, suffix in headings
                if re.fullmatch(SEMVER, label) and (dated := RELEASE_SUFFIX.fullmatch(suffix))]
    for newer, older in zip(releases, releases[1:]):
        if newer[0] <= older[0] or newer[1] < older[1]:
            findings.append("CHANGELOG.md: releases must be listed newest first, by version and date")
            break
    if "Unreleased" in seen and not re.search(r"^\[Unreleased\]: \S+$", text, re.M):
        findings.append("CHANGELOG.md: missing link reference for [Unreleased]")
    return findings


def required_file(root: Path, tracked: set[str], path: str) -> list[str]:
    """A file the source checks depend on must be tracked and present."""
    if path not in tracked:
        return [f"{path}: required file is not tracked"]
    if not (root / path).is_file():
        return [f"{path}: tracked but missing from the working tree"]
    return []


def strip_comment(line: str, state: str) -> tuple[str, str]:
    """Remove PowerShell comments from one line: # to the end of the line and
    <# ... #> blocks, which may span lines; a # or <# inside quotes is text.
    state carries what is open at the end of the previous line: "" (code),
    "<#" (a block comment), '"' or "'" (a string) or '@"' or "@'" (a
    here-string, whose lines are text and are left out of the code).
    In a double-quoted string a backtick escapes the next character."""
    if state in ('@"', "@'"):
        if not line.lstrip().startswith(state[1] + "@"):
            return "", state
        line, state = line.lstrip()[2:], ""
    in_block = state == "<#"
    quote = state if state in ('"', "'") else ""
    code: list[str] = []
    i = 0
    while i < len(line):
        if in_block:
            if line.startswith("#>", i):
                in_block, i = False, i + 2
            else:
                i += 1
            continue
        char = line[i]
        if quote == '"' and char == "`":
            # A backtick escapes the next character of a double-quoted
            # string: `" does not close it and `` is a literal backtick. In
            # a single-quoted string a backtick is literal.
            code.append(line[i:i + 2])
            i += 2
            continue
        if quote:
            quote = "" if char == quote else quote
        elif line.startswith(('@"', "@'"), i) and not line[i + 2:].strip():
            return "".join(code), line[i:i + 2]
        elif char in "\"'":
            quote = char
        elif line.startswith("<#", i):
            in_block, i = True, i + 2
            continue
        elif char == "#":
            break
        code.append(char)
        i += 1
    return "".join(code), "<#" if in_block else quote


def sync_roots(text: str) -> tuple[list[str], list[str]]:
    """Read the active entries of $sourceRoots: the directories, with "/"
    separators and a trailing "/", and the problems found."""
    lines = text.splitlines()
    # Comments and strings are tracked from the top of the file, so that a
    # declaration inside a <# ... #> block, after a # or in a here-string is
    # never taken as the active one, and a "<#" inside a string opens no
    # comment .
    state = ""
    start = None
    for number, line in enumerate(lines):
        code, state = strip_comment(line, state)
        if SYNC_START.match(code):
            start = number
            break
    if start is None:
        return [], ["$sourceRoots list not found"]
    roots: list[str] = []
    problems: list[str] = []
    for line in lines[start + 1:]:
        code, state = strip_comment(line, state)
        code = code.strip()
        if not code:
            continue
        if code == ")":
            return roots, problems
        entry = SYNC_ENTRY.fullmatch(code)
        if entry:
            roots.append(entry.group(1).replace("\\", "/").strip("/") + "/")
        else:
            problems.append(f"unsupported $sourceRoots entry: {code}")
    return roots, problems + ["$sourceRoots list is not closed"]


def check_sync_script(root: Path, tracked: set[str], components: list[str]) -> list[str]:
    """Every VBA component must lie in a directory the synchronizer reads."""
    missing = required_file(root, tracked, SYNC_SCRIPT)
    if missing:
        return missing
    roots, problems = sync_roots((root / SYNC_SCRIPT).read_text(encoding="utf-8"))
    findings = [f"{SYNC_SCRIPT}: {problem}" for problem in problems]
    findings.extend(f"{SYNC_SCRIPT}: source directory {r} does not exist" for r in roots
                    if not (root / r).is_dir())
    findings.extend(f"{p}: not synchronized; add its directory to $sourceRoots in {SYNC_SCRIPT}"
                    for p in components if not p.startswith(tuple(roots)))
    return findings


def check_document_modules(root: Path, tracked: set[str], components: list[str]) -> list[str]:
    """Document modules and the template's sheets must correspond one to one."""
    missing = required_file(root, tracked, TEMPLATE)
    if missing:
        return missing
    try:
        with zipfile.ZipFile(root / TEMPLATE) as package:
            names = {match for part in package.namelist()
                     if part == "xl/workbook.xml" or part.startswith("xl/worksheets/sheet")
                     for match in CODE_NAME.findall(package.read(part).decode("utf-8", errors="replace"))}
    except (OSError, zipfile.BadZipFile):
        return []    # check_workbook reports an unreadable template
    modules = {Path(p).stem for p in components
               if p.startswith(DOCUMENT_HOME) and p.lower().endswith(".cls")}
    findings = [f"{DOCUMENT_HOME}{name}.cls: no sheet or workbook with this CodeName in {TEMPLATE}"
                for name in sorted(modules - names)]
    findings.extend(f"{TEMPLATE}: CodeName {name} has no document module {DOCUMENT_HOME}{name}.cls"
                    for name in sorted(names - modules))
    return findings


def check_version(root: Path) -> list[str]:
    """VERSION exists from the first release on and names the newest dated changelog release."""
    changelog = root / "CHANGELOG.md"
    text = changelog.read_text(encoding="utf-8") if changelog.is_file() else ""
    released = [label for label, _ in VERSION_HEADING.findall(text) if label != "Unreleased"]
    path = root / "VERSION"
    if not path.is_file():
        return [f"VERSION is missing; CHANGELOG.md releases [{released[0]}]"] if released else []
    raw = path.read_text(encoding="utf-8")
    if not re.fullmatch(SEMVER + r"\n", raw):
        return ["VERSION must contain one X.Y.Z line ending in a newline"]
    if not released:
        return ["VERSION exists but CHANGELOG.md has no release heading"]
    if raw.strip() != released[0]:
        return [f"VERSION {raw.strip()} differs from the newest CHANGELOG.md release [{released[0]}]"]
    return []


def run_check(root: Path) -> dict[str, Any]:
    tracked = tracked_files(root)
    findings = check_storage(root)
    names: dict[str, str] = {}
    components = sorted(p for p in tracked if Path(p).suffix.lower() in VBA_SUFFIXES)
    findings.extend(f"{p}: VBA component outside the documented source locations"
                    for p in sorted(tracked)
                    if Path(p).suffix.lower() in VBA_SUFFIXES | {".frx"} and not p.startswith(VBA_HOMES))
    for path in components:
        findings.extend(check_component(root, path, tracked, names))
    for path in sorted(p for p in tracked if Path(p).suffix.lower() in WORKBOOK_SUFFIXES):
        if path != TEMPLATE:
            findings.append(f"{path}: only {TEMPLATE} may be tracked")
        findings.extend(check_workbook(root, path))
    findings.extend(check_sync_script(root, tracked, components))
    findings.extend(check_document_modules(root, tracked, components))
    findings.extend(check_changelog(root))
    findings.extend(check_version(root))
    return {"schema_version": 1, "status": "fail" if findings else "pass",
            "components": len(components), "findings": findings}


def markdown_report(report: dict[str, Any]) -> str:
    lines = [f"Source integrity: {report['status'].upper()} ({report['components']} VBA component(s))"]
    lines.extend(report["findings"])
    return "\n".join(lines) + "\n"


def main() -> int:
    options = parse_report_args(sys.argv[1:], description=__doc__)
    return run_gate(options, build=lambda: run_check(options.root.resolve()),
                    markdown=markdown_report, errors=(OSError, RuntimeError, ValueError))


if __name__ == "__main__":
    raise SystemExit(main())
