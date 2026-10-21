#!/usr/bin/env python3
"""Compare the VBA exported from a built workbook with the repository source.

Part of the source round trip (issue #6): a workbook built from the template
and the source of one commit, then exported with tools/Export-IRRBB-VBA.ps1, must
hold exactly the components that tools/vba-projects.json lists for its mode,
each with the same code and attributes as the source of that commit. Line
endings are ignored, and so are trailing blank lines and the
VERSION/BEGIN/END block of a class file. So is the letter case of code
outside string literals and comments: VBA names are case-insensitive, and
the VBE gives each name one spelling across the project, so it re-cases
identifiers on import; the lines that differ only that way are counted. The
attributes of a document module are not compared either: the synchronizer
replaces only its code, and Excel writes its attributes. Everything else must
match. The source is read from Git at the candidate commit, not from the
working tree. The tool never runs Excel. See docs/EXCEL_EVIDENCE.md.

Source parity says nothing about which workbook the files came from. The
exporter writes export-manifest.json: the SHA-256 of the workbook it
opened and of every file it exported, the .frx of a form included. With --workbook the tool checks the
workbook and the files against it and reports the binding as VERIFIED, as
FAIL, or as UNVERIFIED for an export without a manifest. A manifest is a
record, not a signature: it binds an export to a workbook, but it does not
attest who made either.
"""
from __future__ import annotations

import argparse
import difflib
import hashlib
import json
import re
import sys
import zipfile
from dataclasses import dataclass
from pathlib import Path

from _gatelib import git_bytes

MANIFEST = "tools/vba-projects.json"
EXPORT_MANIFEST = "export-manifest.json"
EXPORT_FORMAT = "irrbb-vba-export-1"
VBA_PROJECT_PART = "xl/vbaProject.bin"
TEMPLATE = "src/workbook/IRRBB_Template.xlsx"
VBA_ENCODING = "cp1252"
SUFFIXES = (".bas", ".cls", ".frm")
# A form's binary properties, which the VBE exports beside its .frm: not
# compared with the source, but bound to the workbook like the components.
RESOURCE_SUFFIX = ".frx"
VB_NAME = re.compile(r'^Attribute VB_Name = "([^"]+)"$')
REM = re.compile(r"rem(?=\s|$)", re.IGNORECASE)
# Lines the VBE writes before the code of a class or form: the VERSION line
# ("VERSION 1.0 CLASS" for a class, "VERSION 5.00" for a form), the
# BEGIN ... END property block (title case and nested in a form) and the
# Attribute lines.
CLASS_HEADER = re.compile(r"^(VERSION \d+\.\d+( CLASS)?|(?i:begin)(\s.*)?|(?i:end)|\s+\S.*)$")


@dataclass(frozen=True)
class Component:
    name: str
    attributes: tuple[str, ...]
    body: tuple[str, ...]


@dataclass(frozen=True)
class Expected:
    path: str
    component: Component
    document: bool


def parse(raw: bytes, label: str) -> Component:
    """Split a VBA file into its VB_Name, its Attribute lines and its code."""
    lines = raw.decode(VBA_ENCODING).replace("\r\n", "\n").replace("\r", "\n").split("\n")
    index = 0
    attributes: list[str] = []
    while index < len(lines):
        line = lines[index]
        if line.startswith("Attribute "):
            attributes.append(line)
        elif not CLASS_HEADER.match(line) or attributes:
            break
        index += 1
    body = lines[index:]
    while body and body[-1] == "":
        body.pop()
    names = [m.group(1) for m in map(VB_NAME.match, attributes) if m]
    if len(names) != 1:
        raise ValueError(f"{label}: expected one Attribute VB_Name line, found {len(names)}")
    return Component(names[0], tuple(attributes), tuple(body))


def fold_case(line: str) -> str:
    """``line`` with its code lower-cased, string literals and the comment kept.

    A comment starts at an apostrophe outside a string, or at ``Rem`` where a
    statement starts: the start of the line, after a line number or after a
    colon."""
    folded: list[str] = []
    in_string = False
    statement_start = True
    index = 0
    while index < len(line):
        char = line[index]
        if in_string:
            in_string = char != '"'
        elif char == '"':
            in_string, statement_start = True, False
        elif char == "'":
            return "".join(folded) + line[index:]
        elif char == ":":
            statement_start = True
        elif statement_start and REM.match(line, index):
            return "".join(folded) + "rem" + line[index + 3:]
        elif statement_start and char.isdigit() and not "".join(folded).strip(" \t0123456789"):
            pass
        elif not char.isspace():
            statement_start = False
        folded.append(char if in_string or char == '"' else char.lower())
        index += 1
    return "".join(folded)


def committed(root: Path, sha: str, path: str) -> bytes:
    completed = git_bytes(root, "show", f"{sha}:{path}")
    if completed.returncode != 0:
        raise ValueError(f"{path}: not found at {sha}")
    return completed.stdout


def expected_components(root: Path, sha: str, mode: str) -> dict[str, Expected]:
    """The components of ``mode`` at ``sha``, by lower-case VB_Name."""
    manifest = json.loads(committed(root, sha, MANIFEST).decode("utf-8"))
    configuration = manifest["configurations"].get(mode)
    if configuration is None:
        raise ValueError(f"{MANIFEST}: no configuration '{mode}'")
    documents = set(manifest["document_modules"])
    expected: dict[str, Expected] = {}
    for path in [*configuration["modules"], *manifest["document_modules"]]:
        component = parse(committed(root, sha, path), path)
        expected[component.name.lower()] = Expected(path, component, path in documents)
    return expected


def exported_components(folder: Path) -> dict[str, tuple[str, Component]]:
    exported: dict[str, tuple[str, Component]] = {}
    for path in sorted(folder.iterdir()):
        if path.suffix.lower() not in SUFFIXES:
            continue
        component = parse(path.read_bytes(), path.name)
        key = component.name.lower()
        if key in exported:
            raise ValueError(f"{path.name}: VB_Name {component.name} exported twice")
        exported[key] = (path.name, component)
    return exported


def first_difference(source: tuple[str, ...], export: tuple[str, ...]) -> str:
    for number, (left, right) in enumerate(zip(source, export), start=1):
        if fold_case(left) != fold_case(right):
            return f"line {number}: source {left!r}, export {right!r}"
    shorter = min(len(source), len(export))
    return f"line {shorter + 1}: source has {len(source)} lines, export {len(export)}"


def compare(expected: dict[str, Expected],
            exported: dict[str, tuple[str, Component]]) -> tuple[list[str], int]:
    """The differences between the source and the export, empty when they
    match, and the number of lines that differ only in letter case."""
    findings = [f"missing from the export: {expected[key].component.name} ({expected[key].path})"
                for key in sorted(set(expected) - set(exported))]
    findings += [f"not in the source for this mode: {exported[key][1].name} ({exported[key][0]})"
                 for key in sorted(set(exported) - set(expected))]
    recased = 0
    for key in sorted(set(expected) & set(exported)):
        path, source = expected[key].path, expected[key].component
        _, export = exported[key]
        if source.name != export.name:
            findings.append(f"{path}: VB_Name {source.name} exported as {export.name}")
        if not expected[key].document and source.attributes != export.attributes:
            changed = [line for line in difflib.unified_diff(
                source.attributes, export.attributes, lineterm="", n=0)
                if line[:1] in "+-" and line[:3] not in ("+++", "---")]
            findings.append(f"{path}: attributes differ: {'; '.join(changed)}")
        if [fold_case(line) for line in source.body] != [fold_case(line) for line in export.body]:
            findings.append(f"{path}: code differs at {first_difference(source.body, export.body)}")
        else:
            recased += sum(left != right for left, right in zip(source.body, export.body))
    return findings, recased


def sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


@dataclass(frozen=True)
class ExportRecord:
    workbook_name: str
    workbook_sha256: str
    files: dict[str, tuple[str, str, str]]  # lower-case file -> (file, VB_Name, sha256)


def load_export_manifest(path: Path) -> ExportRecord:
    """export-manifest.json; ValueError when it is not a valid manifest."""
    try:
        record = json.loads(path.read_bytes().decode("utf-8-sig"))
        if not isinstance(record, dict) or record.get("format") != EXPORT_FORMAT:
            raise ValueError(f"format is not {EXPORT_FORMAT}")
        name = record["workbook"]["file"]
        if not isinstance(name, str) or not name:
            raise ValueError("workbook.file is not a file name")
        files = {item["file"].lower(): (item["file"], item["name"], item["sha256"].lower())
                 for item in record["components"]}
        if len(files) != len(record["components"]):
            raise ValueError("a file is listed twice")
        return ExportRecord(name, record["workbook"]["sha256"].lower(), files)
    except (ValueError, KeyError, TypeError, AttributeError) as error:
        raise ValueError(f"{EXPORT_MANIFEST} is not a valid export manifest: {error}") from error


def file_findings(folder: Path, record: ExportRecord) -> list[str]:
    """The exported files that are missing, unlisted, changed or renamed
    against the manifest."""
    listed = record.files
    present = {item.name.lower(): item for item in folder.iterdir()
               if item.is_file() and item.suffix.lower() in (*SUFFIXES, RESOURCE_SUFFIX)}
    findings = [f"{EXPORT_MANIFEST} lists {listed[key][0]}, which is not in the export"
                for key in sorted(set(listed) - set(present))]
    findings += [f"{present[key].name} is in the export but not in {EXPORT_MANIFEST}"
                 for key in sorted(set(present) - set(listed))]
    for key in sorted(set(listed) & set(present)):
        name, component, digest = listed[key]
        data = present[key].read_bytes()
        if sha256(data) != digest:
            findings.append(f"{name} changed after the export: its SHA-256 is not the one in {EXPORT_MANIFEST}")
        elif present[key].suffix.lower() != RESOURCE_SUFFIX:
            try:
                exported_name = parse(data, name).name
            except ValueError:
                continue  # reported by the source comparison
            if exported_name != component:
                findings.append(f"{name} declares VB_Name {exported_name}, {EXPORT_MANIFEST} names {component}")
    return findings


def check_binding(folder: Path, workbook: Path | None) -> tuple[str, list[str], list[str]]:
    """The binding of the export to the workbook, VERIFIED, UNVERIFIED or
    FAIL, the findings that make it FAIL and notes that do not.

    A manifest that does not match the exported files fails even without
    --workbook; without either, or without a manifest, the binding is
    UNVERIFIED. The workbook is bound by its bytes: a workbook renamed after
    the export, as a download often is, gets a note, not a failure."""
    path = folder / EXPORT_MANIFEST
    if not path.is_file():
        return "UNVERIFIED", [], []
    try:
        record = load_export_manifest(path)
    except ValueError as error:
        return "FAIL", [str(error)], []
    findings = file_findings(folder, record)
    notes: list[str] = []
    if workbook is not None:
        data = workbook.read_bytes()
        same_bytes = sha256(data) == record.workbook_sha256
        if not same_bytes:
            findings.append(f"{workbook.name} is not the workbook the export was made from: "
                            f"sha256 {sha256(data)}, {EXPORT_MANIFEST} records {record.workbook_sha256}")
        if not is_macro_workbook(workbook):
            findings.append(f"{workbook.name} is not a macro-enabled workbook ({VBA_PROJECT_PART} not found)")
        if same_bytes and workbook.name.lower() != record.workbook_name.lower():
            notes.append(f"{EXPORT_MANIFEST} names the workbook {record.workbook_name}; "
                         "it was renamed since the export")
    if findings:
        return "FAIL", findings, notes
    return ("VERIFIED" if workbook is not None else "UNVERIFIED"), [], notes


def is_macro_workbook(path: Path) -> bool:
    try:
        with zipfile.ZipFile(path) as archive:
            return VBA_PROJECT_PART in archive.namelist()
    except (zipfile.BadZipFile, OSError):
        return False


def main(arguments: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--root", type=Path, default=Path(__file__).resolve().parent.parent)
    parser.add_argument("--export", type=Path, required=True, help="folder written by Export-IRRBB-VBA.ps1")
    parser.add_argument("--mode", choices=("production", "development"), required=True)
    parser.add_argument("--candidate-sha", default="HEAD", help="commit the workbook was built from")
    parser.add_argument("--workbook", type=Path,
                        help="built .xlsm the export was made from, checked against export-manifest.json")
    options = parser.parse_args(arguments)
    try:
        resolved = git_bytes(options.root, "rev-parse", "--verify", f"{options.candidate_sha}^{{commit}}")
        if resolved.returncode != 0:
            raise ValueError(f"unknown commit: {options.candidate_sha}")
        sha = resolved.stdout.decode().strip()
        if not options.export.is_dir():
            raise ValueError(f"export folder not found: {options.export}")
        expected = expected_components(options.root, sha, options.mode)
        exported = exported_components(options.export)
        findings, recased = compare(expected, exported)
        if options.workbook is not None and not options.workbook.is_file():
            raise ValueError(f"workbook not found: {options.workbook}")
        binding, binding_findings, binding_notes = check_binding(options.export, options.workbook)
        template = sha256(committed(options.root, sha, TEMPLATE))
    except ValueError as error:
        print(f"ERROR: {error}", file=sys.stderr)
        return 2
    print(f"Candidate: {sha}")
    print(f"Mode:      {options.mode} ({len(expected)} components expected, {len(exported)} exported)")
    print(f"Template:  {TEMPLATE} sha256 {template}")
    if options.workbook is not None:
        print(f"Workbook:  {options.workbook.name} sha256 {sha256(options.workbook.read_bytes())}")
    print(f"Re-cased:  {recased} line(s) differ only in the letter case of code (ignored)")
    print(f"Binding:   {binding} - {binding_note(binding, options.export, options.workbook)}")
    for note in binding_notes:
        print(f"  NOTE {note}")
    for finding in findings + binding_findings:
        print(f"  FAIL {finding}")
    passed = not findings and binding != "FAIL"
    print(f"RESULT={'PASS' if passed else 'FAIL'}; components={len(expected)}; "
          f"differences={len(findings)}; binding={binding}")
    return 0 if passed else 1


def binding_note(binding: str, folder: Path, workbook: Path | None) -> str:
    if binding == "VERIFIED":
        return f"{EXPORT_MANIFEST} records this workbook and every exported file"
    if binding == "FAIL":
        return f"the export, {EXPORT_MANIFEST} and the workbook do not agree"
    if not (folder / EXPORT_MANIFEST).is_file():
        return f"no {EXPORT_MANIFEST}: the export is not bound to a workbook"
    return "no --workbook given: the export is not checked against a workbook"


if __name__ == "__main__":
    sys.exit(main())
