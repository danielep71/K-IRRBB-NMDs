#!/usr/bin/env python3
"""Validate explicit VBA public API declarations and the checked-in manifest."""

from __future__ import annotations

import re
import subprocess
import sys
import tempfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from _gatelib import git_bytes as git
from _gatelib import parse_report_args as parse_args
from _gatelib import run_gate
from _vba_lex import logical_lines
from check_vba_conditionals import reachable_sources

MANIFEST_PATH = "docs/PUBLIC_API.txt"
# IRRBB derives each component's role from its location instead of a profile
# registry; see docs/REPOSITORY_STRUCTURE.md. Only src/modules is public.
ROLE_BY_PREFIX = (
    ("src/modules/", "public"),
    ("src/core/", "internal"),
    ("src/classes/", "internal"),
    ("src/workbook/", "internal"),
    ("src/forms/", "internal"),
    ("tests/", "test"),
    ("examples/", "example"),
)
TOOL_NAME = "VBA public API"
# Enum members share the project namespace, so they take part in collision
# checks; they are recorded in the manifest only through their Enum's body.
ENUM_MEMBER = "Enum Member"
VBA_SUFFIXES = {".bas", ".cls", ".frm"}
SIG_PREFIX = "# SIG\t"
PROPERTY_KINDS = {"property get", "property let", "property set"}

PROC = re.compile(
    r"^\s*(?:(?P<vis>Public|Private|Friend|Global)\s+)?(?:Static\s+)?"
    r"(?P<kind>Sub|Function|Property\s+(?:Get|Let|Set))\s+"
    r"(?P<name>[A-Za-z_]\w*)\b",
    re.I,
)
END_PROC = re.compile(r"^\s*End\s+(?:Sub|Function|Property)\b", re.I)
BLOCK = re.compile(r"^\s*(?P<vis>Public|Global)\s+(?P<kind>Enum|Type)\s+(?P<name>[A-Za-z_]\w*)\b", re.I)
END_BLOCK = {
    "enum": re.compile(r"^\s*End\s+Enum\b", re.I),
    "type": re.compile(r"^\s*End\s+Type\b", re.I),
}
DECLARE = re.compile(
    r"^\s*(?:Public|Global)\s+Declare\s+(?:PtrSafe\s+)?"
    r"(?P<kind>Function|Sub)\s+(?P<name>[A-Za-z_]\w*)\b",
    re.I,
)
EVENT = re.compile(r"^\s*(?:Public|Global)\s+Event\s+(?P<name>[A-Za-z_]\w*)\b", re.I)
CONST = re.compile(
    r"^\s*(?:Public|Global)\s+Const\s+(?P<name>[A-Za-z_]\w*)\b(?P<rest>.*)$",
    re.I,
)
VARIABLE = re.compile(
    r"^\s*(?:Public|Global)\s+(?P<withevents>WithEvents\s+)?"
    r"(?P<name>[A-Za-z_]\w*)\b(?P<rest>.*)$",
    re.I,
)
IMPLICIT = re.compile(
    r"^\s*(?:Static\s+)?(?:Sub|Function|Property\s+(?:Get|Let|Set)|"
    r"Enum|Type|Event|Declare\s+(?:PtrSafe\s+)?(?:Function|Sub))\b",
    re.I,
)


def tracked_vba(root: Path) -> list[str]:
    completed = git(root, "ls-files", "-z")
    if completed.returncode:
        raise RuntimeError(completed.stderr.decode("utf-8", errors="replace").strip())
    return sorted(
        path.decode("utf-8", errors="surrogateescape")
        for path in completed.stdout.split(b"\0")
        if path
        and Path(path.decode("utf-8", errors="surrogateescape")).suffix.casefold()
        in VBA_SUFFIXES
    )


def is_date_delimiter(code: str, index: int, in_date: bool) -> bool:
    if in_date:
        return True
    # An adjacent identifier/numeric token owns its # type suffix.
    previous = code[index - 1] if index else ""
    if previous and (previous.isalnum() or previous in "_.)]"):
        return False
    # A literal needs a closing delimiter; e.g. Print #1 is not a date.
    return "#" in code[index + 1:]


def split_statements(code: str) -> list[str]:
    """Split colons outside VBA strings, preserving named arguments and Rem comments."""
    statements: list[str] = []
    start = 0
    in_string = False
    in_date = False
    index = 0
    while index < len(code):
        char = code[index]
        if char == '"' and not in_date:
            if in_string and code[index:index + 2] == '""':
                index += 2
                continue
            in_string = not in_string
        elif char == "#" and not in_string and is_date_delimiter(code, index, in_date):
            in_date = not in_date
        elif char == ":" and not (in_string or in_date) and code[index:index + 2] != ":=":
            statement = code[start:index].strip()
            if re.match(r"^Rem(?:\s|$)", statement, re.I):
                return statements
            statements.append(statement)
            start = index + 1
        index += 1
    statement = code[start:].strip()
    if not re.match(r"^Rem(?:\s|$)", statement, re.I):
        statements.append(statement)
    return statements


def logical(lines: list[str]) -> list[tuple[int, int, str]]:
    return [(start, end, part) for start, end, code in logical_lines(lines)
            for part in split_statements(code)]


def norm(text: str) -> str:
    return re.sub(r"\s+", " ", text.strip())


def key(component: str, kind: str, name: str) -> str:
    return f"{component}\t{kind}\t{name}"


def top_comma(text: str) -> bool:
    depth = 0
    in_string = False
    index = 0
    while index < len(text):
        character = text[index]
        if character == '"':
            if in_string and index + 1 < len(text) and text[index + 1] == '"':
                index += 2
                continue
            in_string = not in_string
        elif not in_string:
            if character == "(":
                depth += 1
            elif character == ")" and depth:
                depth -= 1
            elif character == "," and depth == 0:
                return True
        index += 1
    return False


def record(
    output: list[dict[str, Any]],
    supported: bool,
    component: str,
    kind: str,
    name: str,
    signature: str,
    path: str,
    line: int,
) -> None:
    output.append(
        {
            "component": component,
            "kind": kind,
            "name": name,
            "signature": signature,
            "path": path,
            "line": line,
            "supported": supported,
        }
    )


@dataclass
class _ComponentScan:
    """Declarations and findings collected while scanning one active component source."""

    path: str
    supported: bool
    declarations: list[dict[str, Any]] = field(default_factory=list)
    findings: list[dict[str, Any]] = field(default_factory=list)

    def declare(self, kind: str, name: str, signature: str, line: int) -> None:
        record(self.declarations, self.supported, Path(self.path).stem, kind, name, signature, self.path, line)

    def reject(self, line: int, message: str) -> None:
        self.findings.append({"path": self.path, "line": line, "message": message})


def _procedure_header(scan: _ComponentScan, match: re.Match[str], code: str, line: int) -> None:
    """A procedure must state its visibility; public procedures are recorded."""
    visibility = (match.group("vis") or "").casefold()
    if not visibility:
        scan.reject(line, f"Implicit public procedure is prohibited: {match.group('name')}.")
    if visibility == "public":
        kind = " ".join(item.capitalize() for item in match.group("kind").split())
        scan.declare(kind, match.group("name"), norm(code), line)


def _public_block(
    scan: _ComponentScan, match: re.Match[str], statements: list[tuple[int, int, str]], index: int, line: int
) -> int:
    """Record a public Type/Enum block with its body; return the index after it.

    Global Enum and Global Type are rejected: sources disagree on whether VBA accepts
    Global Enum at all, so the project declares both Public."""
    kind = match.group("kind").capitalize()
    name = match.group("name")
    body = [norm(statements[index][2])]
    close = END_BLOCK[kind.casefold()]
    if match.group("vis").casefold() == "global":
        scan.reject(line, f"Global {kind} {name} is not allowed; declare it Public.")
    end_index = index + 1
    while end_index < len(statements) and not close.match(statements[end_index][2]):
        if statements[end_index][2].strip():
            body.append(norm(statements[end_index][2]))
        end_index += 1
    if end_index >= len(statements):
        scan.reject(line, f"Public {kind} {name} is not closed.")
        return index + 1
    body.append(norm(statements[end_index][2]))
    scan.declare(kind, name, " | ".join(body), line)
    if kind == "Enum":
        for member_line, _, member_code in statements[index + 1:end_index]:
            member = re.match(r"\s*([A-Za-z_]\w*)", member_code)
            if member:
                scan.declare(ENUM_MEMBER, member.group(1), norm(member_code), member_line)
    return end_index + 1


def _module_declaration(code: str) -> tuple[str, str] | str | None:
    """Classify a single-statement module-level public declaration.

    Returns ``(kind, name)`` to record, a message for a prohibited form, or ``None``.
    """
    match = DECLARE.match(code)
    if match:
        kind = "Declare Function" if match.group("kind").casefold() == "function" else "Declare Sub"
        return kind, match.group("name")
    match = EVENT.match(code)
    if match:
        return "Event", match.group("name")
    match = CONST.match(code)
    if match:
        if top_comma(match.group("rest")):
            return "Public Const declarations must contain one identifier per statement."
        return "Const", match.group("name")
    match = VARIABLE.match(code)
    if match:
        if top_comma(match.group("rest")):
            return "Public variable declarations must contain one identifier per statement."
        return ("WithEvents Variable" if match.group("withevents") else "Variable"), match.group("name")
    return None


def _parse_active_component(
    path: str, text: str, supported: bool
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    scan = _ComponentScan(path, supported)
    statements = logical(text.replace("\r\n", "\n").split("\n"))
    inside_procedure = False
    index = 0
    while index < len(statements):
        line, _, code = statements[index]
        stripped = code.strip()
        next_index = index + 1
        procedure = PROC.match(code)
        block = BLOCK.match(code)
        if not stripped or stripped.startswith(("Attribute ", "Option ", "#")):
            pass
        elif procedure and " declare " not in f" {code.casefold()} ":
            _procedure_header(scan, procedure, code, line)
            inside_procedure = True
        elif END_PROC.match(code):
            inside_procedure = False
        elif inside_procedure:
            pass
        elif IMPLICIT.match(code):
            scan.reject(line, f"Implicit public module-level declaration is prohibited: {norm(code)}")
        elif block:
            next_index = _public_block(scan, block, statements, index, line)
        else:
            outcome = _module_declaration(code)
            if isinstance(outcome, str):
                scan.reject(line, outcome)
            elif outcome is not None:
                if outcome[0] == "Event" and Path(path).suffix.casefold() == ".bas":
                    scan.reject(line, f"Event {outcome[1]} is declared in a standard module; "
                                      "events are valid only in class and form modules.")
                scan.declare(outcome[0], outcome[1], norm(code), line)
        index = next_index
    return scan.declarations, scan.findings


def parse_component(
    path: str, text: str, supported: bool
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    sources, findings = reachable_sources(path, text)
    declarations: dict[tuple[int, str, str, str], dict[str, Any]] = {}
    for environment, source in sources.items():
        parsed, errors = _parse_active_component(path, source, supported)
        for error in errors:
            if error not in findings:
                findings.append(error)
        for item in parsed:
            identity = (item["line"], item["kind"], item["name"], item["signature"])
            if identity not in declarations:
                declarations[identity] = {**item, "environments": []}
            declarations[identity]["environments"].append(environment)
    return list(declarations.values()), findings


def read_manifest(
    root: Path, relative: str
) -> tuple[set[str], dict[str, set[str]], list[dict[str, Any]]]:
    path = root / relative
    if not path.is_file():
        return set(), {}, [
            {"path": relative, "message": "Configured public API manifest is missing."}
        ]
    rows: set[str] = set()
    signatures: dict[str, set[str]] = {}
    findings: list[dict[str, Any]] = []
    for number, raw in enumerate(path.read_text(encoding="utf-8").splitlines(), 1):
        if not raw.strip():
            continue
        if raw.startswith(SIG_PREFIX):
            fields = raw[len(SIG_PREFIX) :].split("\t", 3)
            if len(fields) != 4:
                findings.append(
                    {"path": relative, "line": number, "message": "Malformed # SIG record."}
                )
                continue
            declaration_key = key(fields[0], fields[1], fields[2])
            canonical = next((existing for existing in signatures
                              if existing.casefold() == declaration_key.casefold()), declaration_key)
            variants = signatures.setdefault(canonical, set())
            if fields[3] in variants:
                findings.append({"path": relative, "line": number,
                                 "message": f"Duplicate signature record: {declaration_key}"})
            variants.add(fields[3])
            continue
        if raw.lstrip().startswith("#"):
            continue
        fields = raw.split("\t")
        if len(fields) != 3 or any(not item for item in fields):
            findings.append(
                {
                    "path": relative,
                    "line": number,
                    "message": (
                        "Manifest declaration rows require component, kind, and name."
                    ),
                }
            )
            continue
        declaration_key = key(*fields)
        if any(existing.casefold() == declaration_key.casefold() for existing in rows):
            findings.append(
                {
                    "path": relative,
                    "line": number,
                    "message": f"Duplicate manifest declaration: {declaration_key}",
                }
            )
        rows.add(declaration_key)
    return rows, signatures, findings


def is_property(kind: str) -> bool:
    return kind.casefold() in PROPERTY_KINDS


def property_pair(first: dict[str, Any], second: dict[str, Any]) -> bool:
    return (
        str(first["component"]).casefold() == str(second["component"]).casefold()
        and is_property(str(first["kind"]))
        and is_property(str(second["kind"]))
        and str(first["kind"]).casefold() != str(second["kind"]).casefold()
    )


def _parse_components(
    root: Path, components: dict[str, str]
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    """Parse every tracked component that has a role; every declaration is returned, flagged
    "supported" only for public-role components, which alone enter the manifest."""
    declarations: list[dict[str, Any]] = []
    findings: list[dict[str, Any]] = []
    tracked = set(tracked_vba(root))
    for path, role in sorted(components.items()):
        if path not in tracked or not (root / path).is_file():
            continue
        parsed, errors = parse_component(
            path,
            (root / path).read_bytes().decode("cp1252"),
            role == "public",
        )
        declarations.extend(parsed)
        findings.extend(errors)
    return declarations, findings


def _in_project_namespace(declaration: dict[str, Any]) -> bool:
    """Standard-module Public names, and Public Enums with their members from any
    module type (class modules included), share one project namespace."""
    return (Path(str(declaration["path"])).suffix.casefold() == ".bas"
            or declaration["kind"] in {"Enum", ENUM_MEMBER})


def _name_collisions(
    declaration: dict[str, Any], names: dict[str, list[dict[str, Any]]], scope: str,
    reported: bool = False,
) -> list[dict[str, Any]]:
    """A public name may not collide with another of its scope in an overlapping environment.

    ``reported`` skips pairs of project-wide names, which the project scope already reports.
    Events have their own namespace: an event collides only with another event."""
    findings: list[dict[str, Any]] = []
    prior = names.setdefault(str(declaration["name"]).casefold(), [])
    for previous in prior:
        if reported and _in_project_namespace(previous) and _in_project_namespace(declaration):
            continue
        if (previous["kind"] == "Event") != (declaration["kind"] == "Event"):
            continue
        if (set(previous["environments"]) & set(declaration["environments"])
                and not property_pair(previous, declaration)):
            findings.append(
                {
                    "path": declaration["path"],
                    "line": declaration["line"],
                    "message": (
                        f"Public {scope} name "
                        f"{declaration['name']!r} collides with "
                        f"{previous['component']}.{previous['name']} "
                        f"({previous['kind']})."
                    ),
                }
            )
    prior.append(declaration)
    return findings


def _index_declarations(
    declarations: list[dict[str, Any]],
) -> tuple[dict[str, list[dict[str, Any]]], list[dict[str, Any]]]:
    """Group declarations by case-insensitive key; report duplicates and name collisions,
    project-wide and, for class and form modules, within the module."""
    keys: dict[str, list[dict[str, Any]]] = {}
    names: dict[str, list[dict[str, Any]]] = {}
    modules: dict[str, dict[str, list[dict[str, Any]]]] = {}
    findings: list[dict[str, Any]] = []
    for declaration in declarations:
        declaration_key = key(
            str(declaration["component"]),
            str(declaration["kind"]),
            str(declaration["name"]),
        )
        canonical = next((existing for existing in keys
                          if existing.casefold() == declaration_key.casefold()), declaration_key)
        variants = keys.setdefault(canonical, [])
        if any(set(previous["environments"]) & set(declaration["environments"])
               for previous in variants):
            findings.append({
                "path": declaration["path"], "line": declaration["line"],
                "message": "Public declaration appears more than once: " + declaration_key,
            })
        variants.append(declaration)
        if _in_project_namespace(declaration):
            findings.extend(_name_collisions(declaration, names, "project-wide"))
        if Path(str(declaration["path"])).suffix.casefold() != ".bas":
            module = modules.setdefault(str(declaration["path"]), {})
            findings.extend(_name_collisions(declaration, module, "module", reported=True))
    return keys, findings


def _manifest_findings(
    manifest: str,
    keys: dict[str, list[dict[str, Any]]],
    rows: set[str],
    signatures: dict[str, set[str]],
) -> list[dict[str, Any]]:
    """Source declarations, manifest rows and signature records must match one to one."""
    findings: list[dict[str, Any]] = []
    actual = {item.casefold(): item for item in keys}
    row_map = {item.casefold(): item for item in rows}
    signature_map = {item.casefold(): item for item in signatures}
    for folded, declaration_key in actual.items():
        if folded not in row_map:
            findings.append(
                {
                    "path": manifest,
                    "message": f"Public declaration is not recorded: {declaration_key}",
                }
            )
        signature_key = signature_map.get(folded)
        if signature_key is None:
            findings.append(
                {
                    "path": manifest,
                    "message": (
                        "Normalized signature record is missing: "
                        f"{declaration_key}"
                    ),
                }
            )
        elif signatures[signature_key] != {item["signature"] for item in keys[declaration_key]}:
            findings.append(
                {
                    "path": manifest,
                    "message": (
                        f"Signature mismatch for {declaration_key}: expected "
                        f"{sorted({item['signature'] for item in keys[declaration_key]})!r}, recorded "
                        f"{sorted(signatures[signature_key])!r}"
                    ),
                }
            )
    for folded, declaration_key in row_map.items():
        if folded not in actual:
            findings.append(
                {
                    "path": manifest,
                    "message": (
                        "Manifest declaration is not present in supported source: "
                        f"{declaration_key}"
                    ),
                }
            )
    for folded, declaration_key in signature_map.items():
        if folded not in actual:
            findings.append(
                {
                    "path": manifest,
                    "message": f"Stale signature record: {declaration_key}",
                }
            )
    return findings


def component_roles(root: Path) -> dict[str, str]:
    """Map every tracked VBA component to its role from its documented location."""
    roles: dict[str, str] = {}
    for path in tracked_vba(root):
        for prefix, role in ROLE_BY_PREFIX:
            if path.startswith(prefix):
                roles[path] = role
                break
    return roles


def run_check(root: Path) -> dict[str, Any]:
    manifest = MANIFEST_PATH
    findings: list[dict[str, Any]] = []
    everything, parse_errors = _parse_components(root, component_roles(root))
    findings.extend(parse_errors)
    # Every standard module's Public names share one project namespace, Option
    # Private Module included: a clash anywhere does not compile.
    _, index_errors = _index_declarations(everything)
    findings.extend(index_errors)
    declarations = [item for item in everything if item["supported"] and item["kind"] != ENUM_MEMBER]
    keys, _ = _index_declarations(declarations)
    rows, signatures, errors = read_manifest(root, manifest)
    findings.extend(errors)
    findings.extend(_manifest_findings(manifest, keys, rows, signatures))

    evidence = sorted(
        declarations,
        key=lambda declaration: (
            str(declaration["component"]).casefold(),
            str(declaration["name"]).casefold(),
            str(declaration["kind"]).casefold(),
        ),
    )
    return {
        "schema_version": 1,
        "tool": TOOL_NAME,
        "status": "pass" if not findings else "fail",
        "manifest": manifest,
        "policy": {
            "public_role": "src/modules/ (docs/REPOSITORY_STRUCTURE.md)",
            "implicit_visibility": "prohibited",
            "const_and_variable_form": "one identifier per statement",
            "property_accessors": "same-name Get/Let/Set allowed only within one component",
        },
        "declarations": evidence,
        "findings": findings,
    }


def markdown_report(report: dict[str, Any]) -> str:
    lines = [
        "## VBA public API",
        "",
        f"- **Status:** {str(report['status']).upper()}",
        f"- **Manifest:** `{report['manifest']}`",
        f"- **Declarations:** {len(report['declarations'])}",
        f"- **Public role:** {report['policy']['public_role']}",
        "- **Implicit public visibility:** prohibited",
        "- **Public Const/variable form:** one identifier per statement",
        "- **Property accessors:** same-name Get/Let/Set allowed within one component",
        f"- **Findings:** {len(report['findings'])}",
    ]
    if report["declarations"]:
        lines += [
            "",
            "| Component | Kind | Name | Signature |",
            "| --- | --- | --- | --- |",
        ]
        for declaration in report["declarations"]:
            signature = str(declaration["signature"]).replace("|", "\\|")
            lines.append(
                f"| {declaration['component']} | {declaration['kind']} | "
                f"{declaration['name']} | `{signature}` |"
            )
    if report["findings"]:
        lines += ["", "### Findings", ""]
        for item in report["findings"]:
            location = str(item.get("path", "."))
            if item.get("line"):
                location += f":{item['line']}"
            lines.append(f"- `{location}` — {item['message']}")
    return "\n".join(lines) + "\n"


def init_fixture(
    root: Path, facade: str, manifest: list[str], other: str | None = None,
    extra: dict[str, str] | None = None,
) -> None:
    for path in (
        root / "src/modules",
        root / "src/core",
        root / "tests/modules",
        root / "docs",
    ):
        path.mkdir(parents=True, exist_ok=True)
    files = {
        "src/modules/Facade.bas": facade,
        "src/core/Core.bas": (
            'Attribute VB_Name = "Core"\n'
            "Option Explicit\n"
            "Option Private Module\n"
            "Public Function InternalOnly() As Long\n"
            "End Function\n"
        ),
        "tests/modules/Tests.bas": (
            'Attribute VB_Name = "Tests"\n'
            "Option Explicit\n"
            "Public Sub RunTests()\n"
            "End Sub\n"
        ),
    }
    if other is not None:
        files["src/modules/Other.bas"] = other
    files.update(extra or {})
    for relative, text in files.items():
        (root / relative).parent.mkdir(parents=True, exist_ok=True)
        (root / relative).write_bytes(text.replace("\n", "\r\n").encode("cp1252"))
    (root / "docs/PUBLIC_API.txt").write_text(
        "\n".join(manifest) + "\n", encoding="utf-8"
    )
    for command in (
        ("init", "-b", "main"),
        ("config", "user.name", "API Self-Test"),
        ("config", "user.email", "api@example.invalid"),
        ("add", "--all"),
    ):
        completed = git(root, *command)
        if completed.returncode:
            raise RuntimeError(
                completed.stderr.decode("utf-8", errors="replace").strip()
            )


def fixture(
    facade: str, manifest: list[str], other: str | None = None,
    extra: dict[str, str] | None = None,
) -> dict[str, Any]:
    with tempfile.TemporaryDirectory(prefix="vba-api-") as temporary:
        root = Path(temporary)
        init_fixture(root, facade, manifest, other, extra)
        return run_check(root)


def run_self_test() -> int:
    facade = '''Attribute VB_Name = "Facade"
Option Explicit
Public Const C As Long = 7
Public Declare PtrSafe Function Tick Lib "kernel32" () As Long
Public Number As Long
Public WithEvents Source As Object
Public Enum Mode
    ModeA = 1
End Enum
Public Type Pair
    Left As Long
    Right As Long
End Type
Public Function Echo( _
    ByVal value As Long) As Long
    Echo = value
End Function
Public Property Get Current() As Long
    Current = Number
End Property
Public Property Let Current(ByVal value As Long)
    Number = value
End Property
'''
    manifest = [
        "# Supported VBA declarations: <component><TAB><kind><TAB><name>",
        "Facade\tConst\tC",
        "# SIG\tFacade\tConst\tC\tPublic Const C As Long = 7",
        "Facade\tDeclare Function\tTick",
        '# SIG\tFacade\tDeclare Function\tTick\tPublic Declare PtrSafe Function Tick Lib "kernel32" () As Long',
        "Facade\tVariable\tNumber",
        "# SIG\tFacade\tVariable\tNumber\tPublic Number As Long",
        "Facade\tWithEvents Variable\tSource",
        "# SIG\tFacade\tWithEvents Variable\tSource\tPublic WithEvents Source As Object",
        "Facade\tEnum\tMode",
        "# SIG\tFacade\tEnum\tMode\tPublic Enum Mode | ModeA = 1 | End Enum",
        "Facade\tType\tPair",
        "# SIG\tFacade\tType\tPair\tPublic Type Pair | Left As Long | Right As Long | End Type",
        "Facade\tFunction\tEcho",
        "# SIG\tFacade\tFunction\tEcho\tPublic Function Echo( ByVal value As Long) As Long",
        "Facade\tProperty Get\tCurrent",
        "# SIG\tFacade\tProperty Get\tCurrent\tPublic Property Get Current() As Long",
        "Facade\tProperty Let\tCurrent",
        "# SIG\tFacade\tProperty Let\tCurrent\tPublic Property Let Current(ByVal value As Long)",
    ]
    failures: list[str] = []
    tests: list[tuple[str, dict[str, Any], str, str | None]] = []
    tests.append(("positive", fixture(facade, manifest), "pass", None))
    tests.append(
        (
            "implicit",
            fixture(facade.replace("Public Function Echo", "Function Echo"), manifest),
            "fail",
            "Implicit public procedure",
        )
    )
    tests.append(
        (
            "missing-signature",
            fixture(
                facade,
                [
                    item
                    for item in manifest
                    if not item.startswith("# SIG\tFacade\tFunction\tEcho")
                ],
            ),
            "fail",
            "Normalized signature record is missing",
        )
    )
    tests.append(
        (
            "multi-var",
            fixture(
                facade.replace(
                    "Public Number As Long",
                    "Public Number As Long, Other As Long",
                ),
                manifest,
            ),
            "fail",
            "Public variable declarations must contain one identifier",
        )
    )
    tests.append(
        (
            "multi-const",
            fixture(
                facade.replace(
                    "Public Const C As Long = 7",
                    "Public Const C As Long = 7, D As Long = 8",
                ),
                manifest,
            ),
            "fail",
            "Public Const declarations must contain one identifier",
        )
    )
    duplicate = facade + "\nPublic Function Echo(ByVal value As String) As String\nEnd Function\n"
    tests.append(
        ("same-component-collision", fixture(duplicate, manifest), "fail", "collides")
    )
    other = '''Attribute VB_Name = "Other"
Option Explicit
Public Function Echo(ByVal value As Long) As Long
    Echo = value
End Function
'''
    tests.append(
        (
            "cross-component-collision",
            fixture(facade, manifest, other),
            "fail",
            "collides",
        )
    )
    modern = 'Public Declare PtrSafe Function Tick Lib "kernel32" () As Long'
    legacy = modern.replace("PtrSafe ", "")
    conditional = facade.replace(modern, f"#If VBA7 Then\n{modern}\n#Else\n{legacy}\n#End If")
    variant_manifest = manifest + ["# SIG\tFacade\tDeclare Function\tTick\t" + legacy]
    tests.append(("conditional-declare", fixture(conditional, variant_manifest), "pass", None))
    tests.append(("conditional-missing-signature", fixture(conditional, manifest), "fail", "Signature mismatch"))
    tests.append(("conditional-stale-signature", fixture(facade, variant_manifest), "fail", "Signature mismatch"))
    tests.append(("duplicate-signature", fixture(facade, manifest + [manifest[6]]), "fail", "Duplicate signature"))
    overlap = conditional + "\n" + modern + "\n"
    tests.append(("reachable-collision", fixture(overlap, variant_manifest), "fail", "collides"))
    nested = facade.replace(modern, f"#If VBA7 Then\n#If Win64 Then\n{modern}\n#Else\n{modern}\n#End If\n#ElseIf VBA6 Then\n{legacy}\n#End If")
    tests.append(("nested-elseif", fixture(nested, variant_manifest), "pass", None))
    dead = facade + "\n#If False Then\n" + modern + "\n#End If\n"
    tests.append(("unreachable-declaration", fixture(dead, manifest), "pass", None))
    tests.append(("unknown-condition", fixture(conditional.replace("#If VBA7 Then", "#If Unknown Then"), variant_manifest), "fail", "Indeterminate"))
    tests.append(("unclosed-condition", fixture(conditional.replace("#End If", ""), variant_manifest), "fail", "unclosed"))
    alternate = 'Attribute VB_Name = "Other"\n#If VBA6 Then\n' + legacy + '\n#End If\n'
    complementary = facade.replace(modern, f"#If VBA7 Then\n{modern}\n#End If")
    cross_manifest = manifest + ["Other\tDeclare Function\tTick", "# SIG\tOther\tDeclare Function\tTick\t" + legacy]
    tests.append(("exclusive-cross-component", fixture(complementary, cross_manifest, alternate), "pass", None))
    tests.append(("overlapping-cross-component", fixture(facade, cross_manifest, alternate), "fail", "collides"))
    private = {"src/core/Second.bas": 'Attribute VB_Name = "Second"\nOption Explicit\nOption Private Module\n'
                                      "Public Function internalonly() As Long\nEnd Function\n"}
    tests.append(("private-module-collision", fixture(facade, manifest, extra=private), "fail", "collides"))
    tests.append(("private-module-distinct", fixture(facade, manifest, extra={
        "src/core/Second.bas": private["src/core/Second.bas"].replace("internalonly", "OtherInternal")}), "pass", None))
    exclusive = {"src/core/Second.bas": 'Attribute VB_Name = "Second"\nOption Explicit\nOption Private Module\n'
                                        "#If VBA7 Then\n#Else\nPublic Function InternalOnly() As Long\nEnd Function\n"
                                        "#End If\n"}
    core_vba7 = {"src/core/Core.bas": 'Attribute VB_Name = "Core"\nOption Explicit\nOption Private Module\n'
                                      "#If VBA7 Then\nPublic Function InternalOnly() As Long\nEnd Function\n#End If\n"}
    tests.append(("exclusive-private-modules", fixture(facade, manifest, extra={**exclusive, **core_vba7}), "pass", None))
    member = {"src/core/Second.bas": 'Attribute VB_Name = "Second"\nOption Explicit\nOption Private Module\n'
                                     "Public Const ModeA As Long = 2\n"}
    tests.append(("enum-member-collision", fixture(facade, manifest, extra=member), "fail", "collides"))
    global_enum = {"src/core/Second.bas": 'Attribute VB_Name = "Second"\nOption Explicit\nOption Private Module\n'
                                          "Global Enum Level\n    ModeA = 2\nEnd Enum\n"}
    tests.append(("global-enum-member-collision", fixture(facade, manifest, extra=global_enum), "fail", "collides"))
    tests.append(("global-enum-rejected", fixture(facade, manifest, extra={
        "src/core/Second.bas": global_enum["src/core/Second.bas"].replace("ModeA", "LevelA")}), "fail", "declare it Public"))
    tests.append(("standard-module-event-rejected", fixture(facade, manifest, extra={
        "src/core/Second.bas": 'Attribute VB_Name = "Second"\nOption Explicit\nOption Private Module\n'
                               "Public Event Ping()\n"}), "fail", "valid only in class and form modules"))
    class_enum = {"src/classes/Widget.cls": 'VERSION 1.0 CLASS\nBEGIN\n  MultiUse = -1  \'True\nEND\n'
                                            'Attribute VB_Name = "Widget"\nOption Explicit\n'
                                            "Public Enum Mode\n    WidgetA = 1\nEnd Enum\n"
                                            "Public Function Echo(ByVal value As Long) As Long\nEnd Function\n"}
    tests.append(("class-enum-collision", fixture(facade, manifest, extra=class_enum), "fail", "collides"))
    tests.append(("class-member-distinct", fixture(facade, manifest, extra={
        "src/classes/Widget.cls": class_enum["src/classes/Widget.cls"].replace("Enum Mode", "Enum WidgetMode")}),
        "pass", None))
    same_class = class_enum["src/classes/Widget.cls"].replace("Enum Mode", "Enum WidgetMode")
    tests.append(("class-enum-member-vs-variable", fixture(facade, manifest, extra={
        "src/classes/Widget.cls": same_class + "Public WidgetA As Long\n"}), "fail", "collides"))
    tests.append(("class-member-vs-member", fixture(facade, manifest, extra={
        "src/classes/Widget.cls": same_class + "Public Echo As Long\n"}), "fail", "collides"))
    tests.append(("class-event-vs-procedure", fixture(facade, manifest, extra={
        "src/classes/Widget.cls": same_class + "Public Event Echo()\n"}), "pass", None))
    tests.append(("class-event-vs-event", fixture(facade, manifest, extra={
        "src/classes/Widget.cls": same_class + "Public Event Fired()\nPublic Event fired(ByVal x As Long)\n"}),
        "fail", "collides"))
    tests.append(("global-type-rejected", fixture(facade, manifest, extra={
        "src/core/Second.bas": 'Attribute VB_Name = "Second"\nOption Explicit\nOption Private Module\n'
                               "Global Type Cell\n    Value As Long\nEnd Type\n"}), "fail", "declare it Public"))
    tests.append(("class-global-enum-rejected", fixture(facade, manifest, extra={
        "src/classes/Widget.cls": same_class.replace("Public Enum WidgetMode", "Global Enum WidgetMode")}),
        "fail", "declare it Public"))
    for name, report, expected, needle in tests:
        if report["status"] != expected:
            failures.append(f"{name}: expected {expected}, got {report['status']}")
        if needle and not any(
            needle in item["message"] for item in report["findings"]
        ):
            failures.append(f"{name}: missing diagnostic {needle!r}")
    if failures:
        for failure in failures:
            print(f"[FAIL] {failure}")
        print(f"SELF-TEST FAIL: {len(failures)} failure(s).")
        return 1
    print(
        "SELF-TEST PASS: procedures, paired properties, constants, events, declares, "
        "variables, WithEvents, enums, types, continuations, implicit-public rejection, "
        "signature drift, conditional variants, reachable collisions across every standard module (Option Private Module, Global and enum members included) and with class-module enums, names within a class module (events apart), Global Enum/Type and standard-module events rejected, unknown conditions, and one-identifier Const/variable policy passed."
    )
    return 0


def main(argv: list[str] | None = None) -> int:
    options = parse_args(sys.argv[1:] if argv is None else argv)
    return run_gate(
        options,
        build=lambda: run_check(options.root),
        markdown=markdown_report,
        errors=(
            OSError,
            UnicodeError,
            RuntimeError,
            subprocess.SubprocessError,
        ),
        self_test=run_self_test,
    )

if __name__ == "__main__":
    raise SystemExit(main())
