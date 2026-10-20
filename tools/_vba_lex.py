"""Shared VBA comment/continuation handling; not a full VBA grammar.

MS-VBAL 3.2.2 / 3.3.1: comments can span physical lines. WSC excludes
CP2 NBSP; underscore must be immediately before the line terminator.
"""
from __future__ import annotations

import re

CONTINUATION = re.compile(r"[\t\x19 \u1680\u2000-\u200a\u202f\u205f\u3000]_\Z")
REM = re.compile(r"Rem(?=\s|$)", re.I)


def physical_lines(lines: list[str], *, mask_strings: bool = False) -> list[tuple[int, str, bool]]:
    """Keep physical line numbers, remove comments, optionally blank literals."""
    result = []
    continued_comment = False
    for number, raw in enumerate(lines, 1):
        line = raw.removesuffix("\r")
        if continued_comment:
            continued_comment = bool(CONTINUATION.search(line))
            result.append((number, "", False))
            continue
        output = []
        quoted = False
        statement_start = True
        comment = False
        index = 0
        while index < len(line):
            char = line[index]
            if not quoted and (char == "'" or (statement_start and REM.match(line, index))):
                comment = True
                break
            if char == '"':
                if quoted and line[index:index + 2] == '""':
                    output.append("  " if mask_strings else '""')
                    index += 2
                    continue
                quoted = not quoted
                output.append(" " if mask_strings else char)
            else:
                output.append(" " if quoted and mask_strings else char)
            if not quoted:
                if char == ":" and line[index:index + 2] != ":=":
                    statement_start = True
                elif not char.isspace():
                    statement_start = False
            index += 1
        continued_comment = comment and bool(CONTINUATION.search(line))
        continued_code = not comment and not quoted and bool(CONTINUATION.search(line))
        result.append((number, "".join(output), continued_code))
    return result


def logical_lines(lines: list[str], *, mask_strings: bool = False) -> list[tuple[int, int, str]]:
    result = []
    buffer: list[str] = []
    start = 1
    for number, code, continued in physical_lines(lines, mask_strings=mask_strings):
        if not buffer:
            start = number
        buffer.append(code[:-1] if continued else code)
        if not continued:
            result.append((start, number, " ".join(part.strip() for part in buffer)))
            buffer.clear()
    if buffer:
        result.append((start, len(lines), " ".join(part.strip() for part in buffer)))
    return result
