#!/usr/bin/env python3
"""Validate an Excel evidence bundle offline against its candidate commit.

Usage: python tools/check_excel_evidence.py --bundle FOLDER [--candidate-sha SHA] [--root CHECKOUT]

Reads host.json, judges every stage again from the bundle's logs and the
repository at the candidate commit (tools/excel_evidence.py), recomputes the
source and log digests and compares all of it with what host.json records.
Prints every finding and RESULT=PASS or RESULT=FAIL; exit code 0 only for
PASS. --candidate-sha, when given, must be the commit host.json certifies.
The root must be a Git clone that contains the candidate commit. It never runs
Excel and cannot prove what a person did; NOT_RUN is never a pass.
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

from excel_evidence import validate


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--root", type=Path, default=Path(__file__).resolve().parents[1])
    parser.add_argument("--candidate-sha")
    parser.add_argument("--bundle", type=Path, required=True)
    args = parser.parse_args(argv)
    result, findings = validate(args.root.resolve(), args.bundle, args.candidate_sha)
    for finding in findings:
        print(f"- {finding}")
    print(f"RESULT={result}")
    return 0 if result == "PASS" else 1


if __name__ == "__main__":
    sys.exit(main())
