#!/usr/bin/env python3
"""Write host.json for an Excel evidence bundle and validate it.

Usage: python tools/write_excel_evidence.py --candidate-sha SHA --bundle FOLDER [--root CHECKOUT]

tools/Certify-IRRBB.ps1 runs this once the logs and session.json are in the
bundle. It judges every stage from the logs (tools/excel_evidence.py), writes
host.json, validates the bundle as tools/check_excel_evidence.py does and
prints one line per stage. A host.json that exists already is never replaced.
Exit code 0 only when the bundle validates as PASS. It never runs Excel.
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from excel_evidence import EvidenceError, record, report, validate


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--root", type=Path, default=Path(__file__).resolve().parents[1])
    parser.add_argument("--candidate-sha", required=True)
    parser.add_argument("--bundle", type=Path, required=True)
    args = parser.parse_args(argv)
    host_path = args.bundle / "host.json"
    if host_path.exists():
        print(f"ERROR: {host_path} exists already; a record is never replaced.", file=sys.stderr)
        return 2
    try:
        host = record(args.root.resolve(), args.candidate_sha, args.bundle)
    except EvidenceError as error:
        print(f"ERROR: {error}", file=sys.stderr)
        return 1
    host_path.write_text(json.dumps(host, indent=2) + "\n", encoding="utf-8")
    print("\n".join(report(host)))
    result, findings = validate(args.root.resolve(), args.bundle, args.candidate_sha)
    print(f"Validation: {result}" + ("" if not findings else f" ({len(findings)} findings)"))
    return 0 if result == "PASS" else 1


if __name__ == "__main__":
    sys.exit(main())
