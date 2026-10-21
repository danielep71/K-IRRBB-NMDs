#!/usr/bin/env python3
"""Run the portable K-IRRBB source gates, locally and in CI (never runs Excel)."""
from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path
from typing import Any

from _gatelib import git_text, write_json, write_text


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=Path(__file__).resolve().parents[1])
    parser.add_argument("--output-dir", type=Path, default=Path("test-results"))
    parser.add_argument("--ci", action="store_true", help="Check committed range and require clean checkout")
    parser.add_argument("--base", help="Optional PR base SHA")
    args = parser.parse_args()
    root = args.root.resolve()
    output = (root / args.output_dir).resolve()
    output.mkdir(parents=True, exist_ok=True)
    sha = git_text(root, "rev-parse", "HEAD", check=True).stdout.strip()
    dirty = bool(git_text(root, "status", "--porcelain", "--untracked-files=normal", check=True).stdout)
    commands = [
        ("tool-tests", [sys.executable, "-m", "unittest", "discover", "-s", "tools", "-p", "test_*.py"]),
    ]
    for name in ("check_committed_whitespace", "check_vba_jumps", "check_vba_conditionals", "check_vba_public_api"):
        commands.append((name + "-self-test", [sys.executable, "tools/" + name + ".py", "--self-test"]))
    for name in ("check_source", "check_committed_whitespace", "check_vba_jumps", "check_vba_conditionals", "check_vba_public_api"):
        command = [sys.executable, "tools/" + name + ".py", "--root", str(root), "--output", str(output / (name + ".json"))]
        if name == "check_committed_whitespace":
            command += ["--mode", "committed" if args.ci else "working-tree"]
            if args.ci and args.base:
                command += ["--base", args.base]
        commands.append((name, command))
    results = []
    for name, command in commands:
        r = subprocess.run(command, cwd=root, capture_output=True, text=True, check=False)
        log = r.stdout + r.stderr
        write_text(output / (name + ".log"), log)
        results.append({"name": name, "exit_code": r.returncode})
        print(name + ": " + ("PASS" if r.returncode == 0 else "FAIL"), flush=True)
        if r.returncode:
            print(log, flush=True)
    passed = all(r["exit_code"] == 0 for r in results) and (not args.ci or not dirty)
    if args.ci and dirty:
        print("FAIL: CI candidate must have a clean working tree", flush=True)
    report: dict[str, Any] = {
        "schema_version": 1, "candidate_sha": sha, "dirty": dirty,
        "status": "pass" if passed else "fail", "checks": results,
        "scope": "Static source and repository checks; Excel and IRRBB models NOT executed."
    }
    write_json(output / "static-checks.json", report)
    write_text(output / "static-checks.md", "# Static checks\n\nStatus: **" + report["status"].upper() + "**\n\nCandidate: `" + sha + "`\n\n" + report["scope"] + "\n")
    return 0 if passed else 1


if __name__ == "__main__":
    raise SystemExit(main())
