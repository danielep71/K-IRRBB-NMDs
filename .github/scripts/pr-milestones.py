"""Remove milestones from pull requests: in this repository milestones belong to issues.

On a pull-request event only that pull request is checked; on a schedule, a manual
run or a push, every pull request is. Each item is read back through the issues
endpoint and changed only if it is a pull request, so an issue is never touched.
When GITHUB_STEP_SUMMARY is set, the job summary lists each pull request and the
title of the milestone removed from it.
"""
from __future__ import annotations

import json
import os
import re
import subprocess
import sys
from typing import Any

SECRET = re.compile(r"\b(?:gh[pousr]_[A-Za-z0-9]{20,}|github_pat_[A-Za-z0-9_]{20,})")
DETAIL_LIMIT = 500
PAGE_SIZE = 100
MAX_PAGES = 100


def failure_detail(text: str) -> str:
    """Return API error text safe for a public log: token-like strings redacted, length capped."""
    detail = SECRET.sub("[REDACTED]", " ".join(text.split()))
    return detail if len(detail) <= DETAIL_LIMIT else detail[:DETAIL_LIMIT] + "..."


def table_cell(text: str) -> str:
    """Return `text` as one Markdown table cell: backslashes and pipes escaped, line breaks spaced."""
    escaped = text.replace("\\", "\\\\").replace("|", "\\|")
    return escaped.replace("\r\n", " ").replace("\r", " ").replace("\n", " ")


def summary(removed: list[tuple[int, str]]) -> str:
    """Return the job-summary Markdown for the (pull request, milestone title) pairs removed."""
    lines = ["## Pull-request milestones", ""]
    if not removed:
        lines.append("No milestone was removed from any pull request.")
    else:
        lines += ["| Pull request | Milestone removed |", "| --- | --- |"]
        lines += [f"| #{number} | {table_cell(title)} |" for number, title in removed]
    return "\n".join(lines) + "\n"


def api(endpoint: str, method: str = "GET", payload: Any = None) -> Any:
    command = ["gh", "api", endpoint, "--method", method]
    if payload is not None:
        command += ["--input", "-"]
    completed = subprocess.run(command, input=json.dumps(payload) if payload is not None else None,
                               capture_output=True, text=True)
    if completed.returncode != 0:
        detail = failure_detail(f"{completed.stderr} {completed.stdout}")
        raise RuntimeError(f"{method} {endpoint} failed (gh exit {completed.returncode}): {detail}")
    return json.loads(completed.stdout) if completed.stdout.strip() else None


def clear(repository: str, number: int) -> str | None:
    """Remove the milestone of pull request `number`; return its title, or None if it had none."""
    endpoint = f"repos/{repository}/issues/{number}"
    item = api(endpoint)
    if "pull_request" not in item:
        raise RuntimeError(f"#{number} is an issue, not a pull request; refusing to change it")
    milestone = item.get("milestone")
    if milestone is None:
        return None
    api(endpoint, "PATCH", {"milestone": None})
    if api(endpoint).get("milestone") is not None:
        raise RuntimeError(f"pull request #{number} still has a milestone after the update")
    print(f"Removed the milestone from pull request #{number}")
    return str(milestone.get("title") or "")


def reconcile(repository: str, number: int | None) -> list[tuple[int, str]]:
    """Clear the milestone of one pull request, or of every pull request when `number` is None.

    Return the (pull request, milestone title) pairs removed, in the order they were removed.
    """
    if number is not None:
        title = clear(repository, number)
        return [] if title is None else [(number, title)]
    removed: list[tuple[int, str]] = []
    # Oldest first and unfiltered, so clearing a milestone cannot shift pages still to read.
    for page in range(1, MAX_PAGES + 1):
        items = api(f"repos/{repository}/pulls?state=all&sort=created&direction=asc"
                    f"&per_page={PAGE_SIZE}&page={page}")
        if not isinstance(items, list):
            raise RuntimeError("the pull-request list is not a JSON array")
        for item in items:
            if item.get("milestone") is not None:
                pull_number = int(item["number"])
                title = clear(repository, pull_number)
                if title is not None:
                    removed.append((pull_number, title))
        if len(items) < PAGE_SIZE:
            print(f"Checked every pull request; removed {len(removed)} milestone(s)")
            return removed
    raise RuntimeError(f"more than {MAX_PAGES} pages of pull requests; nothing beyond them was checked")


def main() -> int:
    number = os.environ.get("PR_NUMBER", "").strip()
    removed = reconcile(os.environ["GITHUB_REPOSITORY"], int(number) if number else None)
    summary_path = os.environ.get("GITHUB_STEP_SUMMARY", "").strip()
    if summary_path:
        with open(summary_path, "a", encoding="utf-8") as handle:
            handle.write(summary(removed))
    return 0


if __name__ == "__main__":
    sys.exit(main())
