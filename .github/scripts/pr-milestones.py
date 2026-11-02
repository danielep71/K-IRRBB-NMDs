"""Remove milestones from pull requests: in this repository milestones belong to issues.

On a pull-request event only that pull request is checked; on a schedule, a manual
run or a push, every pull request is. Each item is read back through the issues
endpoint and changed only if it is a pull request, so an issue is never touched.
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


def clear(repository: str, number: int) -> bool:
    """Remove the milestone of pull request `number`; return whether one was removed."""
    endpoint = f"repos/{repository}/issues/{number}"
    item = api(endpoint)
    if "pull_request" not in item:
        raise RuntimeError(f"#{number} is an issue, not a pull request; refusing to change it")
    if item.get("milestone") is None:
        return False
    api(endpoint, "PATCH", {"milestone": None})
    if api(endpoint).get("milestone") is not None:
        raise RuntimeError(f"pull request #{number} still has a milestone after the update")
    print(f"Removed the milestone from pull request #{number}")
    return True


def reconcile(repository: str, number: int | None) -> int:
    """Clear the milestone of one pull request, or of every pull request when `number` is None."""
    if number is not None:
        return int(clear(repository, number))
    removed = 0
    # Oldest first and unfiltered, so clearing a milestone cannot shift pages still to read.
    for page in range(1, MAX_PAGES + 1):
        items = api(f"repos/{repository}/pulls?state=all&sort=created&direction=asc"
                    f"&per_page={PAGE_SIZE}&page={page}")
        if not isinstance(items, list):
            raise RuntimeError("the pull-request list is not a JSON array")
        for item in items:
            if item.get("milestone") is not None:
                removed += clear(repository, int(item["number"]))
        if len(items) < PAGE_SIZE:
            print(f"Checked every pull request; removed {removed} milestone(s)")
            return removed
    raise RuntimeError(f"more than {MAX_PAGES} pages of pull requests; nothing beyond them was checked")


def main() -> int:
    number = os.environ.get("PR_NUMBER", "").strip()
    reconcile(os.environ["GITHUB_REPOSITORY"], int(number) if number else None)
    return 0


if __name__ == "__main__":
    sys.exit(main())
