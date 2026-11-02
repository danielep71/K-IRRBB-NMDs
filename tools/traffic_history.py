#!/usr/bin/env python3
"""Merge GitHub traffic API responses into the CSV history of the traffic-history branch.

The traffic API keeps 14 days, so the Daily traffic export workflow collects
it once a day and this module merges each collection into data/ on the
`traffic-history` branch:

- data/traffic.csv holds one 14-day snapshot per collection date, with the
  repository's stars, forks, watchers and open issues and pull requests;
- data/traffic_daily.csv holds one row of views and clones per day;
- data/referrers.csv and data/popular_paths.csv hold each collection date's
  top referrers and popular paths, ranked as the API returned them;
- data/badges/{clones,visitors,visits}.json are shields.io endpoint files for
  the README, rebuilt from the stored data after every merge.

A collection date's snapshot, referrers and popular paths form one unit: a
later collection on the same day replaces the unit, and a collection replayed
after a later one of that day leaves it unchanged. A collection's days replace
the stored daily rows unless a later collection is already stored, in which
case it only adds the days that are missing. Every response and every
existing file is read and checked before anything is written.

The alert command picks the open milestone with the earliest due date after
today, then compares the two latest snapshots and, on a views spike, more
stars or forks, or a new referrer, writes the title, body and milestone of one
alert issue for the workflow to open. No milestone fails the run, even when no
alert is due.
"""

from __future__ import annotations

import argparse
import csv
import io
import json
import os
from collections.abc import Callable, Mapping, Sequence
from dataclasses import dataclass
from datetime import UTC, date, datetime, timedelta
from pathlib import Path
from typing import Any

TRAFFIC = "traffic.csv"
DAILY = "traffic_daily.csv"
REFERRERS = "referrers.csv"
PATHS = "popular_paths.csv"
HEADERS: dict[str, tuple[str, ...]] = {
    TRAFFIC: ("snapshot_date", "snapshot_datetime_utc", "views_14d", "unique_views_14d",
              "clones_14d", "unique_clones_14d", "stars", "forks", "watchers",
              "open_issues_and_prs"),
    DAILY: ("date", "views", "unique_views", "clones", "unique_clones"),
    REFERRERS: ("snapshot_date", "rank", "referrer", "count", "uniques"),
    PATHS: ("snapshot_date", "rank", "path", "title", "count", "uniques"),
}
RESPONSES = ("views.json", "clones.json", "referrers.json", "paths.json", "repository.json")
BADGES = "badges"
BADGE_COLOR = "217346"
Row = dict[str, str]


class TrafficError(RuntimeError):
    """An invalid response, an unreadable history file or no milestone to file under."""


# --- Reading and writing ---------------------------------------------------------


def to_int(value: Any, default: int = 0) -> int:
    """A count as the API or a CSV cell gives it; anything unreadable is the default."""
    try:
        return int(value)
    except (TypeError, ValueError):
        return default


def read_csv(path: Path) -> list[Row]:
    """Every row of a CSV file, or none when it does not exist; missing cells are blank."""
    if not path.exists():
        return []
    try:
        with path.open("r", newline="", encoding="utf-8") as handle:
            return [{key: value or "" for key, value in row.items() if key is not None}
                    for row in csv.DictReader(handle)]
    except (OSError, ValueError, csv.Error) as error:
        raise TrafficError(f"{path.name}: unreadable ({error})") from error


def render_csv(rows: Sequence[Mapping[str, str]], fieldnames: Sequence[str]) -> str:
    """The file's text: the header, then each row's known columns, CRLF-terminated."""
    buffer = io.StringIO()
    writer = csv.DictWriter(buffer, fieldnames=fieldnames)
    writer.writeheader()
    for row in rows:
        writer.writerow({field: row.get(field, "") for field in fieldnames})
    return buffer.getvalue()


def write_all(data: Path, contents: Mapping[str, str]) -> None:
    """Write every file to a temporary sibling first, then move them into place."""
    staged = []
    for name, content in contents.items():
        target = data / name
        target.parent.mkdir(parents=True, exist_ok=True)
        temporary = target.with_name(f".{target.name}.tmp")
        temporary.write_text(content, encoding="utf-8", newline="")
        staged.append((temporary, target))
    for temporary, target in staged:
        os.replace(temporary, target)


def load_json(path: Path) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except FileNotFoundError as error:
        raise TrafficError(f"{path.name}: response missing") from error
    except (OSError, ValueError) as error:
        raise TrafficError(f"{path.name}: unreadable response ({error})") from error


# --- Collection ------------------------------------------------------------------


@dataclass(frozen=True)
class Collection:
    """The five API responses of one run, and when they were collected."""

    collected_at: datetime
    views: dict[str, Any]
    clones: dict[str, Any]
    referrers: list[dict[str, Any]]
    paths: list[dict[str, Any]]
    repository: dict[str, Any]

    @property
    def snapshot_date(self) -> str:
        return self.collected_at.date().isoformat()

    @property
    def snapshot_datetime(self) -> str:
        return self.collected_at.isoformat(timespec="seconds")


def checked_series(data: Any, key: str) -> dict[str, Any]:
    if not isinstance(data, dict) or key not in data or "count" not in data:
        raise TrafficError(f"Unexpected {key} API response")
    if not isinstance(data[key], list) or not all(isinstance(item, dict) for item in data[key]):
        raise TrafficError(f"Unexpected {key} API response: {key!r} is not a list of objects")
    return data


def checked_ranking(data: Any, name: str) -> list[dict[str, Any]]:
    if not isinstance(data, list) or not all(isinstance(item, dict) for item in data):
        raise TrafficError(f"Unexpected {name} API response")
    return data


def load_collection(raw: Path, collected_at: datetime) -> Collection:
    """Check the five responses of one collection; all of them or none."""
    repository = load_json(raw / "repository.json")
    if not isinstance(repository, dict) or "stargazers_count" not in repository:
        raise TrafficError("Unexpected repository metadata API response")
    return Collection(
        collected_at=collected_at.astimezone(UTC),
        views=checked_series(load_json(raw / "views.json"), "views"),
        clones=checked_series(load_json(raw / "clones.json"), "clones"),
        referrers=checked_ranking(load_json(raw / "referrers.json"), "referrers"),
        paths=checked_ranking(load_json(raw / "paths.json"), "popular paths"),
        repository=repository,
    )


# --- Merge rules -----------------------------------------------------------------


def normalize_traffic_row(row: Mapping[str, str]) -> Row | None:
    """A traffic.csv row in the current layout; earlier layouts are converted.

    The earliest layout keyed rows by ``date`` with ``views``/``clones``
    columns; a later one named the repository count ``open_issues``. A row
    with neither date column is dropped.
    """
    open_issues = row.get("open_issues_and_prs", row.get("open_issues", "0"))
    counts = {key: row.get(key, "0") for key in ("stars", "forks", "watchers")}
    if row.get("snapshot_date"):
        return {
            "snapshot_date": row.get("snapshot_date", ""),
            "snapshot_datetime_utc": row.get("snapshot_datetime_utc", ""),
            "views_14d": row.get("views_14d", "0"),
            "unique_views_14d": row.get("unique_views_14d", "0"),
            "clones_14d": row.get("clones_14d", "0"),
            "unique_clones_14d": row.get("unique_clones_14d", "0"),
            **counts, "open_issues_and_prs": open_issues,
        }
    if row.get("date"):
        return {
            "snapshot_date": row.get("date", "")[:10],
            "snapshot_datetime_utc": row.get("date", ""),
            "views_14d": row.get("views", "0"),
            "unique_views_14d": row.get("unique_views", "0"),
            "clones_14d": row.get("clones", "0"),
            "unique_clones_14d": row.get("unique_clones", "0"),
            **counts, "open_issues_and_prs": open_issues,
        }
    return None


def upsert(existing: Sequence[Row], new: Sequence[Row], key_fields: Sequence[str],
           normalizer: Callable[[Mapping[str, str]], Row | None] | None = None) -> list[Row]:
    """Replace the stored rows that share a key with a new row, then sort by key.

    Rows with a blank key column are dropped.
    """
    def key(row: Mapping[str, str]) -> tuple[str, ...]:
        return tuple(str(row.get(field, "")) for field in key_fields)

    rows = list(existing)
    if normalizer is not None:
        rows = [normalized for normalized in map(normalizer, rows) if normalized is not None]
    new_keys = {key(row) for row in new}
    combined = [row for row in rows if key(row) not in new_keys] + list(new)
    combined = [row for row in combined if all(str(row.get(field, "")) for field in key_fields)]
    return sorted(combined, key=key)


def collected(row: Mapping[str, str]) -> datetime | None:
    """When a snapshot was collected, or None when its time is blank or unreadable."""
    try:
        stamp = datetime.fromisoformat(row.get("snapshot_datetime_utc", ""))
    except ValueError:
        return None
    return stamp if stamp.tzinfo else stamp.replace(tzinfo=UTC)


def traffic_row(collection: Collection) -> Row:
    views, clones, repository = collection.views, collection.clones, collection.repository
    return {
        "snapshot_date": collection.snapshot_date,
        "snapshot_datetime_utc": collection.snapshot_datetime,
        "views_14d": str(to_int(views.get("count", 0))),
        "unique_views_14d": str(to_int(views.get("uniques", 0))),
        "clones_14d": str(to_int(clones.get("count", 0))),
        "unique_clones_14d": str(to_int(clones.get("uniques", 0))),
        "stars": str(to_int(repository.get("stargazers_count", 0))),
        "forks": str(to_int(repository.get("forks_count", 0))),
        "watchers": str(to_int(repository.get("subscribers_count", 0))),
        "open_issues_and_prs": str(to_int(repository.get("open_issues_count", 0))),
    }


def daily_rows(collection: Collection) -> list[Row]:
    """One row per day either series reports; a metric a series omits is zero."""
    daily: dict[str, Row] = {}
    for prefix, series in (("views", collection.views), ("clones", collection.clones)):
        for item in series.get(prefix, []):
            day = str(item.get("timestamp", ""))[:10]
            if not day:
                continue
            row = daily.setdefault(day, {"date": day, "views": "0", "unique_views": "0",
                                         "clones": "0", "unique_clones": "0"})
            row[prefix] = str(to_int(item.get("count", 0)))
            row[f"unique_{prefix}"] = str(to_int(item.get("uniques", 0)))
    return list(daily.values())


def cell(value: Any) -> str:
    """A text field as the CSV writer stores it: None is blank."""
    return "" if value is None else str(value)


def ranked_rows(collection: Collection, items: Sequence[Mapping[str, Any]],
                fields: Sequence[str]) -> list[Row]:
    rows = []
    for rank, item in enumerate(items, start=1):
        row = {"snapshot_date": collection.snapshot_date, "rank": str(rank)}
        row.update({field: cell(item.get(field, "")) for field in fields})
        row["count"] = str(to_int(item.get("count", 0)))
        row["uniques"] = str(to_int(item.get("uniques", 0)))
        rows.append(row)
    return rows


def replace_day(existing: Sequence[Row], day: str, new: Sequence[Row]) -> list[Row]:
    """Replace one snapshot date's ranked rows; other dates keep theirs."""
    combined = [row for row in existing if row.get("snapshot_date") != day] + list(new)
    return sorted(combined, key=lambda row: (row.get("snapshot_date", ""),
                                             to_int(row.get("rank", 999))))


@dataclass
class Outcome:
    """What a merge did: the merged rows and whether the collection was applied."""

    rows: dict[str, list[Row]]
    day_replaced: bool
    days_replaced: bool


def merge_rows(current: Mapping[str, Sequence[Row]], collection: Collection) -> Outcome:
    """Merge one collection into the stored rows of the four history files.

    The collection date's snapshot, referrers and popular paths are replaced
    together unless the stored snapshot of that date was collected later. The
    collection's days replace the stored daily rows unless any stored
    snapshot was collected later; then only days not yet stored are added.
    """
    day = collection.snapshot_date
    stamp = collection.collected_at
    traffic = upsert(current[TRAFFIC], [], ["snapshot_date"], normalize_traffic_row)
    times = [(row["snapshot_date"], collected(row)) for row in traffic]
    stored_day = [time for snapshot, time in times if snapshot == day]
    day_replaced = not stored_day or stored_day[0] is None or stored_day[0] <= stamp
    days_replaced = all(time is None or time <= stamp for _, time in times)

    rows = {name: list(current[name]) for name in (DAILY, REFERRERS, PATHS)}
    rows[TRAFFIC] = traffic
    if day_replaced:
        rows[TRAFFIC] = upsert(traffic, [traffic_row(collection)], ["snapshot_date"])
        rows[REFERRERS] = replace_day(rows[REFERRERS], day, ranked_rows(
            collection, collection.referrers, ("referrer",)))
        rows[PATHS] = replace_day(rows[PATHS], day, ranked_rows(
            collection, collection.paths, ("path", "title")))
    new_days = daily_rows(collection)
    if not days_replaced:
        stored = {row.get("date", "") for row in rows[DAILY]}
        new_days = [row for row in new_days if row["date"] not in stored]
    rows[DAILY] = upsert(rows[DAILY], new_days, ["date"])
    return Outcome(rows, day_replaced, days_replaced)


# --- Badges ----------------------------------------------------------------------


def compact_count(value: int) -> str:
    """A badge number: 950, 1.2k, 11k, 3.4M."""
    if value >= 1000000:
        return f"{value / 1000000:.1f}M".replace(".0M", "M")
    if value >= 1000:
        return f"{value / 1000:.1f}k".replace(".0k", "k")
    return str(value)


def badges(daily: Sequence[Mapping[str, str]],
           traffic: Sequence[Mapping[str, str]]) -> dict[str, str]:
    """shields.io endpoint files for the README, from the stored data only.

    clones: total clones stored in the daily history since tracking began.
    visitors (14d): unique visitors over the API's rolling 14-day window,
    from the latest snapshot; daily uniques cannot be summed into a unique
    count. visits: daily unique visitors summed since tracking began, so a
    visitor counts once per day they visit.
    """
    latest = max(traffic, key=lambda row: row.get("snapshot_date", ""), default=None)
    values = {
        "clones": ("clones", sum(to_int(row.get("clones", 0)) for row in daily)),
        "visits": ("visits", sum(to_int(row.get("unique_views", 0)) for row in daily)),
        "visitors": ("visitors (14d)",
                     to_int(latest.get("unique_views_14d", 0)) if latest else 0),
    }
    return {f"{BADGES}/{name}.json": json.dumps(
                {"schemaVersion": 1, "label": label, "message": compact_count(value),
                 "color": BADGE_COLOR}) + "\n"
            for name, (label, value) in values.items()}


def merge(raw: Path, data: Path, collected_at: datetime) -> Outcome:
    """Check the collection and the history, then rewrite the history files and badges."""
    collection = load_collection(raw, collected_at)
    current = {name: read_csv(data / name) for name in HEADERS}
    outcome = merge_rows(current, collection)
    contents = {name: render_csv(rows, HEADERS[name]) for name, rows in outcome.rows.items()}
    contents.update(badges(outcome.rows[DAILY], outcome.rows[TRAFFIC]))
    write_all(data, contents)
    return outcome


# --- Traffic alert ---------------------------------------------------------------

NEW_REFERRER_VIEWS = 5
ALERT_DAYS = 7


def active_milestone(milestones: Sequence[Mapping[str, Any]], today: str) -> Mapping[str, Any]:
    """The open milestone with the earliest due date after today (UTC)."""
    dated = [m for m in milestones if m.get("state") == "open" and m.get("due_on")
             and str(m["due_on"])[:10] > today]
    if not dated:
        raise TrafficError("No open milestone is due after today; give the next "
                           "milestone a due date so the alert can be filed under it")
    return min(dated, key=lambda m: str(m["due_on"]))


def target_end_date(today: str, milestone_due: str) -> date:
    """A week ahead, or the day before the milestone's due date if that is sooner."""
    return min(date.fromisoformat(today) + timedelta(days=ALERT_DAYS),
               date.fromisoformat(milestone_due) - timedelta(days=1))


def top_by_count(rows: Sequence[Row], limit: int = 10) -> list[Row]:
    return sorted(rows, key=lambda row: to_int(row.get("count", 0)), reverse=True)[:limit]


def reasons_for(previous: Mapping[str, str], current: Mapping[str, str],
                new_referrers: Sequence[Row]) -> list[str]:
    """Why the latest snapshot stands out; clones are reported but never a reason."""
    reasons = []
    before = to_int(previous.get("views_14d"))
    if to_int(current.get("views_14d")) >= max(before * 2, before + 100):
        reasons.append("Views spike")
    if to_int(current.get("stars")) > to_int(previous.get("stars")):
        reasons.append("Stars increased")
    if to_int(current.get("forks")) > to_int(previous.get("forks")):
        reasons.append("Forks increased")
    if new_referrers:
        reasons.append("New referrer detected")
    return reasons


def listed(rows: Sequence[Row], label: str, empty: str) -> str:
    return "\n".join(
        f"- {row.get(label, '')}: {row.get('count', '0')} views, {row.get('uniques', '0')} unique"
        for row in top_by_count(rows)) or empty


def alert_body(previous: Mapping[str, str], current: Mapping[str, str], reasons: Sequence[str],
               lists: Sequence[str], target: date) -> str:
    since = f"on {previous['snapshot_date']}"
    new_referrers, referrers, paths = lists

    def compared(label: str, key: str) -> str:
        return f"- {label}: {current[key]} vs {previous[key]} {since}"

    return "\n".join([
        "Automated traffic alert for **K-IRRBB**.",
        "",
        "Trigger reason:",
        ", ".join(reasons),
        "",
        "Current 14-day snapshot:",
        compared("Views", "views_14d"),
        compared("Unique views", "unique_views_14d"),
        compared("Clones", "clones_14d"),
        compared("Unique clones", "unique_clones_14d"),
        compared("Stars", "stars"),
        compared("Forks", "forks"),
        "",
        "New referrers:",
        new_referrers,
        "",
        "Current top referrers:",
        referrers,
        "",
        "Current popular paths:",
        paths,
        "",
        "Suggested checks:",
        "1. Go to **Insights → Traffic**",
        "2. Check whether one of the referrers is a blog, forum, Reddit, StackOverflow, "
        "GitHub project, or newsletter",
        "3. Check which page, release, screenshot, source file, or README section is "
        "attracting traffic",
        "4. Consider adding a short README CTA asking users to star the repository if they "
        "are using the project",
        "",
        "## Dependencies",
        "",
        "None.",
        "",
        f"- Target end date: {target} (Europe/Rome)",
    ])


def alert(data: Path, milestone: Mapping[str, Any], today: str) -> dict[str, str] | None:
    """The alert issue for the latest snapshot, or None when nothing stands out.

    It compares the latest snapshot with the one before: a views spike (the
    14-day views at least double and grow by 100), more stars or forks, or a
    referrer never seen before that brings at least NEW_REFERRER_VIEWS views.
    """
    traffic = sorted((row for row in read_csv(data / TRAFFIC) if row.get("snapshot_date")),
                     key=lambda row: row.get("snapshot_date", ""))
    if len(traffic) < 2:
        return None
    previous, current = traffic[-2], traffic[-1]
    day = current["snapshot_date"]
    referrers = read_csv(data / REFERRERS)
    current_referrers = [row for row in referrers if row.get("snapshot_date") == day]
    known = {row.get("referrer", "") for row in referrers if row.get("snapshot_date", "") < day}
    new_referrers = [row for row in current_referrers
                     if row.get("referrer", "") and row.get("referrer", "") not in known
                     and to_int(row.get("count", 0)) >= NEW_REFERRER_VIEWS]
    reasons = reasons_for(previous, current, new_referrers)
    if not reasons:
        return None
    current_paths = [row for row in read_csv(data / PATHS) if row.get("snapshot_date") == day]
    lists = (listed(new_referrers, "referrer", "- No new referrer above threshold"),
             listed(current_referrers, "referrer", "- No referrers available"),
             listed(current_paths, "path", "- No popular paths available"))
    target = target_end_date(today, str(milestone["due_on"])[:10])
    return {"title": f"Traffic alert: {day}",
            "body": alert_body(previous, current, reasons, lists, target),
            "milestone": str(milestone["title"])}


# --- Command line ----------------------------------------------------------------


def summary(outcome: Outcome, collected_at: datetime) -> str:
    lines = ["# Daily traffic export", "",
             f"Collection: {collected_at.isoformat(timespec='seconds')}", "",
             "| File | Rows |", "| --- | --- |"]
    lines += [f"| `data/{name}` | {len(rows)} |" for name, rows in outcome.rows.items()]
    if not outcome.day_replaced:
        lines += ["", "A later collection of the same day is stored: its snapshot, "
                  "referrers and popular paths were kept."]
    if not outcome.days_replaced:
        lines += ["", "A later collection is stored: only missing days were added."]
    return "\n".join(lines) + "\n"


def parse_time(value: str) -> datetime:
    try:
        parsed = datetime.fromisoformat(value.replace("Z", "+00:00"))
    except ValueError as error:
        raise argparse.ArgumentTypeError(f"invalid timestamp {value!r}") from error
    if parsed.tzinfo is None:
        raise argparse.ArgumentTypeError("the timestamp needs a UTC offset")
    return parsed.astimezone(UTC)


def parse_args(argv: Sequence[str] | None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=(__doc__ or "").splitlines()[0])
    commands = parser.add_subparsers(dest="command", required=True)
    merge_command = commands.add_parser("merge", help="merge one collection into the history")
    merge_command.add_argument("--raw", type=Path, required=True,
                               help=f"directory holding {', '.join(RESPONSES)}")
    merge_command.add_argument("--data", type=Path, required=True,
                               help="history directory (data/ on traffic-history)")
    merge_command.add_argument("--collected-at", type=parse_time,
                               default=datetime.now(UTC).replace(microsecond=0),
                               help="UTC time of the collection (default: now)")
    merge_command.add_argument("--summary", type=Path, help="append a Markdown report here")
    alert_command = commands.add_parser(
        "alert", help="write the alert issue for the latest snapshot, if one is due")
    alert_command.add_argument("--data", type=Path, required=True)
    alert_command.add_argument("--milestones", type=Path, required=True,
                               help="JSON list of the repository's open milestones")
    alert_command.add_argument("--today", type=date.fromisoformat,
                               default=datetime.now(UTC).date(),
                               help="UTC date of the run (default: today)")
    alert_command.add_argument("--output", type=Path, required=True,
                               help="JSON file for the issue's title, body and milestone; "
                                    "not written when no alert is due")
    return parser.parse_args(argv)


def run_alert(args: argparse.Namespace) -> int:
    today = args.today.isoformat()
    milestones = load_json(args.milestones)
    if not isinstance(milestones, list):
        raise TrafficError("Unexpected milestones API response")
    # Chosen before any early exit, so a missing milestone fails every run
    # instead of only the first one that raises an alert.
    milestone = active_milestone(milestones, today)
    issue = alert(args.data, milestone, today)
    if issue is None:
        print("No alert: nothing stands out in the latest snapshot.")
        return 0
    args.output.write_text(json.dumps(issue), encoding="utf-8")
    print(issue["title"])
    return 0


def main(argv: Sequence[str] | None = None) -> int:
    args = parse_args(argv)
    try:
        if args.command == "alert":
            return run_alert(args)
        outcome = merge(args.raw, args.data, args.collected_at)
    except TrafficError as error:
        print(f"::error::{error}")
        return 2
    report = summary(outcome, args.collected_at)
    print(report, end="")
    if args.summary:
        with args.summary.open("a", encoding="utf-8") as handle:
            handle.write(report)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
