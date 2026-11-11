"""The traffic history merge, badges and alert (tools/traffic_history.py), and its workflow."""
from __future__ import annotations

import contextlib
import csv
import io
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import textwrap
import unittest
from collections.abc import Iterator, Mapping, Sequence
from datetime import UTC, date, datetime, timedelta
from pathlib import Path
from typing import Any

import traffic_history
from test_workflow_pins import jobs
from traffic_history import DAILY, PATHS, REFERRERS, TRAFFIC, TrafficError

ROOT = Path(__file__).resolve().parents[1]
WORKFLOW = ROOT / ".github/workflows/daily-traffic.yml"
MILESTONE = {"title": "next", "state": "open", "due_on": "2026-11-20T08:00:00Z"}
Day = tuple[str, int, int]


def at(stamp: str) -> datetime:
    return datetime.fromisoformat(stamp).replace(tzinfo=UTC)


def window(last: str, values: Sequence[tuple[int, int]]) -> list[Day]:
    """Days ending on ``last``, oldest first, one per value pair."""
    end = date.fromisoformat(last)
    return [((end - timedelta(days=len(values) - 1 - offset)).isoformat(), total, unique)
            for offset, (total, unique) in enumerate(values)]


def series(key: str, days: Sequence[Day], total: int = 0, uniques: int = 0) -> dict[str, Any]:
    return {"count": total, "uniques": uniques,
            key: [{"timestamp": f"{day}T00:00:00Z", "count": count, "uniques": unique}
                  for day, count, unique in days]}


class Workspace:
    """A raw response directory and a history directory under one temporary root."""

    def __init__(self, root: Path) -> None:
        self.raw = root / "raw"
        self.data = root / "data"
        self.raw.mkdir()

    def collect(self, view_days: Sequence[Day] = (), clone_days: Sequence[Day] = (), *,
                totals: tuple[int, int, int, int] = (10, 2, 4, 1), stars: int = 1,
                referrers: Sequence[tuple[str, int, int]] = (("github.com", 8, 2),),
                paths: Sequence[tuple[str, str, int, int]] = (("/r", "Overview", 9, 2),),
                replace: Mapping[str, Any] | None = None) -> None:
        responses: dict[str, Any] = {
            "views": series("views", view_days, totals[0], totals[1]),
            "clones": series("clones", clone_days, totals[2], totals[3]),
            "referrers": [{"referrer": name, "count": count, "uniques": unique}
                          for name, count, unique in referrers],
            "paths": [{"path": path, "title": title, "count": count, "uniques": unique}
                      for path, title, count, unique in paths],
            "repository": {"stargazers_count": stars, "forks_count": 0,
                           "subscribers_count": 1, "open_issues_count": 4},
        }
        responses.update(replace or {})
        for name, value in responses.items():
            (self.raw / f"{name}.json").write_text(json.dumps(value), encoding="utf-8")

    def merge(self, stamp: str) -> traffic_history.Outcome:
        return traffic_history.merge(self.raw, self.data, at(stamp))

    def rows(self, name: str) -> list[dict[str, str]]:
        with (self.data / name).open(newline="", encoding="utf-8") as handle:
            return list(csv.DictReader(handle))

    def files(self) -> dict[str, bytes]:
        return {str(path.relative_to(self.data)): path.read_bytes()
                for path in sorted(self.data.rglob("*")) if path.is_file()}

    def write(self, name: str, header: str, *lines: str) -> None:
        self.data.mkdir(exist_ok=True)
        (self.data / name).write_text("\r\n".join([header, *lines]) + "\r\n",
                                      encoding="utf-8", newline="")


@contextlib.contextmanager
def workspace() -> Iterator[Workspace]:
    with tempfile.TemporaryDirectory() as directory:
        yield Workspace(Path(directory))


class MergeTests(unittest.TestCase):
    def test_the_first_collection_writes_every_file_in_the_branch_layout(self) -> None:
        with workspace() as space:
            space.collect(window("2026-11-02", [(3, 1), (5, 2)]), window("2026-11-02", [(7, 3)]))
            space.merge("2026-11-02T06:17:03")
            files = space.files()
            self.assertEqual(sorted(files), ["badges/clones.json", "badges/visitors.json",
                                             "badges/visits.json", "popular_paths.csv",
                                             "referrers.csv", "traffic.csv", "traffic_daily.csv"])
            headers = {name: files[name].split(b"\r\n", 1)[0].decode() for name in
                       (TRAFFIC, DAILY, REFERRERS, PATHS)}
            self.assertEqual(headers, {
                TRAFFIC: "snapshot_date,snapshot_datetime_utc,views_14d,unique_views_14d,"
                         "clones_14d,unique_clones_14d,stars,forks,watchers,open_issues_and_prs",
                DAILY: "date,views,unique_views,clones,unique_clones",
                REFERRERS: "snapshot_date,rank,referrer,count,uniques",
                PATHS: "snapshot_date,rank,path,title,count,uniques",
            })
            self.assertEqual(files[TRAFFIC].split(b"\r\n")[1],
                             b"2026-11-02,2026-11-02T06:17:03+00:00,10,2,4,1,1,0,1,4")
            self.assertEqual(files[DAILY], b"date,views,unique_views,clones,unique_clones\r\n"
                                           b"2026-11-01,3,1,0,0\r\n2026-11-02,5,2,7,3\r\n")
            self.assertEqual(space.rows(REFERRERS), [
                {"snapshot_date": "2026-11-02", "rank": "1", "referrer": "github.com",
                 "count": "8", "uniques": "2"}])
            self.assertEqual(space.rows(PATHS)[0]["title"], "Overview")

    def test_reported_days_replace_stored_ones_and_older_days_are_kept(self) -> None:
        with workspace() as space:
            space.collect(window("2026-11-01", [(1, 1), (2, 1)]))
            space.merge("2026-11-01T06:00:00")
            space.collect(window("2026-11-02", [(9, 4), (6, 3)]))
            space.merge("2026-11-02T06:00:00")
            daily = {row["date"]: row["views"] for row in space.rows(DAILY)}
            self.assertEqual(daily, {"2026-10-31": "1", "2026-11-01": "9", "2026-11-02": "6"})
            self.assertEqual([row["snapshot_date"] for row in space.rows(TRAFFIC)],
                             ["2026-11-01", "2026-11-02"])
            self.assertEqual([row["snapshot_date"] for row in space.rows(REFERRERS)],
                             ["2026-11-01", "2026-11-02"])

    def test_ranked_rows_keep_the_api_order(self) -> None:
        with workspace() as space:
            space.collect(referrers=[("b.example", 1, 1), ("a.example", 9, 3)],
                          paths=[("/r", "Overview, main", 4, 1), ("/r/issues", "Issues", 2, 1)])
            space.merge("2026-11-02T06:00:00")
            self.assertEqual([(row["rank"], row["referrer"]) for row in space.rows(REFERRERS)],
                             [("1", "b.example"), ("2", "a.example")])
            self.assertEqual(space.rows(PATHS)[0]["title"], "Overview, main")

    def test_earlier_traffic_layouts_are_converted(self) -> None:
        with workspace() as space:
            space.write(TRAFFIC, "date,views,unique_views,clones,unique_clones,open_issues",
                        "2026-10-01T06:00:00+00:00,4,1,2,1,3")
            space.collect()
            space.merge("2026-11-02T06:00:00")
            rows = space.rows(TRAFFIC)
            self.assertEqual((rows[0]["snapshot_date"], rows[0]["snapshot_datetime_utc"],
                              rows[0]["views_14d"], rows[0]["open_issues_and_prs"]),
                             ("2026-10-01", "2026-10-01T06:00:00+00:00", "4", "3"))
            self.assertEqual(rows[0]["stars"], "0")

    def test_an_invalid_response_leaves_the_history_unchanged(self) -> None:
        with workspace() as space:
            space.collect(window("2026-11-01", [(1, 1)]))
            space.merge("2026-11-01T06:00:00")
            before = space.files()
            overrides: list[dict[str, Any]] = [
                {"views": {"views": []}}, {"clones": {"count": 1, "clones": 3}},
                {"referrers": {"referrer": "x"}}, {"paths": [1]}, {"repository": {"name": "r"}}]
            for override in overrides:
                with self.subTest(override=override):
                    space.collect(window("2026-11-02", [(5, 1)]), replace=override)
                    with self.assertRaisesRegex(TrafficError, "Unexpected"):
                        space.merge("2026-11-02T06:00:00")
                    self.assertEqual(space.files(), before)
            space.collect(window("2026-11-02", [(5, 1)]))
            (space.raw / "paths.json").unlink()
            with self.assertRaisesRegex(TrafficError, "missing"):
                space.merge("2026-11-02T06:00:00")
            self.assertEqual(space.files(), before)

    def test_the_command_reports_an_invalid_response_and_writes_nothing(self) -> None:
        with workspace() as space:
            space.collect(replace={"views": {"count": 1}})
            out = io.StringIO()
            with contextlib.redirect_stdout(out):
                code = traffic_history.main(["merge", "--raw", str(space.raw),
                                             "--data", str(space.data)])
            self.assertEqual(code, 2)
            self.assertIn("::error::Unexpected views API response", out.getvalue())
            self.assertFalse(space.data.exists())


class ReplayTests(unittest.TestCase):
    def test_a_later_collection_on_the_same_day_replaces_that_day(self) -> None:
        with workspace() as space:
            space.collect(window("2026-11-02", [(1, 1)]), referrers=[("a", 5, 1), ("b", 1, 1)])
            space.merge("2026-11-02T06:00:00")
            space.collect(window("2026-11-02", [(4, 2)]), totals=(20, 5, 4, 1),
                          referrers=[("c", 7, 2)], paths=[])
            space.merge("2026-11-02T18:00:00")
            traffic = space.rows(TRAFFIC)
            self.assertEqual(len(traffic), 1)
            self.assertEqual((traffic[0]["snapshot_datetime_utc"], traffic[0]["views_14d"]),
                             ("2026-11-02T18:00:00+00:00", "20"))
            self.assertEqual([row["referrer"] for row in space.rows(REFERRERS)], ["c"])
            self.assertEqual(space.rows(PATHS), [])
            self.assertEqual([(row["date"], row["views"]) for row in space.rows(DAILY)],
                             [("2026-11-02", "4")])

    def test_replaying_the_same_collection_changes_nothing(self) -> None:
        with workspace() as space:
            space.collect(window("2026-11-02", [(1, 1), (3, 2)]), window("2026-11-02", [(2, 1)]))
            space.merge("2026-11-02T06:00:00")
            first = space.files()
            space.merge("2026-11-02T06:00:00")
            self.assertEqual(space.files(), first)

    def test_an_older_collection_replayed_after_a_newer_one_is_ignored(self) -> None:
        with workspace() as space:
            space.collect(window("2026-11-02", [(4, 2)]), totals=(20, 5, 4, 1),
                          referrers=[("new", 7, 2)])
            space.merge("2026-11-02T18:00:00")
            newer = space.files()
            space.collect(window("2026-11-02", [(1, 1)]), totals=(10, 2, 4, 1),
                          referrers=[("old", 5, 1)])
            outcome = space.merge("2026-11-02T06:00:00")
            self.assertFalse(outcome.day_replaced)
            self.assertFalse(outcome.days_replaced)
            self.assertEqual(space.files(), newer)

    def test_an_older_collection_only_fills_what_is_missing(self) -> None:
        with workspace() as space:
            space.collect(window("2026-11-02", [(9, 4), (8, 3)]))
            space.merge("2026-11-02T06:00:00")
            # Yesterday's collection, replayed afterwards: its snapshot date and
            # its first two days are not stored yet; the day it shares is newer.
            space.collect(window("2026-11-01", [(5, 1), (1, 1), (1, 1)]), totals=(7, 1, 0, 0),
                          referrers=[("y", 6, 1)])
            outcome = space.merge("2026-11-01T06:00:00")
            self.assertTrue(outcome.day_replaced)
            self.assertFalse(outcome.days_replaced)
            self.assertEqual([(row["date"], row["views"]) for row in space.rows(DAILY)],
                             [("2026-10-30", "5"), ("2026-10-31", "1"), ("2026-11-01", "9"),
                              ("2026-11-02", "8")])
            self.assertEqual([(row["snapshot_date"], row["views_14d"])
                              for row in space.rows(TRAFFIC)],
                             [("2026-11-01", "7"), ("2026-11-02", "10")])
            self.assertEqual([(row["snapshot_date"], row["referrer"])
                              for row in space.rows(REFERRERS)],
                             [("2026-11-01", "y"), ("2026-11-02", "github.com")])
            badge = json.loads((space.data / "badges/visitors.json").read_text())
            self.assertEqual(badge["message"], "2")


class BadgeTests(unittest.TestCase):
    def test_compact_counts(self) -> None:
        cases = {0: "0", 999: "999", 1000: "1k", 1234: "1.2k", 10_000: "10k",
                 999_999: "1000k", 1_000_000: "1M", 3_400_000: "3.4M"}
        for value, expected in cases.items():
            with self.subTest(value=value):
                self.assertEqual(traffic_history.compact_count(value), expected)

    def test_badges_come_from_the_stored_history(self) -> None:
        with workspace() as space:
            space.write(DAILY, "date,views,unique_views,clones,unique_clones",
                        "2026-10-01,10,600,1500,3", "2026-10-02,2,450,1000,1")
            space.collect(window("2026-11-02", [(4, 3)]), window("2026-11-02", [(40, 2)]),
                          totals=(4, 3, 40, 2))
            space.merge("2026-11-02T06:00:00")
            files = space.files()
            self.assertEqual(files["badges/clones.json"],
                             b'{"schemaVersion": 1, "label": "clones", "message": "2.5k", '
                             b'"color": "217346"}\n')
            self.assertEqual(files["badges/visits.json"],
                             b'{"schemaVersion": 1, "label": "visits", "message": "1.1k", '
                             b'"color": "217346"}\n')
            self.assertEqual(files["badges/visitors.json"],
                             b'{"schemaVersion": 1, "label": "visitors (14d)", "message": "3", '
                             b'"color": "217346"}\n')


TRAFFIC_HEADER = ",".join(traffic_history.HEADERS[TRAFFIC])


def snapshot(day: str, views: int = 10, clones: int = 4, stars: int = 1, forks: int = 0) -> str:
    return f"{day},{day}T06:00:00+00:00,{views},2,{clones},1,{stars},{forks},1,4"


class AlertTests(unittest.TestCase):
    def history(self, space: Workspace, previous: str, current: str,
                referrers: Sequence[str] = ()) -> None:
        space.write(TRAFFIC, TRAFFIC_HEADER, previous, current)
        space.write(REFERRERS, ",".join(traffic_history.HEADERS[REFERRERS]), *referrers)
        space.write(PATHS, ",".join(traffic_history.HEADERS[PATHS]),
                    "2026-11-02,1,/r,Overview,9,2")

    def alert(self, previous: str, current: str, referrers: Sequence[str] = (),
              today: str = "2026-11-02",
              milestone: Mapping[str, Any] = MILESTONE) -> dict[str, str] | None:
        with workspace() as space:
            self.history(space, previous, current, referrers)
            return traffic_history.alert(space.data, milestone, today)

    def test_quiet_traffic_raises_no_alert(self) -> None:
        self.assertIsNone(self.alert(snapshot("2026-11-01"), snapshot("2026-11-02", views=19)))

    def test_one_snapshot_raises_no_alert(self) -> None:
        with workspace() as space:
            space.write(TRAFFIC, TRAFFIC_HEADER, snapshot("2026-11-02", stars=9))
            self.assertIsNone(traffic_history.alert(space.data, MILESTONE, "2026-11-02"))

    def test_a_views_spike_doubles_and_grows_by_100(self) -> None:
        self.assertIsNone(self.alert(snapshot("2026-11-01", views=10),
                                     snapshot("2026-11-02", views=109)))
        issue = self.alert(snapshot("2026-11-01", views=10), snapshot("2026-11-02", views=110))
        assert issue is not None
        self.assertEqual(issue["title"], "Traffic alert: 2026-11-02")
        self.assertIn("Trigger reason:\nViews spike\n", issue["body"])

    def test_clones_are_reported_but_never_a_reason(self) -> None:
        self.assertIsNone(self.alert(snapshot("2026-11-01", clones=4),
                                     snapshot("2026-11-02", clones=4000)))
        issue = self.alert(snapshot("2026-11-01", clones=4),
                           snapshot("2026-11-02", clones=4000, stars=2))
        assert issue is not None
        self.assertIn("Trigger reason:\nStars increased\n", issue["body"])
        self.assertNotIn("Clones spike", issue["body"])
        self.assertIn("- Clones: 4000 vs 4 on 2026-11-01", issue["body"])

    def test_stars_and_forks_increases(self) -> None:
        issue = self.alert(snapshot("2026-11-01"), snapshot("2026-11-02", stars=2, forks=1))
        assert issue is not None
        self.assertIn("Trigger reason:\nStars increased, Forks increased\n", issue["body"])
        self.assertIsNone(self.alert(snapshot("2026-11-01", stars=3), snapshot("2026-11-02")))

    def test_a_new_referrer_needs_five_views_and_no_earlier_appearance(self) -> None:
        quiet = (snapshot("2026-11-01"), snapshot("2026-11-02"))
        self.assertIsNone(self.alert(*quiet, ["2026-11-02,1,blog.example,4,1"]))
        self.assertIsNone(self.alert(*quiet, ["2026-10-20,1,blog.example,1,1",
                                              "2026-11-02,1,blog.example,50,9"]))
        issue = self.alert(*quiet, ["2026-11-02,1,github.com,3,1", "2026-11-02,2,blog.example,5,2"])
        assert issue is not None
        self.assertIn("Trigger reason:\nNew referrer detected\n", issue["body"])
        self.assertIn("New referrers:\n- blog.example: 5 views, 2 unique\n", issue["body"])
        self.assertIn("Current top referrers:\n- blog.example: 5 views, 2 unique\n"
                      "- github.com: 3 views, 1 unique\n", issue["body"])
        self.assertIn("Current popular paths:\n- /r: 9 views, 2 unique\n", issue["body"])

    def test_the_body_meets_the_triage_rules(self) -> None:
        spike = (snapshot("2026-11-01", stars=1), snapshot("2026-11-02", stars=2))
        cases = {("2026-11-02", "2026-11-20T08:00:00Z"): "2026-11-09",
                 ("2026-11-02", "2026-11-05T08:00:00Z"): "2026-11-04"}
        for (today, due), target in cases.items():
            with self.subTest(due=due):
                issue = self.alert(*spike, today=today,
                                   milestone={"title": "m", "state": "open", "due_on": due})
                assert issue is not None
                self.assertEqual(issue["milestone"], "m")
                self.assertTrue(
                    issue["body"].startswith("Automated traffic alert for **K-IRRBB**.\n"))
                self.assertIn("\n## Dependencies\n\nNone.\n", issue["body"])
                self.assertTrue(issue["body"].endswith(
                    f"\n- Start date: {today} (Europe/Rome)"
                    f"\n- Target end date: {target} (Europe/Rome)"), issue["body"][-80:])


class AlertCommandTests(unittest.TestCase):
    def run_alert(self, space: Workspace, milestones: Any) -> tuple[int, str, Path]:
        listing = space.raw / "milestones.json"
        listing.write_text(json.dumps(milestones), encoding="utf-8")
        output = space.raw / "alert.json"
        out = io.StringIO()
        with contextlib.redirect_stdout(out):
            code = traffic_history.main(["alert", "--data", str(space.data), "--milestones",
                                         str(listing), "--today", "2026-11-02",
                                         "--output", str(output)])
        return code, out.getvalue(), output

    def test_no_milestone_fails_even_without_an_alert(self) -> None:
        with workspace() as space:
            code, out, output = self.run_alert(space, [])
            self.assertEqual(code, 2)
            self.assertIn("::error::No open milestone is due after today", out)
            self.assertFalse(output.exists())

    def test_no_alert_writes_no_issue(self) -> None:
        with workspace() as space:
            code, out, output = self.run_alert(space, [MILESTONE])
            self.assertEqual(code, 0)
            self.assertIn("No alert", out)
            self.assertFalse(output.exists())

    def test_an_alert_writes_the_issue_with_its_milestone(self) -> None:
        with workspace() as space:
            space.write(TRAFFIC, TRAFFIC_HEADER, snapshot("2026-11-01"),
                        snapshot("2026-11-02", forks=1))
            code, _, output = self.run_alert(space, [MILESTONE, {**MILESTONE, "title": "later",
                                                                 "due_on": "2027-01-01T00:00:00Z"}])
            self.assertEqual(code, 0)
            issue = json.loads(output.read_text(encoding="utf-8"))
            self.assertEqual((issue["title"], issue["milestone"]),
                             ("Traffic alert: 2026-11-02", "next"))


class WorkflowTests(unittest.TestCase):
    text: str
    jobs: dict[str, str]

    @classmethod
    def setUpClass(cls) -> None:
        cls.text = WORKFLOW.read_text(encoding="utf-8")
        cls.jobs = jobs(cls.text)

    def steps(self, job: str) -> dict[str, str]:
        parts = re.split(r"^      - name: (.+)$", self.jobs[job], flags=re.M)
        return dict(zip(parts[1::2], parts[2::2], strict=True))

    def test_the_jobs_hold_least_privilege(self) -> None:
        self.assertIn("\npermissions: {}\n", self.text)
        self.assertEqual(sorted(self.jobs), ["alert", "collect"])
        granted = {job: re.findall(r"^      ([\w-]+): (read|write)$", text, re.M)
                   for job, text in self.jobs.items()}
        self.assertEqual(granted, {"collect": [("contents", "write")],
                                   "alert": [("contents", "read"), ("issues", "write")]})
        self.assertIn("    needs: collect\n", self.jobs["alert"])
        self.assertIn("    environment: analytics\n", self.jobs["collect"])
        self.assertNotRegex(self.jobs["alert"], r"(?m)^\s+environment:")
        self.assertNotIn("TRAFFIC_APP", self.jobs["alert"])
        self.assertIn("ref: ${{ needs.collect.outputs.history-sha }}", self.jobs["alert"])

    def test_only_the_fetch_step_receives_the_app_token(self) -> None:
        self.assertNotIn("TRAFFIC_TOKEN", self.text)
        self.assertIn("uses: actions/create-github-app-token@"
                      "64ab729a6294000f5aa4a4bc0b73d1f32bc45510 # v3.3.0", self.text)
        steps = self.steps("collect")
        token = steps["Create the traffic App token"]
        self.assertIn("client-id: ${{ vars.TRAFFIC_APP_CLIENT_ID }}", token)
        self.assertIn("private-key: ${{ secrets.TRAFFIC_APP_PRIVATE_KEY }}", token)
        self.assertEqual(re.findall(r"permission-[\w-]+: \w+", token),
                         ["permission-administration: read"])
        users = [name for name, text in steps.items() if "steps.app-token.outputs" in text]
        self.assertEqual(users, ["Fetch GitHub traffic API data"])
        keys = [name for name, text in steps.items() if "TRAFFIC_APP_PRIVATE_KEY" in text]
        self.assertEqual(keys, ["Check the traffic App configuration",
                                "Create the traffic App token"])

    def test_no_checkout_keeps_credentials_and_only_the_push_authenticates(self) -> None:
        self.assertEqual(len(re.findall(r"^\s+uses: actions/checkout@", self.text, re.M)),
                         len(re.findall(r"^\s+persist-credentials: false$", self.text, re.M)))
        self.assertNotIn("persist-credentials: true", self.text)
        users = {job: [name for name, text in self.steps(job).items() if "github.token" in text]
                 for job in self.jobs}
        self.assertEqual(users, {"collect": ["Commit and push traffic history"],
                                 "alert": ["Create alert issue on spike or new referrer"]})

    def test_it_runs_daily_off_the_hour_and_only_from_the_default_branch(self) -> None:
        crons = re.findall(r'^    - cron: "(\d+) (\d+) \* \* \*"$', self.text, re.M)
        self.assertEqual(len(crons), 1)
        self.assertEqual(len(re.findall(r"cron:", self.text)), 1)
        minute, hour = (int(field) for field in crons[0])
        self.assertNotIn(minute, (0, 30))
        documented = re.search(r"at (\d{2}):(\d{2}) UTC every day",
                               " ".join((ROOT / "docs/TRAFFIC.md").read_text(encoding="utf-8").split()))
        assert documented is not None, "docs/TRAFFIC.md states no daily UTC time"
        self.assertEqual((hour, minute), (int(documented.group(1)), int(documented.group(2))))
        self.assertIn("\n  workflow_dispatch:\n", self.text)
        self.assertIn("github.ref == format('refs/heads/{0}', "
                      "github.event.repository.default_branch)", self.jobs["collect"])

    def test_the_alert_keeps_its_assignee_label_and_duplicate_check(self) -> None:
        step = self.steps("alert")["Create alert issue on spike or new referrer"]
        self.assertIn("--assignee danielep71 --label P3", step)
        self.assertIn('--milestone "$milestone"', step)
        self.assertIn('gh issue list --state all --search "\\"$title\\" in:title"', step)
        self.assertIn("tools/traffic_history.py alert", step)

    def test_the_push_retries_rebuilds_and_always_reports_its_commit(self) -> None:
        step = self.steps("collect")["Commit and push traffic history"]
        self.assertIn("COLLECTED_AT: ${{ steps.fetch.outputs.collected_at }}", step)
        self.assertIn("for attempt in 1 2 3; do", step)
        rebuild = step.split("fetch origin traffic-history", 1)[1]
        self.assertIn("git reset -q --hard FETCH_HEAD", rebuild)
        self.assertIn('tools/traffic_history.py" merge', rebuild)
        self.assertIn("commit_history", rebuild)
        self.assertEqual(step.count('echo "sha=$(git rev-parse HEAD)" >> "$GITHUB_OUTPUT"'), 1)
        self.assertLess(step.index("sha=$(git rev-parse HEAD)"), step.index("exit 0"))

    def test_the_alert_checks_the_history_commit_before_checking_it_out(self) -> None:
        names = list(self.steps("alert"))
        self.assertLess(names.index("Check the pushed history commit"),
                        names.index("Checkout the pushed traffic history"))
        check = self.steps("alert")["Check the pushed history commit"]
        self.assertIn("HISTORY_SHA: ${{ needs.collect.outputs.history-sha }}", check)
        self.assertIn('[[ "$HISTORY_SHA" =~ ^[0-9a-f]{40}$ ]]', check)
        script = check.split("run: |\n", 1)[1]
        for value, code in (("", 1), ("main", 1), ("A" * 40, 1), ("a" * 39, 1), ("0" * 40, 0)):
            with self.subTest(value=value):
                result = subprocess.run(["bash", "-c", textwrap.dedent(script)], capture_output=True,
                                        text=True, env={**os.environ, "HISTORY_SHA": value})
                self.assertEqual(result.returncode, code, result.stdout)


GIT = shutil.which("git")


def push_step() -> str:
    collect = jobs(WORKFLOW.read_text(encoding="utf-8"))["collect"]
    parts = re.split(r"^      - name: (.+)$", collect, flags=re.M)
    return dict(zip(parts[1::2], parts[2::2], strict=True))["Commit and push traffic history"]


@unittest.skipUnless(GIT and shutil.which("bash"), "needs git and bash")
class PushStepTests(unittest.TestCase):
    """The push step itself, run with bash against local repositories."""

    def setUp(self) -> None:
        self.root = Path(tempfile.mkdtemp())
        self.addCleanup(shutil.rmtree, self.root)
        self.origin = self.root / "origin.git"
        self.git("init", "-q", "--bare", "-b", "main", str(self.origin))
        seed = self.clone("seed", branch=None)
        self.git("checkout", "-q", "--orphan", "traffic-history", cwd=seed)
        self.collect(seed, "2026-10-31")
        self.git("add", "data", cwd=seed)
        self.git("commit", "-q", "-m", "seed", cwd=seed)
        self.git("push", "-q", "origin", "traffic-history", cwd=seed)
        self.history = self.clone("history")
        self.raw = self.collect(self.history, "2026-11-02")
        shims = self.root / "bin"
        shims.mkdir()
        (shims / "python").symlink_to(sys.executable)
        self.output = self.root / "output"
        self.output.touch()
        self.env = {**os.environ, "PATH": f"{shims}{os.pathsep}{os.environ['PATH']}",
                    "HISTORY": str(self.history), "RAW": str(self.raw),
                    "COLLECTED_AT": "2026-11-02T06:17:00Z", "GITHUB_OUTPUT": str(self.output),
                    "GITHUB_WORKSPACE": str(ROOT), "GITHUB_SERVER_URL": "https://github.invalid",
                    "GITHUB_TOKEN": "not-a-token", "GIT_CONFIG_GLOBAL": str(self.root / "gitconfig")}
        (self.root / "gitconfig").write_text("[user]\n\tname = test\n\temail = test@invalid\n"
                                             "[init]\n\tdefaultBranch = main\n", encoding="utf-8")

    def git(self, *args: str, cwd: Path | None = None) -> str:
        assert GIT is not None
        env = {**os.environ, "GIT_AUTHOR_NAME": "test", "GIT_AUTHOR_EMAIL": "test@invalid",
               "GIT_COMMITTER_NAME": "test", "GIT_COMMITTER_EMAIL": "test@invalid"}
        return subprocess.run([GIT, *args], cwd=cwd, check=True, capture_output=True,
                              text=True, env=env).stdout.strip()

    def clone(self, name: str, branch: str | None = "traffic-history") -> Path:
        target = self.root / name
        extra = ["-b", branch] if branch else []
        self.git("clone", "-q", *extra, str(self.origin), str(target))
        return target

    def collect(self, checkout: Path, day: str) -> Path:
        """Merge one collection for ``day`` into ``checkout``/data; return its raw folder."""
        folder = self.root / f"collection-{checkout.name}-{day}"
        folder.mkdir()
        workspace = Workspace(folder)
        workspace.collect(window(day, [(3, 1)]), window(day, [(2, 1)]))
        traffic_history.merge(workspace.raw, checkout / "data", at(f"{day}T06:17:00"))
        return workspace.raw

    def run_step(self) -> subprocess.CompletedProcess[str]:
        script = textwrap.dedent(push_step().split("run: |\n", 1)[1])
        return subprocess.run(["bash", "-c", script], capture_output=True, text=True, env=self.env)

    def snapshot_dates(self) -> list[str]:
        text = self.git("show", "traffic-history:data/traffic.csv", cwd=self.origin)
        return [row["snapshot_date"] for row in csv.DictReader(io.StringIO(text))]

    def test_a_rejected_push_is_rebuilt_on_the_moved_branch(self) -> None:
        other = self.clone("other")
        self.collect(other, "2026-11-01")
        self.git("commit", "-qam", "concurrent", cwd=other)
        self.git("push", "-q", "origin", "traffic-history", cwd=other)
        result = self.run_step()
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertEqual(result.stdout.count("::warning::Push attempt"), 1)
        self.assertEqual(self.snapshot_dates(), ["2026-10-31", "2026-11-01", "2026-11-02"])
        head = self.git("rev-parse", "traffic-history", cwd=self.origin)
        self.assertEqual(self.output.read_text(encoding="utf-8"), f"sha={head}\n")

    def test_no_change_still_reports_the_commit(self) -> None:
        self.assertEqual(self.run_step().returncode, 0)
        self.output.write_text("", encoding="utf-8")
        result = self.run_step()
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertIn("No traffic changes to commit", result.stdout)
        head = self.git("rev-parse", "traffic-history", cwd=self.origin)
        self.assertEqual(self.output.read_text(encoding="utf-8"), f"sha={head}\n")

    def test_three_rejections_fail_without_a_commit(self) -> None:
        hook = self.origin / "hooks/pre-receive"
        hook.write_text("#!/bin/sh\nexit 1\n", encoding="utf-8")
        hook.chmod(0o755)
        result = self.run_step()
        self.assertEqual(result.returncode, 1)
        self.assertEqual(result.stdout.count("::warning::Push attempt"), 2)
        self.assertIn("rejected 3 times", result.stdout)
        self.assertEqual(self.output.read_text(encoding="utf-8"), "")
        self.assertEqual(self.snapshot_dates(), ["2026-10-31"])


if __name__ == "__main__":
    unittest.main()
