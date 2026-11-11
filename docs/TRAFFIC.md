# Repository traffic analytics

`Daily traffic export` (`.github/workflows/daily-traffic.yml`) collects GitHub
views, clones, referrers and popular paths at 06:17 UTC every day; GitHub may
start a scheduled run late. I can also run it from the Actions tab using
`main`; a run from any other branch is skipped. It runs separately from CI and
does not validate VBA or models.

The workflow has two jobs:

| Job | Permissions | What it does |
| --- | --- | --- |
| `collect` | `contents: write`, `analytics` environment | Creates a GitHub App token, fetches the traffic API, merges the collection into the history with `tools/traffic_history.py` and pushes it to `traffic-history`. |
| `alert` | `contents: read`, `issues: write`, no environment | Checks out the history the collect job pushed and opens an alert issue when one is due. |

Every checkout sets `persist-credentials: false`. Only the push step
authenticates, with the workflow token, through a Git header that is never
written to disk.

If `traffic-history` moved after the checkout, the push is rejected. The push
step then fetches the branch, resets to it, merges the same collection again
and commits, up to three attempts in all; after the third rejection the run
fails and the collection is not saved. Every successful exit, including "no
change", reports the commit it pushed, and the alert job refuses to start
unless that report is a full 40-character commit SHA, so it never reads the
default branch by mistake.

## Credentials and access

The traffic API requires **Administration: read**, which the workflow token
cannot have. A GitHub App that I own, with that permission only, issues a
short-lived installation token in each run:

- Environment: `analytics`, with deployment access restricted to `main`. The
  settings baseline records this branch policy and
  `tools/setup_repository.py --check` verifies it
  ([tools/README.md](../tools/README.md#repository-settings)).
- Environment variable: `TRAFFIC_APP_CLIENT_ID`, the App's client ID.
- Environment secret: `TRAFFIC_APP_PRIVATE_KEY`, a private key of the App.
- `actions/create-github-app-token` turns them into an installation token for
  this repository with `Administration: read` only. The token expires after an
  hour and is revoked when the collect job ends, so it needs no rotation.
- Only the token step receives the private key, and only the fetch step
  receives the token. Neither runs repository code. GitHub's built-in
  `GITHUB_TOKEN` pushes the history branch and creates alert issues.
- Keep the environment out of build, test and release jobs. Any workflow on an
  allowed branch can request the environment; its name is not a boundary that
  binds the key to this one workflow.
- Replace the private key immediately after suspected exposure: generate a new
  key in the App's settings, update the environment secret, then delete the old
  key there. Never commit a key.

### Setting up the App

1. Create a GitHub App under my account with **Repository permissions →
   Administration: Read-only** and nothing else, webhooks inactive, installable
   only on my account.
2. Install it on this repository only.
3. In **Settings → Environments → analytics**, add the App's client ID as the
   variable `TRAFFIC_APP_CLIENT_ID` and a newly generated private key as the
   secret `TRAFFIC_APP_PRIVATE_KEY`.
4. Run the workflow once from `main` and confirm both jobs succeed.
5. Then delete the earlier `TRAFFIC_TOKEN` environment secret and revoke the
   fine-grained personal access token it held; the workflow no longer reads it.

A run without the variable or the secret fails at its first step that needs
them, with a message naming both.

## Outputs

The first successful run creates the orphan `traffic-history` branch. Runs
maintain these files on it:

| File | Columns |
| --- | --- |
| `data/traffic.csv` | `snapshot_date`, `snapshot_datetime_utc`, `views_14d`, `unique_views_14d`, `clones_14d`, `unique_clones_14d`, `stars`, `forks`, `watchers`, `open_issues_and_prs` |
| `data/traffic_daily.csv` | `date`, `views`, `unique_views`, `clones`, `unique_clones` |
| `data/referrers.csv` | `snapshot_date`, `rank`, `referrer`, `count`, `uniques` |
| `data/popular_paths.csv` | `snapshot_date`, `rank`, `path`, `title`, `count`, `uniques` |
| `data/badges/clones.json`, `visitors.json`, `visits.json` | shields.io endpoint badges |

`tools/traffic_history.py` merges each collection by these rules:

- A collection date's snapshot, referrers and popular paths are one unit. A
  later run on the same day replaces that day's unit rather than duplicating
  it; a collection replayed after a later one of the same day leaves it
  unchanged.
- The days a collection reports replace the stored rows of
  `data/traffic_daily.csv`, and earlier days keep theirs. If a later
  collection is already stored, an older one only adds the days that are
  missing.
- Every response and existing file is read and checked before anything is
  written. A failed fetch or an unexpected response stops the workflow instead
  of replacing existing history.

GitHub's traffic API returns a rolling 14-day window; retained snapshots keep
a longer record. Unique visitors over several days cannot be obtained by adding
daily unique counts. The clones badge totals the daily clones; the visitors
badge is the latest snapshot's 14-day unique count; the visits badge sums daily
unique counts and can count a returning visitor on several days.

The repository is public, so the `traffic-history` branch is readable by
anyone. It holds aggregate counts only (views, clones, referrers, popular
paths), the same data GitHub shows repository admins, and no visitor
identities. Its badge JSON files can therefore back shields.io endpoint badges.
Traffic does not require GitHub Pages.

## Traffic alert

After at least two snapshots, the alert job compares the latest snapshot with
the one before. A views spike (the 14-day views at least double and grow by
100), an increase in stars or forks, or a new referrer with at least five views
creates a `P3` issue assigned to `danielep71`. Clones are not a reason, because
CI checkouts of this repository count as clones; they still appear in the
CSV files, in the alert body and in the clones badge.

The issue goes into the open milestone with the earliest due date after today
(UTC). The milestone is chosen before anything else, so the alert job fails,
with a message saying so, whenever no open milestone is due after today, even
when no alert is due. The issue body carries a dependencies section and a
target end date, a week ahead or the day before the milestone's due date if
that is sooner, so it meets the [issue triage rules](GOVERNANCE.md#issue-triage).
The job opens at most one alert per snapshot date.

## Verification

After integration into `main`, use **Actions → Daily traffic export → Run
workflow → main**. Confirm that both jobs succeed, that `traffic-history`
contains the expected files, and that they hold aggregate counts only. A
missing App configuration, an API error or an unexpected response fails the
run; it must not be reported as a successful traffic snapshot. The merge, the
badges and the alert are tested offline by `tools/test_traffic_history.py`.

## Repository backlog metric

`data/traffic.csv` uses `open_issues_and_prs` for the REST repository count,
which includes both open issues and pull requests. It is not an issue-only
backlog. A legacy `open_issues` column is renamed on the next export,
preserving its historical values; no historical issue-only count is inferred.
Consumers of the CSV must use the new column name.
