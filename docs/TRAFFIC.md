# Repository traffic analytics

`Daily traffic export` collects GitHub views, clones, referrers and popular
paths at 06:00 UTC every day. A maintainer can also run it from the Actions tab
using `main`. It runs separately from CI and does not validate VBA or models.

## Credentials and access

- Environment: `analytics`, with deployment access restricted to `main`.
- Environment secret: `TRAFFIC_TOKEN`.
- Use a fine-grained personal access token limited to this repository with
  **Administration: read** (and GitHub's mandatory metadata read permission).
- Rotate the token every 90 days and immediately after suspected exposure.
  Update the environment secret with the replacement; never commit the value.
- The token reads traffic only. GitHub's built-in `GITHUB_TOKEN` writes the
  history branch and creates alert issues.
- Keep the environment secret out of build, test and release jobs. Any workflow
  on an allowed branch can request the environment; its name is not a boundary
  that binds the token to this one workflow.

## Outputs

The first successful run creates the orphan `traffic-history` branch. Runs
maintain `data/traffic.csv`, `data/traffic_daily.csv`,
`data/referrers.csv`, `data/popular_paths.csv` and badge JSON files.
Repeated runs on the same day replace that day's snapshot rather than duplicate
it. A failed remote lookup or fetch stops the workflow instead of replacing
existing history.

GitHub's traffic API returns a rolling 14-day window; retained snapshots keep
a longer record. Unique visitors over several days cannot be obtained by adding
daily unique counts. The visitors badge is the 14-day unique count; the visits
badge sums daily unique counts and can count a returning visitor on several days.

The repository is public, so the `traffic-history` branch is readable by
anyone. It holds aggregate counts only (views, clones, referrers, popular
paths), the same data GitHub shows repository admins, and no visitor
identities. Its badge JSON files can therefore back shields.io endpoint badges.
Traffic does not require GitHub Pages.

After at least two snapshots, qualifying traffic spikes, increases in stars or
forks, or new referrers create a `P3` issue assigned to `danielep71` in the
milestone selected by repository variable `ISSUE_MILESTONE_NUMBER` (default: `1`,
v0.1.0 foundation). The selected milestone must exist and be open; other open
milestones do not affect selection. Update this variable before retiring that
milestone. The issue-metadata workflow uses the same selection.
The workflow prevents duplicate alerts for the same day.

## Verification

After integration into `main`, use **Actions → Daily traffic export → Run
workflow → main**. Confirm that the run succeeds, that `traffic-history`
contains the expected files, and that they hold aggregate counts only.
A run from a different branch is skipped. A missing/expired token or API error
fails the run; it must not be reported as a successful traffic snapshot.

## Repository backlog metric

`data/traffic.csv` uses `open_issues_and_prs` for the REST repository count,
which includes both open issues and pull requests. It is not an issue-only
backlog. The next export renames the legacy `open_issues` column while
preserving its historical values; no historical issue-only count is inferred.
Consumers of the CSV must use the new column name.
