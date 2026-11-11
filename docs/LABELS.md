# Issue labels

The canonical label catalog is `.github/labels.json`; `.github/label-policy.json` selects the `application` profile.

- Priority: exactly one of `P1`, `P2`, `P3`. `P1` is reserved for a
  confirmed calculation bug, wrong results for valid inputs, and always comes
  with `bug`; other correctness defects are `P2`. See
  [Issue triage](GOVERNANCE.md#issue-triage).
- `blocked` is the only dependency-status label; it is present exactly while
  a `Blocked by` line in an issue's dependencies section names an open issue,
  or the section names an external condition.
- Type/area: `repository`, `enhancement`, `tests`, `documentation`, `ci`, `security`, etc.
- Labels are synchronized by `.github/workflows/labels-sync.yml` on trusted pushes to main or manual dispatch.
- With `"prune": true`, a live label outside the catalog is deleted, which
  removes it from every issue and pull request and cannot be undone. Before
  any change, the reconcile run checks each label it would delete against all
  issues and pull requests, open and closed, and refuses, changing nothing,
  if any is still in use; the message lists every such label. Remove the
  label from those items or add it to `.github/labels.json`, then run again.
- To preview a reconcile, set `GITHUB_TOKEN` and run
  `node .github/scripts/labels-sync.mjs --policy .github/label-policy.json --mode reconcile --repository <owner>/<name> --dry-run`.
  The dry run reads the live labels and the use of each label it would delete,
  lists every planned change in its summary, marks each deletion blocked by a
  label in use with the issues or pull requests that use it (up to ten
  numbers, then "and others"), changes nothing and exits 0.
- Scheduled drift checks report differences; they must not silently change the repository.

Do not equate version labels with GitHub milestone assignment.
