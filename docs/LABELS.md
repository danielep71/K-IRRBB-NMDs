# Issue labels

The canonical label catalog is `.github/labels.json`; `.github/label-policy.json` selects the `application` profile.

- Priority: exactly one of `P1`, `P2`, `P3`. `P1` is reserved for a
  confirmed calculation bug, wrong results for valid inputs, and always comes
  with `bug`; other correctness defects are `P2`. See
  [Issue triage](GOVERNANCE.md#issue-triage).
- `blocked` is the only dependency-status label; it is present exactly while
  a `Blocked by` line in an issue's dependencies section names an open issue,
  or the section names an external condition.
- Type: `bug` (a confirmed defect), `enhancement` (a new capability or an
  improvement), `refactor` (restructuring that preserves supported behavior).
  Add `behavior-change` to any of them that changes externally observable
  behavior on purpose.
- Area: `repository`, `ci`, `tests`, `documentation`, `security`, `release`.
- Contribution: `good first issue` (a well-scoped first task) and `help wanted`
  (I welcome outside work on it).
- Status and resolution: `question` (further information or my decision is
  required), `duplicate`, `invalid`, `wontfix`.
- Each label's description in `.github/labels.json` states its meaning.
- Labels are synchronized by `.github/workflows/labels-sync.yml` on trusted pushes to main or manual dispatch.
- With `"prune": true`, a live label outside the catalog is deleted, which
  removes it from every issue and pull request and cannot be undone. Before
  any change, the reconcile run checks each label it would delete against all
  issues and pull requests, open and closed, and refuses, changing nothing,
  if any is still in use; the message lists every such label. Remove the
  label from those items or add it to `.github/labels.json`, then run again.
- Scheduled drift checks report differences; they must not silently change the repository.
