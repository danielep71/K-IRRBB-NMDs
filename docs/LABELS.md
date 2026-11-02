# Issue labels

The canonical label catalog is `.github/labels.json`; `.github/label-policy.json` selects the `application` profile.

- Priority: exactly one of `P1`, `P2`, `P3`. `P1` is reserved for a
  confirmed calculation or correctness bug and always comes with `bug`; see
  [Issue triage](GOVERNANCE.md#issue-triage).
- `blocked` is the only dependency-status label; it is present exactly while
  an issue's dependencies section names an open prerequisite or an external
  condition.
- Type/area: `repository`, `enhancement`, `tests`, `documentation`, `ci`, `security`, etc.
- Labels are synchronized by `.github/workflows/labels-sync.yml` on trusted pushes to main or manual dispatch.
- Scheduled drift checks report differences; they must not silently change the repository.

Do not equate version labels with GitHub milestone assignment.
