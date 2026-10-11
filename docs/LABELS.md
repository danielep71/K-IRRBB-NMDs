# Issue labels

The canonical label catalog is `.github/labels.json`; `.github/label-policy.json` selects the `application` profile. This catalog and reconciliation scripts are adapted from VBA-SACCR-Toolkit.

- Priority: exactly one of `P1`, `P2`, `P3`.
- Type/area: `repository`, `enhancement`, `tests`, `documentation`, `ci`, `security`, etc.
- Labels are synchronized by `.github/workflows/labels-sync.yml` on trusted pushes to main or manual dispatch.
- Scheduled drift checks report differences; they must not silently change the repository.

Do not equate version labels with GitHub milestone assignment.
