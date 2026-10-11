# Governance

## Workflow

- Active integration branch: `release/v0.2.0`.
- Start focused task branches from the release branch; open PRs back into it.
- No direct commits to `release/**` or `main` after the initial repository bootstrap.
- A release-to-main PR requires the owner's explicit approval. No version tags or releases are implied.
- Merge only with reviewer sign-off and green repository checks; Excel evidence is separately recorded when VBA changes exist.

## Issue metadata

- Every issue, open or closed, must include `danielep71` as assignee.
- Exactly one priority: `P1` (blocking correctness/security), `P2` (significant), `P3` (non-blocking).
- Add suitable type labels, a milestone and verifiable checkboxes; titles do not carry version prefixes.
- Issues may be closed only when their acceptance criteria are evidenced; links to code and tests should be recorded.

The trusted default-branch `issue-metadata.yml` workflow reconciles issue events
and a daily sweep, including closed issues, and verifies the saved metadata.
It preserves an existing milestone; missing milestones use repository variable
`ISSUE_MILESTONE_NUMBER` (default `1`), which must identify an open milestone.
It adds `P3` only when no priority exists; conflicting priorities retain the
highest severity (`P1`, then `P2`, then `P3`). Other labels and assignees remain.
A manual dispatch on the default branch performs a full sweep. Change the
variable before closing the configured default milestone. Issue events that
GitHub raises for pull requests are skipped before any API call, and a failed
API call is logged with its HTTP status and response, with token-like strings
redacted (#41).

Administrator bypass is an exceptional capability, not permission to skip review
or failing checks. The intended `main` and `release/**` bypass mode is **pull
requests only**. Verify the saved GitHub ruleset after any settings change; an
unsaved edit or an authentication prompt does not establish enforcement.

## Privacy and intellectual property

Use only synthetic examples and independent work. The repository is public, and no visibility setting or private fork is authorization to store client files, personal data or confidential vendor methods. Third-party licensing and method provenance must be resolved before implementation. Contributor rules are in [`CONTRIBUTING.md`](../CONTRIBUTING.md#data-confidentiality-and-provenance).

## Licensing and distribution

The [Mozilla Public License 2.0](../LICENSE) (MPL-2.0, copyright 2026 Daniele Penza) covers original repository content only. It grants no rights over third-party material and does not decide whether the project may be distributed. It replaced the MIT License on 2026-10-09 (#49); copies obtained earlier under MIT keep that license. The root `LICENSE` file is the license notice for every file, as MPL-2.0 Exhibit A permits; files carry no per-file header.

**External distribution** means making the repository public, sharing a clone, archive or built workbook with anyone outside the maintainer, or publishing a GitHub Release. Before any of these, the owner reviews and records in an issue:

1. that the license is still the intended one;
2. that every file is original or carries a compatible license and attribution (the static-check tooling is adapted from the maintainer's own VBA-SACCR-Toolkit and EXCEL-VBA-PROJECT-TEMPLATE, relicensed here by their copyright holder);
3. that no employer, client or vendor has a claim on the content or on the methods it implements;
4. that a full-history scan finds no sensitive material (see below); and
5. that issue and pull-request text, which becomes visible with the repository, contains no confidential material.

No external distribution is authorized until that issue is closed by the owner.

### History scan

Scan every blob reachable from any branch, not only the current tree:

```shell
git fetch origin '+refs/heads/*:refs/remotes/origin/*'
git rev-list --objects --all
git log --all --numstat --format= | awk '$1 == "-"'   # binary files ever committed
```

Search each blob for credentials, keys, email addresses, institution and vendor names, and real-data markers; inspect every binary and every data file. Record the scanned commit range, the patterns and the findings in the issue.

## Definition of done

A change is complete only when (1) the issue contract is satisfied; (2) tests, diagnostics and failure behavior are addressed; (3) docs and changelog are current; (4) `python tools/check.py --ci` passes on the exact source commit; and (5) genuine Excel compilation, execution and independent result evidence is attached when relevant.

Static CI does not compile or execute VBA.
