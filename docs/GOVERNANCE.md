# Governance

## Workflow

- Active integration branch: the release branch named in
  [`RELEASING.md`](../RELEASING.md#current-status).
- Start focused task branches from the release branch; open PRs back into it
  and squash-merge them. The release-branch ruleset also allows a merge
  commit, used only to level a release branch with `main`.
  The release branch goes into `main` with a merge commit; see
  [`RELEASING.md`](../RELEASING.md#integrating-the-release-branch-into-main).
- No direct commits to `release/**` or `main` after the initial repository bootstrap; even levelling a release branch with `main` is a pull request ([`RELEASING.md`](../RELEASING.md#integrating-the-release-branch-into-main)).
- A release-to-main PR needs my explicit approval. No version tags or releases are implied.
- Merge only with review sign-off and green repository checks, as recorded in the PR's final-head review ([pull requests](#pull-requests)). VBA may be merged before Excel evidence exists; a release needs Excel evidence on its candidate ([`EXCEL_EVIDENCE.md`](EXCEL_EVIDENCE.md)).

<a id="pull-requests"></a>

## Pull requests

- **Linking issues.** A task PR names the issues it completes with
  `Resolves #N`. An integration PR (release branch into `main`), or any PR
  that only carries work merged earlier, names them with `Refs #N`: a closing
  keyword there would close an issue at the wrong time, and GitHub closes
  issues only when a PR merges into the default branch. Close a task issue by
  hand when its PR merges into the release branch. Milestones belong to
  issues, never to PRs.
- **Final-head review.** Before merging, the PR records its final head SHA,
  the commit the review covered and its result, and the Static repository
  checks run on that head. The reviewed commit must be the final head: any
  push after the record renews both the review and the CI run, and every
  review finding is fixed or answered, with its thread resolved. For a
  Dependabot PR, which does not use the template, I record the same in a
  comment.

<a id="issue-triage"></a>

## Issue triage

Every open issue carries what I need to plan it:

- **Assignee:** `danielep71`.
- **Priority:** exactly one of `P1`, `P2`, `P3`.
  - `P1` only for a confirmed calculation bug: the application gives wrong
    results for valid inputs. It always comes with `bug`.
  - `P2` for other material correctness defects (for example a broken import
    or a contract violation) and for robustness or maintainability work.
  - `P3` for non-blocking quality, testing, documentation or technical-debt
    improvements.
- **Milestone:** an open milestone with a due date.
- **Start date:** an ISO date (`YYYY-MM-DD`, Europe/Rome) on or before the
  target end date, as a line `- Start date: ...` or under a "Start date"
  heading, which is how the issue form renders it.
- **Target end date:** an ISO date (`YYYY-MM-DD`, Europe/Rome) before that
  due date, as a line `- Target end date: ...` or under a "Target end date"
  heading.
- **Dependencies:** a section headed "Dependencies". Each issue it names must
  exist; it may describe parents, children and related work freely. Only a
  line starting `Blocked by` names prerequisites (`Blocked by #N`); an
  `External condition: ...` line names an outside blocker. Write "None." when
  there is nothing to say.
- **`blocked`:** present exactly while a `Blocked by` line names an open issue
  or the section names an external condition.
- Labels come from `.github/labels.json`; titles carry no version prefixes,
  and acceptance criteria are verifiable checkboxes.

The issue forms set the assignee and a provisional priority (`P2` for bugs and
features, `P3` for documentation) and require the dependencies, a start date
and a target end date, which you propose: the start on or before the target,
the target before the milestone's due date. Each form also
requires acceptance criteria; the bug form requires the host environment and a
validation plan, and the documentation form requires the document and section,
the problem, the proposed correction, a supporting source and how the
correction is validated. I confirm or change the priority, set the milestone
and may change the start and target end dates during triage. No automation changes an
issue's priority or milestone.

`.github/workflows/issue-triage.yml` runs `tools/check_issues.py` from the
default branch on issue and milestone events and on demand, with read-only
permissions. It reports every finding in the job summary and fails when there
is any; I resolve each finding by hand. `python tools/check_issues.py
--check-config`, one of the portable gates, checks that the forms and the label
catalog still support these rules.

Where an automation files an issue itself (the traffic alert), it uses the
open milestone with the earliest due date after today; with none, the run
fails with a message saying so. Give the next milestone a due date before the
current one passes.

Close an issue only when its acceptance criteria are evidenced, with links to
the code and tests.

<a id="repository-settings"></a>

## Repository settings

The general settings, security features and rulesets are recorded in
[`repository-settings.json`](repository-settings.json), the baseline; this
document does not restate them. `tools/setup_repository.py` compares the saved
settings with it ([tools/README.md](../tools/README.md#repository-settings)).
After any settings change I run it with `--check` and keep the evidence file;
an unsaved edit or an authentication prompt does not establish enforcement.
The script never changes or deletes a ruleset the baseline does not name.
No ruleset has a bypass, for admins included: every change reaches `main` and
`release/**` through a pull request whose required check passed on a head that
is up to date with its base. "Protect release branches" accepts a squash (task
pull requests) or a merge commit (levelling a release branch with `main` after
an integration,
[`RELEASING.md`](../RELEASING.md#integrating-the-release-branch-into-main)).
"Prevent release branch deletion" forbids deleting `release/**`, so merging an
integration pull request never deletes the release branch.

## Privacy and intellectual property

Use only synthetic examples and independent work. The repository is public, and no visibility setting or private fork is authorization to store client files, personal data or confidential vendor methods. Third-party licensing and method provenance must be resolved before implementation. Contributor rules are in [`CONTRIBUTING.md`](../CONTRIBUTING.md#data-confidentiality-and-provenance).

## Licensing and distribution

The [Mozilla Public License 2.0](../LICENSE) (MPL-2.0, copyright 2026 Daniele Penza) covers original repository content only. It grants no rights over third-party material and does not decide whether the project may be distributed. It replaced the MIT License on 2026-10-09 (#49); copies obtained earlier under MIT keep that license. The root `LICENSE` file is the license notice for every file, as MPL-2.0 Exhibit A permits; files carry no per-file header.

**External distribution** means making the repository public, sharing a clone, archive or built workbook with anyone but me, or publishing a GitHub Release. Before any of these, I review and record in an issue:

1. that the license is still the intended one;
2. that every file is original or carries a compatible license and attribution (adapted files, their sources and licenses are listed in [`docs/licenses/THIRD_PARTY_NOTICES.txt`](licenses/THIRD_PARTY_NOTICES.txt));
3. that no employer, client or vendor has a claim on the content or on the methods it implements;
4. that a full-history scan finds no sensitive material (see below); and
5. that issue and pull-request text, which becomes visible with the repository, contains no confidential material.

No external distribution is authorized until I close that issue.

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
