<div align="center">

# 🚀 IRRBB Toolkit Release Guide

### Integration and release sequence for the IRRBB behavioral-model toolkit

[![Release model](https://img.shields.io/badge/Release-Exact_source-0969da?style=flat-square)](#release-invariants)
[![Versioning](https://img.shields.io/badge/Versioning-SemVer-3f4551?style=flat-square)](CHANGELOG.md#date-and-version-rules)
[![Status](https://img.shields.io/badge/Status-Pre--release-6e7781?style=flat-square)](#current-status)
[![Security](https://img.shields.io/badge/Security-Private-d73a49?style=flat-square)](SECURITY.md)

<br>

**Owner-approved · One frozen SHA · Certified in Excel · Annotated tag · Immutable once published**

</div>

---

This document is authoritative for the **maintainer integration and release
sequence**. Day-to-day contribution is owned by
[`CONTRIBUTING.md`](CONTRIBUTING.md); branch and review policy by
[`docs/GOVERNANCE.md`](docs/GOVERNANCE.md); version and changelog format by
[`CHANGELOG.md`](CHANGELOG.md#date-and-version-rules).

> [!IMPORTANT]
> Nothing is released, versioned, tagged or stamped unless the owner asks for it
> explicitly. Closing a milestone does not authorize a release, and unreleased
> work stays under `## [Unreleased]` in the changelog.

<a id="current-status"></a>

## 🧭 Current status

| Property | State |
| --- | --- |
| Released versions | None |
| Active branch | `release/v0.2.0` |
| Active milestone | `v0.2.0 - Workbook, data and Excel tests` |
| Closed milestones | `v0.1.0 - Repository foundation` (2026-10-09, no release) |
| Version file | Not present; created when the first release is prepared ([versioning](#versioning)) |
| Release evidence tooling | Static checks (`tools/check.py`) |
| Excel certification procedure | Policy in [`docs/EXCEL_EVIDENCE.md`](docs/EXCEL_EVIDENCE.md); harness *(to be defined)* |

Steps below marked *(to be defined)* are settled with the first VBA source and
the first distributed artifact. The v0.1.0 foundation milestone closed without
producing a release.

## 🌿 Branch model

| Branch | Role |
| --- | --- |
| `release/<version>` | Active integration branch. Task PRs merge into it after review and green checks. |
| `main` | Default branch. Receives the release branch only through an owner-requested PR. |
| Task branches | One per change, from the release branch: `fix/<issue>-<slug>`, `docs/<slug>`, `test/<slug>`. |

<a id="versioning"></a>

## 🏷️ Versioning

| Item | Convention |
| --- | --- |
| Version | Semantic Versioning `MAJOR.MINOR.PATCH`; rules in [`CHANGELOG.md`](CHANGELOG.md#date-and-version-rules) |
| Release branch | `release/vX.Y.Z`, matching the milestone, opened only by owner decision |
| Milestone | `vX.Y.Z - <title>`, matching the release branch |
| `VERSION` | Root file holding one `X.Y.Z` line. Created in the release-preparation PR, never before; afterwards it always names the newest released version |
| Changelog | Work stays under `## [Unreleased]` until release preparation is approved; the release PR moves it to `## [X.Y.Z] - YYYY-MM-DD` |
| Tag | Annotated `vX.Y.Z` on the certified commit on `main` |
| GitHub Release | Created from the tag, with certified artifacts only |

`python tools/check.py` enforces the link between `VERSION` and the changelog:
once the changelog has a dated release, `VERSION` must exist and equal the
newest one; a `VERSION` without a dated release fails.

<a id="release-invariants"></a>

## 🔒 Release invariants

A release is valid only when:

1. one exact candidate SHA is frozen and reviewable;
2. the changelog entry for the version is dated and matches the scope;
3. `python tools/check.py --ci` passes on that candidate, and the hosted
   **Repository integrity** check is green;
4. VBA compile, the synthetic regression harness and the release scenarios pass
   in Excel on that candidate;
5. every implemented calculation has independent numerical reference cases
   within stated tolerances, recorded against that candidate;
6. any distributed artifact is built from and tested against that candidate;
7. the licensing and distribution review in
   [`docs/GOVERNANCE.md`](docs/GOVERNANCE.md#licensing-and-distribution) is
   recorded and closed by the owner;
8. an annotated `v*` tag targets the certified commit on `main`; and
9. post-publication checks pass.

If source changes after certification, the evidence is stale and must be
rerun. Never compensate by editing an already-tested artifact.

<a id="integrating-the-release-branch-into-main"></a>

## 🔀 Integrating the release branch into main

Done only when the owner asks. This is an integration, not a release.

Both `main` and `release/**` accept changes **only through pull requests**
with the **Repository integrity** check green; direct pushes, including
fast-forwards, are rejected by the branch rulesets.

1. Confirm the release branch is green and has no open review findings.
2. **Always create an integration branch** from the release tip and use it as
   the PR head, never `release/vX.Y.Z` itself: the repository deletes a PR's
   head branch after merging, which would remove the release branch.
3. **Bring `main` in if it has moved.** If `main` has commits the release
   branch lacks, merge `main` into the integration branch with a merge commit
   and resolve any conflict so the result keeps the release branch's content.
4. Open the PR into `main`, wait for **Repository integrity** and the review
   findings, then merge it with a **merge commit**, not a squash, so both
   branches share history.
5. **Bring the release branch level with `main`** with a second PR from `main`
   into `release/vX.Y.Z`, merged with a **merge commit**. `main` then becomes an
   ancestor of the release branch, and the next integration merges without
   conflicts. At a milestone closeout, open the next `release/vX.Y.Z` branch
   from the merge commit instead, and point the docs and
   `ISSUE_MILESTONE_NUMBER` at its milestone.
6. Verify that the release branch contains the `main` merge commit, that both
   trees are identical, and that CI passed on both.

Never rebase, amend, force-push or reset either branch.

## 📦 Releasing a version

Performed only on the owner's explicit request.

### 1. Freeze and identify the candidate

Record the release-branch SHA, the previous tag (if any) and the intended
version.

### 2. Finalize the changelog

Move the `[Unreleased]` entries into `## [X.Y.Z] - YYYY-MM-DD`, add the link
reference for the new version, keep an empty `[Unreleased]` section, and create
or update `VERSION` to `X.Y.Z` in the same PR. `tools/check_source.py` enforces
the heading format, real calendar dates, link references, newest-first order
and the `VERSION` match. Merge this through a task PR into the release branch.

### 3. Run the static gates

```sh
python tools/check.py --ci
```

Confirm the hosted **Repository integrity** check is green on the same SHA.

### 4. Certify in Excel *(to be defined)*

Import the exact candidate into a clean workbook, compile, run the full
regression harness, the reference cases and the release scenarios, and record
the evidence against the candidate SHA as described in
[`docs/EXCEL_EVIDENCE.md`](docs/EXCEL_EVIDENCE.md).

### 5. Build artifacts *(to be defined)*

If a workbook is distributed, build it from the certified source, test the
packaged file and record its SHA-256 hash. No artifact is distributed today.

### 6. Integrate into main

Follow [Integrating the release branch into main](#integrating-the-release-branch-into-main).
If the merge changes the source identity, repeat the certification on the
merged SHA.

### 7. Create the annotated tag

Tag only the certified commit on `main`, from the command line so the tag is
annotated:

```sh
git switch main
git pull --ff-only
git rev-parse HEAD
git tag -a vX.Y.Z -F ../tag-message-vX.Y.Z.txt
git cat-file -t vX.Y.Z        # must print: tag
git push origin vX.Y.Z
```

The tag message is the release's permanent record:

```text
vX.Y.Z - <short release title>

Certified at <full 40-character SHA>.
Excel: compile <result>; harness <passed>/<run>, Failed=0.
Reference cases: <passed>/<run> within tolerance.
Fixes #<n>, #<n>.
Deviations: none.
```

`Deviations:` is mandatory: write `none`, or list each gate that was not run.
Copy totals verbatim from the run; never round or reuse earlier figures. Never
delete and recreate a published tag.

### 8. Publish and verify

Publish the GitHub Release from the tag, attach only certified artifacts with
their hashes, and verify that a fresh clone of the tag passes the static checks
and imports cleanly.

## 🧾 Evidence record

Retain at least: candidate SHA, previous tag, static-check result, Excel
version and bitness, compile and harness results, reference-case results and
tolerances, artifact names, sizes and SHA-256 hashes, the integration PR,
validation timestamps and any deviations. A release note summarizes evidence;
it does not replace it.

## 🧯 Recovery

- **Before tagging:** fix through a task PR, update the evidence and rerun the
  affected gates.
- **After tagging, before publishing:** pause and document the correction;
  prefer a new patch version if the tag may have propagated.
- **After publishing:** do not replace assets or move the tag. Publish a
  corrected patch, record the problem in the changelog, and use
  [`SECURITY.md`](SECURITY.md) for vulnerabilities.

## ☑️ Maintainer checklist

| # | Gate | Done |
| ---: | --- | :---: |
| 1 | Owner requested the release | ☐ |
| 2 | Scope frozen and diff reviewed | ☐ |
| 3 | Licensing and distribution review recorded and closed | ☐ |
| 4 | Changelog finalized and `VERSION` created or updated | ☐ |
| 5 | Static checks pass locally and in CI | ☐ |
| 6 | Excel certification and reference cases pass | ☐ |
| 7 | Artifacts built, tested and hashed (if any) | ☐ |
| 8 | Release branch merged into `main` with a merge commit | ☐ |
| 9 | Release branch level with `main` again (PR from `main`, merge commit) | ☐ |
| 10 | Annotated tag targets the certified `main` SHA | ☐ |
| 11 | GitHub Release published and verified | ☐ |

## 📚 Related documents

- [`CHANGELOG.md`](CHANGELOG.md) — history and version rules
- [`CONTRIBUTING.md`](CONTRIBUTING.md) — contribution workflow
- [`INSTALLATION.md`](INSTALLATION.md) — setup, import and validation record
- [`docs/GOVERNANCE.md`](docs/GOVERNANCE.md) — branches, issues and definition of done
- [`SECURITY.md`](SECURITY.md) — vulnerability handling
