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
| Release evidence tooling | Static checks (`tools/check.py`); development-workbook builder, synchronizer and source round trip ([`tools/README.md`](tools/README.md)) |
| Excel certification procedure | Stages 1–6 of [`docs/EXCEL_EVIDENCE.md`](docs/EXCEL_EVIDENCE.md), with the harness `TEST_Harness.RunTests` |

The step below marked *(to be defined)* is settled with the first distributed
artifact. The v0.1.0 foundation milestone closed without producing a release.

## 🌿 Branch model

| Branch | Role |
| --- | --- |
| `release/<version>` | Active integration branch. Task PRs are squash-merged into it after review and green checks. |
| `release/v0.3.0` | Opened early to hold the reviewed decay prototype (#36); takes no other work until v0.2.0 is released. |
| `main` | Default branch. Receives the release branch only through an owner-requested PR merged with a merge commit. |
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

The rulesets mirror VBA-SACCR-Toolkit's. `main` accepts changes only through
pull requests merged with a **merge commit**. `release/**` accepts pull requests
merged with a **squash** only, rejects non-fast-forward pushes and cannot be
deleted; repository admins may bypass it. Both require the **Repository
integrity** check.

1. Confirm the release branch is green and has no open review findings.
2. Open a PR from `release/vX.Y.Z` into `main`. `main` changes only through
   these integrations, so it has nothing the release branch lacks.
3. Merge it with a **merge commit**, not a squash, once the hosted checks are
   green, so provenance remains visible. The automatic head-branch deletion
   fails on the protected release branch, which is kept.
4. **Bring the release branch level with `main`:** the owner, as an admin
   bypassing the ruleset, fast-forwards it to the merge commit
   (`git push origin <merge-sha>:refs/heads/release/vX.Y.Z`). A merge commit or
   a squash of `main` into the release branch is not used: the ruleset allows
   only squash, which would cut the shared history. At a milestone closeout,
   open the next `release/vX.Y.Z` branch from the merge commit instead, and
   point the docs and `ISSUE_MILESTONE_NUMBER` at its milestone.

   **If the next release branch already exists**, as `release/v0.3.0` does,
   it is not recreated. The owner merges the merge commit into it locally with
   a merge commit, resolving any conflict in favour of the later branch's work,
   and pushes the result as an admin:
   `git switch release/vX.Y.Z && git merge --no-ff <merge-sha> && git push origin release/vX.Y.Z`.
   The push is a fast-forward of that branch, since the new commit has its tip
   as first parent, so nothing is reset or rewritten. CI must pass on the
   pushed commit before more work lands on it.
5. Verify that both branches point to the merge commit and that CI passed on
   it.

Never rebase, amend, force-push or reset either branch.

## 📦 Releasing a version

Performed only on the owner's explicit request.

The candidate is the merge commit that brings the release branch into `main`:
it is certified, any artifact is built from it, and the tag is put on it, so
all three name one SHA.

### 1. Finalize the changelog

Move the `[Unreleased]` entries into `## [X.Y.Z] - YYYY-MM-DD`, add the link
reference for the new version, keep an empty `[Unreleased]` section, and create
or update `VERSION` to `X.Y.Z` in the same PR. `tools/check_source.py` enforces
the heading format, real calendar dates, link references, newest-first order
and the `VERSION` match. Merge this through a task PR into the release branch.

### 2. Integrate into main and freeze the candidate

Follow [Integrating the release branch into main](#integrating-the-release-branch-into-main),
which leaves the release branch level with `main`. The merge commit on `main`
is the candidate: record its full SHA, the previous tag (if any) and the
intended version. Until the tag exists, the release branch takes only fixes for
a gate that failed on the candidate (steps 3 to 5); each is integrated again,
and the new merge commit becomes the candidate.

### 3. Run the static gates

```sh
git switch --detach <candidate-sha>
python tools/check.py --ci
```

Confirm the hosted **Repository integrity** check is green on the same SHA.

### 4. Certify in Excel

Build a clean workbook from the candidate, compile, run the full regression
harness, the reference cases and the release scenarios, and record the evidence
against the candidate SHA as described in
[`docs/EXCEL_EVIDENCE.md`](docs/EXCEL_EVIDENCE.md).

### 5. Build artifacts *(to be defined)*

If a workbook is distributed, build it from the candidate, test the packaged
file and record its SHA-256 hash. No artifact is distributed today.

### 6. Create the annotated tag

Tag only the candidate, the certified merge commit on `main`, from the command
line so the tag is annotated:

```sh
git fetch origin main
git rev-parse <candidate-sha>
git tag -a vX.Y.Z <candidate-sha> -F ../tag-message-vX.Y.Z.txt
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

### 7. Publish and verify

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
| 5 | Release branch merged into `main` with a merge commit; candidate SHA recorded | ☐ |
| 6 | Release branch fast-forwarded to the merge commit | ☐ |
| 7 | Static checks pass on the candidate locally and in CI | ☐ |
| 8 | Excel certification and reference cases pass on the candidate | ☐ |
| 9 | Artifacts built, tested and hashed (if any) | ☐ |
| 10 | Annotated tag targets the candidate | ☐ |
| 11 | GitHub Release published and verified | ☐ |

## 📚 Related documents

- [`CHANGELOG.md`](CHANGELOG.md) — history and version rules
- [`CONTRIBUTING.md`](CONTRIBUTING.md) — contribution workflow
- [`INSTALLATION.md`](INSTALLATION.md) — setup, import and validation record
- [`docs/GOVERNANCE.md`](docs/GOVERNANCE.md) — branches, issues and definition of done
- [`SECURITY.md`](SECURITY.md) — vulnerability handling
