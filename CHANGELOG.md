<div align="center">

# 📜 Changelog

### Release history for K-IRRBB, IRRBB behavioral models in Excel/VBA

[![Format](https://img.shields.io/badge/Format-Keep_a_Changelog-0969da?style=flat-square)](https://keepachangelog.com/en/1.1.0/)
[![Versioning](https://img.shields.io/badge/Versioning-SemVer-6f42c1?style=flat-square)](https://semver.org/spec/v2.0.0.html)
[![Dates](https://img.shields.io/badge/Dates-YYYY--MM--DD-217346?style=flat-square)](#date-and-version-rules)
[![Staging](https://img.shields.io/badge/Staging-Unreleased_first-d97706?style=flat-square)](#unreleased)
[![Contributing](https://img.shields.io/badge/Changes-Contribution_guide-2ea44f?style=flat-square)](CONTRIBUTING.md)

<br>

**User-visible history · Explicit compatibility · Reproducible evidence · Immutable releases**

</div>

---

All notable changes to **K-IRRBB** are documented here.

This changelog follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/)
and [Semantic Versioning](https://semver.org/spec/v2.0.0.html). It records
released behavior and material unreleased changes; it is not a commit log, issue
tracker, or substitute for release evidence.

---

## 🧭 Maintenance policy

- Add material changes under **Unreleased** in the same pull request as the
  behavior or documentation they describe.
- Write from the user's perspective: describe the observable result, contract,
  compatibility impact, and migration need.
- Link the owning issue or pull request when it contains useful engineering
  detail.
- Record only validation actually performed. Static checks are never Excel
  evidence; state what was and was not run in Excel.
- Record any change to estimated parameters, stable shares, decay profiles or
  backtest statistics for unchanged inputs, with its cause.
- Move Unreleased entries into a dated version section only when a release is
  explicitly prepared.
- Do not edit a published release entry except to correct a demonstrable factual
  or link error; annotate material corrections instead of rewriting history.

See [CONTRIBUTING.md](CONTRIBUTING.md) for change and evidence requirements and
[SECURITY.md](SECURITY.md) for private vulnerability reporting.

<a id="date-and-version-rules"></a>

### Date and version rules

| Rule | Standard |
|---|---|
| Version | `MAJOR.MINOR.PATCH`, without the leading `v` in headings |
| Release heading | `## [X.Y.Z] - YYYY-MM-DD` |
| Date | Gregorian calendar date in ISO `YYYY-MM-DD` format |
| Ordering | Unreleased first; released versions newest to oldest |
| Comparison | Unreleased → `main` against the active release branch; each release → preceding tag |
| Patch | Backward-compatible correction or hardening |
| Minor | Backward-compatible capability |
| Major | Incompatible public-contract change |

The repository remains below `1.0.0` while its supported surface is still
forming. Pre-release status does not excuse undocumented breaking changes.
`python tools/check.py` enforces these rules and the match with `VERSION`.

<details>
<summary><strong>Entry categories</strong></summary>

<br>

| Category | Use for |
|---|---|
| **Added** | New supported capabilities, APIs, files, or tests |
| **Changed** | Changes to existing behavior, contracts, tooling, or documentation |
| **Deprecated** | Supported behavior scheduled for removal |
| **Removed** | Removed capabilities or compatibility |
| **Fixed** | Corrected defects |
| **Security** | Safely disclosed security corrections |
| **Documentation** | Material documentation-only changes |
| **Validation** | Evidence actually produced |
| **Compatibility** | Upgrade or migration effects |
| **Known limitations** | Deliberate, unresolved boundaries |

Use only the categories needed by a release.

</details>

---

<a id="unreleased"></a>

## [Unreleased]

> Not yet released. Development takes place on the active release branch,
> named in [`RELEASING.md`](RELEASING.md#current-status); changes to `main`
> need my explicit instruction.

### Added

- Workbook synchronization and source round trip (#6), following
  VBA-SACCR-Toolkit: `tools/Sync-IRRBB-VBA.ps1` with the launchers
  `Sync-IRRBB-VBA.cmd` (drag a workbook onto it to import the checkout's source)
  and `Sync-IRRBB-VBA-Downloads.cmd` (import modules from a folder, by default
  Downloads, without changing the checkout) updates the VBA project of an
  existing `.xlsm` in production or development mode without touching sheets;
  `tools/Export-IRRBB-VBA.ps1` exports a workbook's components with an
  `export-manifest.json`, and `tools/compare_vba_export.py` checks them against
  the source of one commit. `tools/vba-projects.json` lists both workbooks'
  components. The document modules `ThisWorkbook.cls` and one `sh<Sheet>.cls`
  per template sheet are now source files in `src/workbook/`; the development
  builder embeds them. `check_source.py` requires every component to lie in a
  synchronized directory and the document modules to match the template's
  CodeNames; `docs/EXCEL_EVIDENCE.md` adds the round trip as stage 6.
- **Decay model** per `MODEL_CONTRACTS.md` (draft `DECAY-1.0-draft`, not
  validated; review in #36): host-independent core `CORE_Math`, `CORE_Decay`
  and `CORE_Overlay`, and the public facade `IRRBB_Decay` (`IRRBB_DecayFit`,
  `IRRBB_DecayOverlay`, `IRRBB_DecayBacktest`, layout enums), listed in
  `docs/PUBLIC_API.txt`. Log changes with declared impulse breaks, residual
  sigma with n - 1 - K degrees of freedom, prudential drift, minimum probable
  amount at confidence c, conserved runoff profile, core and non-core split,
  mean lives in months, Ljung-Box and Jarque-Bera diagnostics with the
  empirical quantile, the `REG-1` overlay (non-core overnight, core-share cap,
  average-maturity cap by cutoff reduction; `REG-2` shown provisional and not
  applied) and the out-of-sample exceedance score with frozen parameters.
  Input is rejected below 36 log changes, on non-month-end dates, gaps,
  non-positive balances, unknown segments and currencies outside the run
  allowlist; the backtest rejects a zero balance at a forecast origin and
  accepts it only as the last realized balance.
- Decay cases `DEC-LOG-01`, `DEC-MPA-01`, `DEC-PROF-01`, `DEC-LIFE-01`,
  `DEC-CAP-01`, the new `DEC-E2E-01` and `DEC-ERR-01`, and the decay parts of
  `OOS-FREEZE-01`, `OOS-LEAK-01` and `OOS-SCORE-01`, run by the harness through
  `TEST_DecayCases`; synthetic fixtures with generator and seed
  (`tests/fixtures/decay/`); independent expected values from numpy and scipy
  with provenance (`tests/expected/decay/`), generated into
  `tests/modules/TEST_DecayData.bas`, which `tools/test_vba_harness.py` now
  keeps equal to the expected files.
- Example `examples/modules/EX_DecayExample.bas` (facade only). It reads a
  `DecayExample` sheet, which the template-based development workbook does not
  have yet (#36); `lo_smoke.py` runs the example and its invalid-input path
  when the sheet exists.
- Deliberate failure run (#8): `TEST_Harness.TEST_RunTestsWithInjectedFailures`
  runs every case with one failing assertion added to `HARN-OUT-01` and an
  unexpected error raised in `DATA-CODES-01`, so an Excel run can show that a
  failure ends `COMPLETE - FAILURES` with those cases `FAIL` and `ERROR`, never
  `PASS`, and that the next `RunTests` withdraws them. `docs/EXCEL_EVIDENCE.md`
  stage 4 now names this run and `TEST_RunTestsWithWriteFault`, and stage 3
  says to copy the `Checks` results. `tools/devbuild/lo_smoke.py` runs it. Not
  run in Excel.
- Per-period outcome lineage (#71): an accepted import now writes, per month
  end, segment and currency, the transitions from that month end by outcome
  (observed, verified closures, gaps, censored disappearances with the sum of
  their last balances, end of sample) and the accounts whose `close_date` the
  point-in-time view at that month end masks, to the new `DataSummary_Outcomes`
  area (`DataSummary!S14:AB20013`) through the new `IRRBB_ImportOutcomes`.
  The data contract now says how a masked account is counted (under its
  latest segment and currency on or before the origin). `DATA-PIT-01` and
  `DATA-IMPORT-01` check it against hand-assigned values for
  `availability.csv`. Not run in Excel.
- Segment selection (#71): the segments marked `Yes` on `Segments` form the
  run's segment list, recorded in the lineage. The import still validates and
  reconciles every segment, so leaving one out never turns a migration into a
  gap or a disappearance; the contract applies the selection when a model
  builds its sample (a transition enters only when its origin segment is
  selected). `IRRBB_ImportData` takes the list as a new argument; a run with
  no segment is refused. `tools/devbuild/make_large_panel.py` writes a 47 MB
  synthetic panel for timing the import in Excel.
- Data fingerprint (#71, stage 4): `CORE_Sha256` computes SHA-256 in VBA
  (DATA_CONTRACT.md decision 5) while the import reads each file, and the
  lineage records both fingerprints and the toolkit commit, taken from the
  workbook's build stamp (`Readme_SourceCommit`). `IRRBB_ImportData` takes the
  commit as a new last argument. `DATA-SHA-01` checks the published FIPS
  180-2 examples and padding boundaries. The import time is local, without
  UTC offset (DATA_CONTRACT.md decision 9). Not run in Excel; speed on large
  panels not measured.
- Import from files (#71, stage 3): `IRRBB_ImportData`, the first public
  entry point, reads the account panel and market-rate CSV files (read only,
  other writers locked out while reading, streamed in 64 KB chunks, LF or
  CRLF), refuses a window whose dates are not month ends without a time of
  day, refuses to run without the timely closure reporting attestation, and
  returns totals, findings, warnings and
  lineage as tables. **Import data** on `Config` writes them to `DataSummary`
  under the status protocol: earlier results withdrawn first, totals only
  for an accepted import, primary and cleanup errors kept. `DATA-IMPORT-01`
  runs it end to end. `SECURITY.md` records the file access. Not run in Excel.
- Transition outcomes and point-in-time view (#71, stage 2): the import
  summary assigns every observed month its outcome (observed, verified
  closure, gap, censored disappearance, end of sample), and `CORE_ViewBuild`
  gives the rows known at a forecast origin with later close dates masked.
  The view exists only after an accepted import (no finding, no `E12`) and
  keeps to the data window.
  `DATA-PIT-01` and `DATA-PIT-02` now run in VBA: the hand-assigned outcomes
  and visible closures, and a leakage test showing that changing everything
  learned after an origin leaves the view at that origin unchanged. Not run
  in Excel.
- VBA import core (#71, stage 1): `CORE_Import` validates the account panel
  and the market-rate file line by line in the contract's rule order (`E01`–
  `E11`, cross-row `E11`, `E10`, `E05`, `E08`), and for a data window reports
  warnings `W01`–`W06`, totals per date, segment and currency (compensated
  sums, never across currencies) and missing extracts (`E12`).
  `DATA-VALID-01` and `DATA-INVALID-01` now run in VBA against the
  hand-computed expected files. No file input, workbook output or lineage
  yet; not run in Excel.
- Workbook template (#5): `src/workbook/IRRBB_Template.xlsx`, macro-free and
  sanitized, with ten sheets, their CodeNames, named input and output areas,
  input validation and a status cell per output sheet that starts as not
  current. `docs/WORKBOOK_CONTRACT.md` defines the layout and the result
  status protocol; the template is generated from
  `tools/workbook/build_template.py` and checked against it, the contract and
  the harness by `tools/test_workbook_template.py`. Placeholder sheets are
  marked as such. Not yet opened in Excel.
- `WB-LAYOUT-01`: a harness case that checks the sheets, CodeNames and names of
  the built workbook through Excel.
- Regression harness for Excel (#8): `TEST_Harness.RunTests` runs every case
  in `docs/methodology/TEST_CASES.md` in register order and writes per-case and
  per-assertion results to the `Checks` sheet. Earlier results are withdrawn
  before a run; a case with no assertion or an unexpected error never passes; an
  unbuilt case is `NOT RUN`. If writing the results fails, the run ends
  `FAILED` with both the primary and the cleanup error shown, and no earlier
  result remains. First VBA cases: `HARN-OUT-01` (outcome and
  tolerance rules), `DATA-PARSE-01`, `DATA-PARSE-02` and `DATA-CODES-01`.
  `tools/test_vba_harness.py` checks their expectations against the contract
  and the registry against the register without Excel. No Excel run is
  recorded yet.
- First core modules: `CORE_Parse` parses the data contract's dates,
  decimals and rates without Excel's locale-dependent conversions, returning
  missing and invalid separately; `CORE_Codes` holds the segment codes and the
  currency allowlist.
- Development workbook builder (`tools/devbuild/`, import route for #6):
  writes `build/IRRBB_Dev.xlsm`, the workbook template with the VBA project
  generated from source, so no VBA trust setting is needed; standard library
  only; checked by `tools/test_devbuild.py` with an independent decoder.

- Daily repository traffic export with history on a separate branch,
  main-only analytics access, and traffic alerts, following the SA-CCR
  repository setup.
  Setup and verification are documented in `docs/TRAFFIC.md`.

- `tools/check_source.py` rejects Excel object-model and UI identifiers
  (`Range`, `Cells`, `Application`, `MsgBox` and others) in `src/core/` code,
  ignoring comments and string literals, so the model core stays
  host-independent; covered by `tools/test_core_boundary.py` (#2).

- Initial repository foundation for a source-first Excel/VBA IRRBB behavioral
  modeling application.
- Documentation of model scope, validation requirements, source boundaries and
  synthetic-data policy.
- Portable static-checking tools and CI configuration adapted from my SA-CCR
  repository.

- CodeQL scanning of the workflows, the label scripts and the Python tooling
  on pushes, pull requests (analysed without uploading), weekly and on demand;
  an offline Markdown link check in the portable gates (`tools/check_links.py`)
  with a weekly run that also requests external links; and a workflow that
  removes milestones from pull requests, since milestones belong to issues.
  Each comes with offline tests (#135).

### Changed

- No ruleset can be bypassed any more, by admins included. "Protect main" and
  "Protect release branches" lose their admin bypass; release branches now
  require pull requests that are up to date with their base, and accept a
  squash (task pull requests) or a merge commit. After an integration, a
  release branch is levelled with `main` through a merge-commit pull request
  from a task branch that merges `main` into it, instead of an admin push. `tools/setup_repository.py` refuses a baseline
  ruleset with bypass actors. RELEASING.md, GOVERNANCE.md, CONTRIBUTING.md and
  tools/README.md describe the procedure (#160).
- The issue forms require a target end date: you propose one before the
  milestone's due date, and I may change it during triage. The bug form adds
  required fields for the host environment (Windows and Excel version, Office
  bitness, locale), acceptance criteria and a validation plan, and an optional
  field for evidence, workaround and recovery. The documentation form replaces
  its two combined fields with required fields for the document and section,
  the problem type, the problem, the proposed correction, a supporting source,
  acceptance criteria and validation. The feature form adds an optional
  "Alternatives considered" field. `tools/check_issues.py --check-config` fails
  when a form does not require the target end date or the dependencies (#161).
- `P1` is now reserved for a confirmed calculation bug that gives wrong
  results for valid inputs; other correctness defects, such as a broken
  import or a contract violation, are `P2`. The label description, the triage
  message, GOVERNANCE.md, LABELS.md and CONTRIBUTING.md use that definition
  (#162).
- mypy also reports unreachable code (`warn_unreachable`, which `strict` does
  not include), and CI proves the setting is active with a probe; the tooling
  had no such code (#166).
- Dependabot pull requests are assigned to `danielep71`, so they meet the
  triage rules without a manual step (#164).
- The daily traffic export reads the traffic API with a short-lived token of
  a GitHub App that has only `Administration: read`, created in each run from
  the `analytics` environment's `TRAFFIC_APP_CLIENT_ID` variable and
  `TRAFFIC_APP_PRIVATE_KEY` secret, instead of the `TRAFFIC_TOKEN` personal
  access token, which is no longer read and no longer needs rotating. Its
  inline Python moves to `tools/traffic_history.py`, with offline tests in
  `tools/test_traffic_history.py`; the `traffic-history` files, columns and
  badges are unchanged, and a collection replayed after a later one no longer
  overwrites it. A `collect` job with `contents: write` pushes the history
  without persisted checkout credentials, and a separate `alert` job with
  `issues: write` and no App token reads the pushed history. The schedule
  moves to 06:17 UTC, and a clone spike no longer raises an alert, since CI
  checkouts count as clones; clones stay in the CSV files, the alert body and
  the badge (#137).
- The repository settings and rulesets are a checked baseline,
  `docs/repository-settings.json`, instead of prose. `tools/setup_repository.py`
  previews the difference without writing, applies it with read-back
  (`--apply`) or verifies read-only (`--check`), writes a dated evidence file,
  refuses to run without admin rights and never touches a ruleset the baseline
  does not name. The baseline adds a ruleset for `traffic-history`, turns off
  rebase merging and Dependabot security updates, and keeps the existing
  rulesets for `main`, `release/**` and `v*` tags unchanged. GOVERNANCE.md and
  RELEASING.md point to it (#138).
- Issue triage reports instead of defaulting. The issue-metadata workflow,
  which added `P3`, dropped extra priorities and set a default milestone, is
  replaced by the read-only Issue triage workflow and `tools/check_issues.py`.
  It reports each open issue that lacks an assignee, exactly one priority, an
  open milestone with a due date, a target end date before it or a
  dependencies section, that has `P1` without `bug`, a label outside the
  catalog, or `blocked` out of step with its prerequisites. The issue forms
  keep a provisional priority and ask for a target end date and the
  dependencies. The traffic alert files its issue under the open milestone
  with the earliest due date after today, so the `ISSUE_MILESTONE_NUMBER`
  variable is no longer read (#136).
- The active release branch is named only in RELEASING.md's current-status
  table, in `.github/dependabot.yml` and in CHANGELOG.md; the other documents
  link to that table. `tools/test_workflow_pins.py` fails when another file
  names the branch, and Static repository checks fails when a Dependabot
  `target-branch` is not an existing branch (#134).
- CI runs one exact Python, 3.13.18, in every workflow, including the traffic
  and issue-metadata workflows that used the runner's `python3`. A non-required
  job runs the portable gates on the minimum supported Python, 3.11, which
  replaces 3.10 (end of life) in `pyproject.toml` and the documentation.
  `tools/requirements-quality-ci.txt` moves to Ruff 0.17.3 and mypy 2.4.1,
  hash-locked for CPython 3.13, and Dependabot now proposes their updates as
  one grouped pull request. `tools/test_workflow_pins.py` checks the versions,
  the setup in every job that runs Python and every Dependabot target (#131).
- The pull request template says to use `Resolves #` on task PRs and `Refs #`
  on integration PRs and PRs that only carry earlier work, so that no issue
  closes at the wrong time, and adds a final-head review: the final head SHA,
  the reviewed commit and its result, and the Static repository checks run,
  renewed after any push. `docs/GOVERNANCE.md` describes both under "Pull
  requests" (#133).
- CI now installs Ruff and mypy from the hash-locked
  `tools/requirements-quality-ci.txt` and runs `ruff check tools` with a
  complexity-ceiling probe, and checksum-verified actionlint 1.7.12 on the
  workflows, as VBA-SACCR-Toolkit does, and `mypy --strict` on every module in
  `tools/`, tests included, with a strict-mode probe; `tools/` is clean under
  both (#115).
- Dependabot proposes weekly updates of the pinned GitHub Actions as pull
  requests against the active release branch, labelled `ci` and `P3`, as in
  VBA-SACCR-Toolkit; `actions/upload-artifact` is updated to v7.0.2 and every
  action pin carries its version comment (#116).
- The project is named K-IRRBB, like its sibling K-SACCR: the README heading,
  the governance documents, the PR template, the traffic alert and the title
  on the workbook's Readme sheet use the new name. The repository was
  subsequently renamed, from `VBA-IRRBB-Toolkit` through
  `K-IRRBB_Behavioral_Models` to `K-IRRBB-NMDs`; clone instructions, badges,
  issue and milestone links and security/help links follow, and GitHub
  redirects the earlier names. Identifiers, file names and the `IRRBB_`
  prefix are unchanged.
- The workbook template follows my SACCR visual identity:
  black-and-gold `K-IRRBB` branding, blue title bands, Aptos typography,
  distinct input/output headers, colored tabs, consistent spacing and landscape
  print setup, with a short Readme workflow, section labels on DataSummary and
  Checks, and readable amount, count and diagnostic formats. Development-workbook buttons share the
  palette and avoid the status block. The template remains reproducible from
  its generator; sheet order, CodeNames, named ranges, validations, status rules
  and calculation behavior are unchanged. Native Excel visual verification is
  still required.
- The data contract now states that dates have a year from 1900 to 9999 (the
  Excel date system) and that a decimal has digits on both sides of any `.`
  (#8).
- The repository is now licensed under the Mozilla Public License 2.0
  (MPL-2.0) instead of MIT (#49). `LICENSE` holds the MPL-2.0 text and is the
  license notice for every file; contributions are accepted under MPL-2.0.
  Copies obtained before the change keep the MIT License.
- The active release branch is now `release/v0.2.0`, opened from `main` when
  milestone v0.1.0 closed (#55); RELEASING.md's current-status table names
  it. `release/v0.1.0`, opened after the foundation was integrated (PR #13),
  carried the foundation work; release branches are named `release/vX.Y.Z`.

### Removed

- The one-off `Bootstrap v0.1.0 tracking` workflow and
  `tools/bootstrap_tracking.py`. The v0.1.0 milestone and issue metadata they
  created remain; the workflow no longer runs with `issues: write` on every
  push to the release branch.

### Fixed

- The settings baseline now records the `analytics` environment's deployment
  branches (`main` only) and the CodeQL default setup state (`not-configured`).
  `tools/setup_repository.py --check` verifies both and lists them in the
  evidence file; `--apply` restricts the environment to the baseline branches,
  keeping its wait timer and required reviewers and never touching its secrets
  or variables, and turns the default setup off (#198).
- `WB-LAYOUT-01` counts only workbook-level defined names. The template also
  sets 13 sheet-level print titles and areas, which Excel lists in
  `Workbook.Names` (LibreOffice does not), so the expected 24 would have failed
  in Excel (#193).
- INSTALLATION and the CONTRIBUTING review checklist no longer require Excel
  evidence for every VBA change before merge: as GOVERNANCE states, a VBA
  change records Excel results when run, or states them as not run, and a
  release candidate needs them (#191).
- The release checklist no longer asks for a fast-forward of the release
  branch, which the rulesets refuse: it asks for the levelling pull request of
  step 4. RELEASING and CONTRIBUTING state that this pull request is the one
  merged with a merge commit, and name its branch `fix/level-<release>-after-<PR>`
  (#190).
- The decay facade and the data contract name the same source for the decay
  balance series: in production the per-segment totals of the import
  (`IRRBB_ImportTotals`), never a series the user supplies; tests and examples
  pass synthetic series (#189).
- The data contract no longer requires a reconciliation the import does not
  run: it states that the import reports per date, segment and currency
  totals, each account row counted once, and takes no control totals. The
  check against control totals from the source system is tracked in #202
  (#188).
- Dependabot updates every `github/codeql-action` step in one grouped pull
  request, since CodeQL fails when `init` and `analyze` differ in version; both
  steps are pinned to v4.39.3, in place of the separate #180 and #181 (#183).
- The settings check no longer treats paused Dependabot security updates as
  in effect: when the baseline wants them on but GitHub has paused them,
  `--check` reports drift and `--apply` fails its read-back. When the baseline
  wants them off, enabled but paused updates count as on and are turned off,
  since they resume on their own. The evidence file shows "on (paused)"
  (#165).
- The Markdown summaries of label synchronization and label drift no longer
  break on a label name, color, description, path or error message that holds
  a backtick, a pipe or a line break. A shared helper,
  `.github/scripts/markdown.mjs`, writes each value as a code span whose fence
  is longer than any run of backticks inside it, escapes pipes in table cells
  and turns CR, LF and CRLF into a space; the drift check reads the
  reconciler's plan table with the same rules. Both self-tests cover these
  values, and the label workflows now also run when the helper changes (#163).
- The link check no longer accepts plain `http://` links. The offline check,
  part of `tools/check.py` on every push, reports each external `http://` link
  in the Markdown as "insecure link (use https)"; `https://` and
  protocol-relative `//host` links stay allowed. The weekly `--external` run
  requests HTTPS only, also on redirects, and reports a redirect to another
  protocol as an error instead of following it without TLS (#168).
- Static repository checks no longer cancel the run of an earlier push to the
  same branch: only a newer commit on the same pull request cancels a run, and
  every other run keeps its own group, so each pushed commit keeps its
  result (#167).
- The label reconcile job runs only from the default branch: a manual
  dispatch from another branch no longer applies a manifest that has not been
  integrated (#169).
- Merging an integration pull request no longer deletes the release branch:
  "delete head branches on merge" had removed `release/v0.2.0` when its
  integration pull request merged into `main`. The settings baseline now has
  "Prevent release branch deletion", a deletion rule for `release/**` with no
  bypass (#158).
- Label synchronization sends its API requests to the host GitHub Actions
  names in `GITHUB_API_URL` (trailing slashes removed), falling back to
  `https://api.github.com`, instead of a hard-coded address. Its self-test now
  also runs a full reconcile against a local API, which checks that a label
  in use is never deleted and an unused one is (#152).
- The daily traffic export no longer loses a collection when
  `traffic-history` moves during the run: a rejected push is retried up to
  three times, each time rebuilding the history on the latest branch from the
  same collection. Every successful exit, including "no change", reports the
  pushed commit, and the alert job fails unless that report is a full commit
  SHA, so an empty report cannot make it check out the default branch (#153).
- Label reconciliation no longer deletes labels that are still in use. Before
  any change, each label that pruning would delete is checked against every
  issue and pull request, open and closed (names with a comma by reading all
  items, since the labels filter cannot express them); if any is in use, the
  run refuses with one message listing them and changes nothing. The reconcile
  job also holds `pull-requests: read`, so labels used only on pull requests
  are seen (#151).
- Issue triage reads a target end date under a level-2 "Target end date"
  heading, as I write it, and treats only issues on a `Blocked by` line as
  prerequisites; other issues named in a dependencies section are context and
  need only exist. GOVERNANCE.md, LABELS.md and the issue forms describe the
  convention (#146).
- The workbook template gives DataSummary, which the data import writes, the
  status `NOT RUN` instead of `NOT RUN - not implemented`; only placeholder
  sheets now say they are not implemented, as `docs/WORKBOOK_CONTRACT.md`
  states. `TEST_DecayCases` raises the harness's `TEST_ERR_NO_CASE` for an
  unknown case, as the other case modules do.
- Comments that no longer matched the code, or came from the project template
  or VBA-SACCR-Toolkit unadapted, are corrected across the VBA source, tests,
  Python tooling, workflows, the synchronizer and the repository control files;
  `.gitattributes` marks `docs/assets/` as documentation instead of `assets/`
  and `images/`, which do not exist.
- The issue-metadata and traffic-alert workflows no longer fall back to
  milestone 1, the closed v0.1.0, when the repository variable
  `ISSUE_MILESTONE_NUMBER` is unset; they fail with a message naming the
  variable, and a value that is not a positive number is rejected (#117).
- The synchronizer drops procedure-level attribute lines (for example
  `Attribute Worksheet_Change.VB_Description = ...`) from document modules
  before pasting their code, where they would be invalid VBA; the code around
  them is kept (#118).
- The issue-metadata workflow no longer fails on issue events that GitHub
  raises for pull requests (#41): it skips them before reading the item, which
  its token, without pull-request access, cannot do. Failed API calls now log
  the HTTP status and response, with token-like strings redacted, instead of a
  bare `CalledProcessError`.
- Label commands now default to the checked-in `.github/label-policy.json`.
- Workbook package checks reject missing core parts and archive corruption;
  these structural checks do not certify Excel compatibility.
- Traffic history labels the repository count `open_issues_and_prs`, including
  pull requests, and preserves values when migrating the old CSV column.
- Traffic alert issues are now created in a milestone, so every issue
  carries one: the milestone named by the repository variable
  `ISSUE_MILESTONE_NUMBER`, which must be open; the run fails otherwise.
- The data-contract reference check applies the cross-row history rule
  (`E11`) before duplicate (`E05`) and currency-change (`E08`) rules and to
  every affected row, as `DATA_CONTRACT.md` orders them; duplicates and
  currency changes no longer hide inconsistent account dates.
- Repository audit corrections (#25): shared VBA comment/continuation lexer,
  exact workbook path and encoding-independent XML reference checks, stronger
  account/market fixture validation, and reconciliation without premature rounding.
- Core-boundary checks now reject `CreateObject` and `GetObject`, including
  variable or concatenated ProgIDs that could previously hide Excel access.
  External COM automation belongs in host adapters; these names are reserved
  in core code even for non-Excel objects. Comments and strings remain allowed
  ([#22](https://github.com/danielep71/K-IRRBB-NMDs/issues/22)).
- Continued comments now recognize VBA whitespace, including tabs, and require
  the continuation underscore immediately before the newline.

### Security

- The synchronizer and the export script need Excel's *Trust access to the VBA
  project object model*; `SECURITY.md` says to enable it only while running
  them. The development-workbook builder still needs neither Excel automation
  nor that setting (#6).
- `.gitignore` now keeps local data, extracts, snapshots and parameter
  exports out of Git (rooted `private/`, `local-data/`, `extracts/`,
  `snapshots/`, `exports/` and `parameter-exports/` folders), ignores binary
  data, statistical-package and archive formats, and covers more key and
  credential files. Synthetic CSV and JSON fixtures stay trackable (#11).

### Documentation

- `docs/licenses/THIRD_PARTY_NOTICES.txt` lists the files adapted from
  VBA-SACCR-Toolkit, EXCEL-VBA-PROJECT-TEMPLATE and MyTemplate, with each
  source's license where verified and the changes made, and is linked from
  `README.md` and `CONTRIBUTING.md`.
  Workflows, configuration, scripts and tests no longer name other
  repositories or cite issue and pull request history in their comments; the
  Markdown documents record my decisions in the first person and address
  contributors directly. `CONTRIBUTING.md` states both rules (#132).
- Correct the contributor guide's current status: the draft Decay engine exists
  but remains unvalidated under #36; Rates and Stable are not implemented.

- The manual-import component list in `INSTALLATION.md` names every module of
  `tools/vba-projects.json`, production and development, including the seven
  decay and example modules it omitted; `tools/test_devbuild.py` keeps the two
  equal.
- The milestone guides in `docs/Milestones/` follow one schema, the same as
  the GitHub milestone descriptions: Objective, Scope, Exit criteria, Tracked
  issues and dependency, Not in scope and Release evidence. Each keeps its
  rationale, a synthetic acceptance case and its references; the v0.3.0 guide
  names the open decay decisions of #36 and the missing `DecayExample` sheet,
  and the v0.4.0 and v0.5.0 guides name their registered cases.
- Social preview image `docs/assets/IRRBB_social_preview.png`, titled
  K-IRRBB and generated from `IRRBB_social_preview.html` in the layout of
  K-SACCR and K-FRTB-SA, and shown at the top of the root README. Its figures are the
  decay reference case `DEC-E2E-01` (draft, #36) and the REG-1
  core-deposit caps; `docs/assets/README.md` records each source and how to
  regenerate the image.
- Data-contract amendment for point-in-time data (#32), decided on
  2026-10-09: a back-filled `close_date` is accepted but masked in the
  point-in-time view at earlier forecast origins, and closures are assumed
  reported by the month end in which they occur (attested per import, recorded
  in the lineage); only a recorded closure is a
  cash-out, and an account that disappears without one is right-censored
  (`W05`, redefined); a month end with no rows inside the window rejects the
  import (new `E12`); `E10` now also rejects a row dated on its own `close_date`, so a
  closure never sits on a month-end row. `DATA_CONTRACT.md` defines the outcome of every
  transition (observed, verified closure, gap, censored, end of sample) and its
  eligibility and lineage; `MODEL_CONTRACTS.md` updates the stable-model
  cash-out rule and the predictor timing. A hand-built fixture and leakage
  tests (`DATA-PIT-01`, `DATA-PIT-02`) check it.
- Translate all milestone guides into US English and establish US English
  as the repository language standard. Standardize surrounding documentation
  while preserving glyphs, navigation, technical scope, and official source
  titles.

- Detailed milestone guides in `docs/Milestones/`, one per milestone
  from v0.1.0 through v1.0.0, with a navigation index, purpose, deliverables,
  dependencies, synthetic examples and completion evidence. Implementation
  tracking gaps are explicit; the guides do not change model contracts or
  assert that planned functions are delivered.
- The v0.1.0 guide records the milestone as closed on 2026-10-09: each
  deliverable links its issue and the document where it lives, the
  distribution review (#42) and MPL-2.0 relicensing (#49) are included, and
  the completion criteria are marked met. The status sections of README,
  `CONTRIBUTING.md`, `INSTALLATION.md` and `RELEASING.md` now name v0.2.0 as
  the active milestone.
- `docs/FOUNDATION_CLOSEOUT.md` records the v0.1.0 closeout: every checklist
  item is marked met with the evidence it rests on.

- Root documents (`CHANGELOG.md`, `CODE_OF_CONDUCT.md`, `CONTRIBUTING.md`,
  `INSTALLATION.md`, `RELEASING.md`, `SECURITY.md`) expanded to my
  documentation standard, with headers, badges and complete
  workflow, evidence, security and release sections.
- `CONTRIBUTING.md` gains a "Data, confidentiality and provenance" section:
  what counts as synthetic, private visibility is not permission, no
  third-party code or restricted manuals without rights, provenance in every
  PR, and what to do if sensitive material is committed. `docs/GOVERNANCE.md`
  adds a licensing and distribution gate with a full-history scan procedure
  (#11).
- The architecture, validation plan and feature issue form refer to a
  generic downstream ALM platform instead of a named vendor product (#11).
- Model contracts accepted (#9) in `docs/methodology/MODEL_CONTRACTS.md`:
  rate pass-through (unit roots, Engle-Granger long run and cointegration,
  symmetric and asymmetric ECM, constraints rejected not clipped), stable
  amount (one-month survival share, balance-weighted fractional logit,
  predictors known at the forecast date, convergence and separation
  failures) and decay (log changes, prudential drift, minimum probable
  amount, conserved profile, mean life), with segment granularity,
  out-of-sample freeze, confidence bands, scenarios, and the regulatory
  overlay kept separate from the fit. `SOURCES.md` registers the cited
  sources: BCBS d368 (core-deposit categories and caps) and Delegated
  Regulation (EU) 2024/856 (outlier-test scenarios and floor) verified,
  EBA/GL/2022/14 not yet; `TEST_CASES.md` registers the numerical
  reference cases with provenance rules and tolerances.
- Data contract accepted (#4) in `docs/methodology/DATA_CONTRACT.md`: account
  panel and market-rate field dictionaries with types, units and null policy;
  segment codes `RET_TX`, `RET_NTX`, `WHS_NFC`; reconciliation; all-or-nothing
  import with error codes `E01`–`E11` and warnings `W01`–`W06`; gaps,
  openings, closures, migrations, outliers and look-ahead; CSV streaming for
  panels beyond worksheet limits; and lineage. Synthetic fixtures with
  hand-computed totals and warnings are checked by `tools/test_data_contract.py`.
- Supported hosts decided (#3): Windows 64-bit Microsoft 365 or Excel 2016+
  is the supported target, 32-bit is best effort (kept compiling, not
  certified), Mac, web and Excel 2013 or earlier are not supported. Only the
  four default VBA references are allowed, and the workbook has no runtime
  dependency beyond Excel; Python or R may produce reference values offline.
  `INSTALLATION.md` separates these commitments from test evidence, and
  `docs/EXCEL_EVIDENCE.md` defines the clean-build, compile, smoke-run,
  failure-path and cleanup stages, their outcomes and per-bitness rules.
- `docs/REPOSITORY_STRUCTURE.md` rewritten as the architecture, which I
  accepted: profile rationale, the intended Excel workflow, layout,
  dependency direction, public API boundary, canonical units at the facade,
  immutable versioned parameter sets and intentional non-goals (#2).
- The v0.1.0-to-v1.0.0 roadmap I approved: workbook/import/harness work moves
  to v0.2.0, with all Excel obligations retained; model-contract amendments
  are split into milestone-specific prerequisites (#26, #32-#37). Independent
  benchmarks accompany each model, and Stable-Decay composition is explicit.
- Accepted host/data/model records now have direct evidence links; unresolved
  model interpretation and reproducibility requirements are tracked in #26.
- `RELEASING.md` and `CONTRIBUTING.md` follow the rulesets copied from
  VBA-SACCR-Toolkit: task PRs are squash-merged into the release branch, the
  release branch is merged into `main` with a merge commit from its own PR, and
  the release branch, which cannot be deleted, is then levelled with `main`
  through a merge-commit PR (#170). VBA may be merged into the release branch before Excel
  evidence exists; nothing is released without it.
- Documentation prepared for public visibility (#42): `SECURITY.md` routes
  reports through GitHub private vulnerability reporting, with email as the
  fallback; `docs/TRAFFIC.md` states that the traffic branch is public and
  holds aggregate counts only; README, `CONTRIBUTING.md`, `GOVERNANCE.md` and
  the foundation closeout no longer assume a private repository, and public
  visibility is recorded as my decision, separate from any release.

### Validation

- Decay model as merged into `release/v0.3.0`: `python tools/check.py --ci`
  passes; LibreOffice 24.2 headless proxy run of the development workbook, see
  the merge commit messages for counts. **No Excel compile or run has been recorded**;
  LibreOffice is not a supported host and this is not Excel evidence.

### Known limitations

- The decay model is a draft under review (#36). Nothing yet builds its
  aggregate series from an accepted import; the caller supplies the series.
  The rate pass-through and stable-amount models are not built.
- Decay choices not fixed by the contract await my decision in #36: the
  treatment of censored disappearances (`W05`) in the aggregate series, the
  break-date convention, the median path, the diagnostic settings and the
  backtest horizons 1 to 12.
- No Excel compile or run has been recorded for the VBA source; the
  LibreOffice proxy runs are not Excel evidence.
- No automated check compiles VBA, runs Excel or validates a model; passing CI
  is not evidence of correct IRRBB calculations.
- The methodology is a roadmap, not a validated specification. Of its
  regulatory references, only `REG-1` and `REG-3` are verified.

---

[Unreleased]: https://github.com/danielep71/K-IRRBB-NMDs/compare/main...release/v0.2.0
