<div align="center">

# 🤝 Contributing to IRRBB Toolkit

### Interest Rate Risk in the Banking Book — behavioral models in Excel/VBA

[![Conduct](https://img.shields.io/badge/Conduct-Required-6f42c1?style=flat-square)](CODE_OF_CONDUCT.md)
[![Security](https://img.shields.io/badge/Security-Private_reporting-d73a49?style=flat-square)](SECURITY.md)
[![Evidence](https://img.shields.io/badge/Evidence-Exact_source-0969da?style=flat-square)](#validation-and-evidence)
[![Workflow](https://img.shields.io/badge/Workflow-PR_into_release-217346?style=flat-square)](#development-workflow)

<br>

**Focused scope · Reviewable source · Independent references · Honest limitations**

</div>

---

This document is authoritative for the **contribution and review workflow**.
Branch, issue and review policy is owned by
[`docs/GOVERNANCE.md`](docs/GOVERNANCE.md); developer setup by
[`INSTALLATION.md`](INSTALLATION.md); vulnerability handling by
[`SECURITY.md`](SECURITY.md); publication by [`RELEASING.md`](RELEASING.md).

Participation is governed by the [Code of Conduct](CODE_OF_CONDUCT.md).
Suspected vulnerabilities must never be disclosed in an issue or pull request.

> [!NOTE]
> Milestone **v0.2.0 – Workbook, data and Excel tests** is in progress;
> **v0.1.0 – Repository foundation** closed on 2026-10-09. No VBA source,
> workbook, regression harness or model engine exists yet. The model families
> and their obligations are described in
> [`docs/methodology/`](docs/methodology/README.md); the source layout is set in
> [`docs/REPOSITORY_STRUCTURE.md`](docs/REPOSITORY_STRUCTURE.md). Sections below
> that refer to VBA describe the rules those changes must follow once they exist.

## 🌐 Repository language

Use **US English** for all repository-authored content: documentation, code
comments, diagnostics and user-facing messages, issue and pull request titles
and bodies, review comments, milestone descriptions, commit messages, and
release notes. Use US spellings such as `behavior`, `modeling`, `license`,
`analyze`, and `color`. Write new identifiers and descriptive filenames in
English, while preserving existing API and file-format compatibility.

Keep proper names, official publication titles, quotations, and external API
identifiers accurate to their sources. Explain quoted non-English material
in US English. Preserve glyphs and mathematical notation where useful; the
language standard does not require ASCII-only text. Do not rewrite published
Git history to apply an editorial correction.

<a id="ways-to-contribute"></a>

## 🌱 Ways to contribute

| Contribution | First action |
| --- | --- |
| 🐛 Reproducible defect | Open an issue with exact commit, environment, expected and observed behavior, and synthetic evidence. |
| ✨ Model or feature change | Open an issue defining the methodological basis, observable contract, compatibility, non-goals and validation. |
| 📐 Estimator or calibration change | State the reference implementation, tolerances, convergence criteria and failure behavior before writing code. |
| 🧪 Test or evidence improvement | Explain provenance, independence, coverage and what failure the test detects. |
| 📖 Documentation | Identify the authoritative document being corrected. |
| ⚙️ Tooling or governance | Explain failure behavior, portability, trust boundary and maintenance cost. |
| 🔐 Security concern | Follow [`SECURITY.md`](SECURITY.md) privately. |

Open issues with the forms (bug, feature, documentation). Every issue is
assigned to `danielep71`, carries **exactly one** priority label (`P1`, `P2` or
`P3`) and has a milestone; the forms set the first two, and triage sets the
milestone. See [`docs/GOVERNANCE.md`](docs/GOVERNANCE.md#issue-metadata). Pull
requests follow [the PR template](.github/PULL_REQUEST_TEMPLATE.md).

<a id="development-workflow"></a>

## 🌿 Development workflow

1. Start from the active release branch, currently `release/v0.2.0`. Never assume
   GitHub's default branch is the right base.
2. Create one focused task branch named in the repository convention:
   `fix/<issue>-<slug>`, `docs/<slug>` or `test/<slug>`.
3. Reproduce the current behavior before changing it.
4. Define the observable contract, affected callers, compatibility impact and
   validation plan.
5. Make the smallest coherent change; avoid unrelated formatting, generated
   output or opportunistic refactoring.
6. Run `python tools/check.py` before every push. All gates must pass.
7. Update the affected documentation and add a user-facing entry under
   `## [Unreleased]` in [`CHANGELOG.md`](CHANGELOG.md).
8. Review the complete diff, then open a pull request **against the release
   branch** with evidence and explicit limitations.
9. Merge only after review, with the **Repository integrity** check green and
   every review finding resolved. The release branch accepts changes only
   through pull requests; a merge commit keeps each commit as authored, and a
   squash must carry a commit message that follows the rules below.

Nothing is committed directly to the release branch or to `main`. Integration
of the release branch into `main` happens only on the owner's request; see
[`RELEASING.md`](RELEASING.md). Pushed history on a shared branch is never
rewritten: no rebase, amend or force-push. Bring `main` in with a merge commit.

### Commit messages

Use a conventional subject line, for example `fix: …`, `docs(plan): …`,
`test: …` or `fix(ci): …`, then a body that states:

- **what** changed;
- **why** it changed; and
- **what was validated** at that commit, separating static checks from Excel
  runs, and stating plainly what was **not** validated.

Reference the issue when one exists. Do not place credentials, private links,
attribution boilerplate or unverifiable test claims in commit messages.

## 📦 Source-change discipline

The repository-control files define how source is stored:

- exported VBA (`.bas`, `.cls`, `.frm`) is the reviewable source of truth. It
  is stored with LF in Git and checked out with CRLF;
- component filenames match their `VB_Name`, and every module declares
  `Option Explicit`;
- `.frm`/`.frx` pairs stay together, and `.frx` stays binary;
- workbooks and add-ins are generated artifacts, ignored by Git unless an exact
  path is re-included; a workbook is never the only record of a source change;
  and
- test data and examples are synthetic.

`python tools/check.py` enforces storage, naming, placement and
`Option Explicit`, plus the VBA jump, conditional-compilation and public-API
checks. Component placement, the dependency direction (`src/workbook/` →
`src/modules/` → `src/core/`) and the public-API boundary are defined in
[`docs/REPOSITORY_STRUCTURE.md`](docs/REPOSITORY_STRUCTURE.md); supported
declarations are listed in [`docs/PUBLIC_API.txt`](docs/PUBLIC_API.txt). Naming,
contracts, errors and Excel-state handling follow
[`docs/VBA_HOUSE_STYLE.md`](docs/VBA_HOUSE_STYLE.md).

Do not weaken a calculation, numerical or packaging gate merely because the
generic repository gate passes.

<a id="data-confidentiality-and-provenance"></a>

## 🔐 Data, confidentiality and provenance

**Synthetic data only.** Every dataset, fixture, example and screenshot in Git,
in an issue or in a pull request must be synthetic. Synthetic means generated
from a documented procedure, with its generator or seed committed or described,
and not derived from real records: rescaled, shifted, anonymized or sampled
real balances, rates or account histories are **not** synthetic.

**Visibility is not permission.** The repository is public, and its full
history, branches and pull-request refs are readable by anyone. A private fork
or branch does not authorize storing client, employer or personal data, and
does not waive confidentiality, GDPR or contractual restrictions. Treat every
commit as published, because history is hard to erase.

**Third-party material needs explicit rights.** Do not commit, attach or
paraphrase at length:

- vendor or third-party source code, workbooks or add-ins;
- restricted manuals, model documentation, training material or methodology
  papers from a vendor, bank or consultancy;
- internal bank models, parameters, calibration results or validation reports;
  or
- licensed data that cannot be redistributed.

Cite public sources (regulation, supervisory guidelines, academic papers) by
reference and implement from that citation. Name a commercial product only
where an integration contract requires it, never its internals.

**State provenance.** A pull request that adds data, a fixture, an expected
value or adapted code states where it came from: generator and seed, the public
source cited, or the license of the adapted material.

**If sensitive material reaches Git**, stop pushing to that branch and report
it privately as described in [`SECURITY.md`](SECURITY.md#data-and-secrets).
Rotate any exposed secret first. Removing it from history is an owner decision;
a follow-up commit that deletes the file does not remove it.

<a id="compatibility-and-model-contracts"></a>

## 🔄 Compatibility and model contracts

A change to documented procedures, functions, classes, enums, parameters,
defaults, return values, errors, side effects, input or output layouts,
parameter-export formats or supported platforms is a contract change. So is any
change that alters an estimated coefficient, stable share, decay profile, mean
life or backtest statistic for the same inputs and settings.

Such a contribution must identify callers and migration impact, update
regression coverage, update the user documentation and state the release
impact. A methodology change must also cite the source it implements and record
it in [`docs/methodology/`](docs/methodology/README.md).

Statistical estimators carry extra obligations:

- **independent numerical reference cases** with stated tolerances;
- **documented failure behavior** for non-convergence, separation, singular
  designs and insufficient data;
- **convergence and reproducibility** settings recorded with the result; and
- **no look-ahead**: backtesting must never use information unavailable at the
  forecast date.

Excel and Windows state belongs to the caller or host unless a component
explicitly owns it. Capture before mutation, restore only state successfully
changed and still owned, and never let cleanup conceal the original failure.

<a id="validation-and-evidence"></a>

## 🧪 Validation and evidence

Validation must be reproducible from the exact source under review. Record:

```text
Source
------
Commit:
Files or components changed:

Environment
-----------
Excel / Office build:
Office bitness:
Operating system:
Locale / date system:

Checks
------
python tools/check.py:
Compile (Debug → Compile VBAProject):
Regression harness:
Reference cases and tolerances:
Specific scenario verified:
Cleanup:

Limitations
-----------
Not run or unverified:
Follow-up:
```

Use only the applicable fields, but never omit a material limitation.

- **Static checks are not Excel evidence.** `tools/check.py` and the hosted
  workflow inspect files. They do not compile VBA, run Excel or validate any
  model. Evidence bound to an exact commit follows
  [`docs/EXCEL_EVIDENCE.md`](docs/EXCEL_EVIDENCE.md).
- **VBA changes are not merged until verified in Excel:** compile, run the
  harness, and exercise the specific scenario the change addresses. State
  plainly what was and was not run in Excel.
- **A skipped check is not a pass.**
- **Reference values must be independent** of the implementation under test,
  for example a published worked example or an independent statistical
  package. A model must never validate itself.

<a id="pull-requests"></a>

## 🚀 Pull requests

A pull request answers five questions:

1. What problem does this solve?
2. What observable contract changes?
3. What remains compatible?
4. How was the exact source validated?
5. What remains unverified?

### Review checklist

- [ ] The PR targets the active release branch, and the related issue is linked.
- [ ] Scope is focused; model, API and compatibility impact are assessed.
- [ ] Exported VBA and required binary companions are synchronized.
- [ ] `python tools/check.py` passes, and the **Repository integrity** check is green.
- [ ] Excel compile, harness and scenario results are recorded for VBA changes.
- [ ] Estimator changes carry independent references, tolerances and failure cases.
- [ ] Documentation and the `[Unreleased]` changelog entry are updated.
- [ ] No confidential, real-portfolio, accidental binary or generated material is added.
- [ ] New data, fixtures and adapted code are synthetic or licensed, with provenance stated.
- [ ] Unverified environments and skipped checks are stated plainly.
- [ ] The final diff contains no unrelated formatting or local artifacts.

The definition of done is in
[`docs/GOVERNANCE.md`](docs/GOVERNANCE.md#definition-of-done). Discussion stays
technical and respectful under the [Code of Conduct](CODE_OF_CONDUCT.md).

## 📚 Where detailed rules live

| Need | Document |
| --- | --- |
| Source layout and public API | [`docs/REPOSITORY_STRUCTURE.md`](docs/REPOSITORY_STRUCTURE.md) |
| VBA naming, contracts and errors | [`docs/VBA_HOUSE_STYLE.md`](docs/VBA_HOUSE_STYLE.md) |
| Model families and obligations | [`docs/methodology/`](docs/methodology/README.md) |
| Backtesting and model-risk controls | [`docs/methodology/VALIDATION_PLAN.md`](docs/methodology/VALIDATION_PLAN.md) |
| Commit-bound Excel evidence | [`docs/EXCEL_EVIDENCE.md`](docs/EXCEL_EVIDENCE.md) |
| Branches, issues, definition of done | [`docs/GOVERNANCE.md`](docs/GOVERNANCE.md) |
| Developer setup | [`INSTALLATION.md`](INSTALLATION.md) |
| Static checks and CI | [`tools/README.md`](tools/README.md) |
| Issue labels | [`docs/LABELS.md`](docs/LABELS.md) |
| Vulnerability reporting | [`SECURITY.md`](SECURITY.md) |
| Release procedure | [`RELEASING.md`](RELEASING.md) |
| Change history | [`CHANGELOG.md`](CHANGELOG.md) |

## 📄 Licensing and maintainer

This project is distributed under the [Mozilla Public License 2.0](LICENSE)
(MPL-2.0); contributions are accepted under the same license. Contributors must
have the right to submit every part of a contribution, including code, tests,
data, images and generated material. Third-party methodologies are cited, not
copied. The license is reviewed before any external distribution; see
[`docs/GOVERNANCE.md`](docs/GOVERNANCE.md#licensing-and-distribution).

Maintained by **Daniele Penza**.

---

**Contribution principle:** make the contract explicit, keep the diff focused,
and leave evidence another person can reproduce.
