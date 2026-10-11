# 🗺️ Milestones — from repository foundation to v1.0.0

These guides explain **what each milestone must deliver, why it matters,
and what evidence is required to consider it complete**. The titles match
the GitHub milestones. The reference roadmap was approved on October 9,
2026, and is recorded in the [foundation closeout](../FOUNDATION_CLOSEOUT.md).

## 🧭 Roadmap and expected outcomes

| Milestone | Question it answers | Main deliverable |
| --- | --- | --- |
| [🏗️ v0.1.0](v0.1.0.md) | What rules will guide development? | Architecture, baseline contracts, checks, and governance |
| [📊 v0.2.0](v0.2.0.md) | How do the workbook, data, and tests become operational? | Reproducible Excel application foundation and controlled data import |
| [📉 v0.3.0](v0.3.0.md) | How do we describe aggregate balance runoff? | Decay engine verified against independent references |
| [📈 v0.4.0](v0.4.0.md) | How do customer rates respond to market rates? | Rates model with estimation, diagnostics, and simulation |
| [🏦 v0.5.0](v0.5.0.md) | How much of the existing balance persists over time? | Stable model with consistent cohort aggregation |
| [🔗 v0.6.0](v0.6.0.md) | How do the models work together? | Integrated, reproducible calibration and backtesting |
| [⚖️ v0.7.0](v0.7.0.md) | How do we apply relevant scenarios and constraints? | Consistent scenarios and constraints traceable to their sources |
| [🖥️ v0.8.0](v0.8.0.md) | How does a user operate the application and export results? | User workflow, persistence, and ALM export |
| [🧪 v0.9.0](v0.9.0.md) | Is the complete application ready for release? | Frozen, qualified candidate with complete evidence |
| [🚀 v1.0.0](v1.0.0.md) | What product do we accept and release? | First stable NMD application within its stated limits |

## 📖 How to read these guides

Each file covers purpose, rationale, activities and observable outcomes,
dependencies, a synthetic example, completion criteria, and references.
Activities break down the approved scope into practical work; they do not
certify that a feature is already available. The listed criteria are not
a record of completed tests.

GitHub issues hold current status, ownership, priority, and closure evidence.
These guides do not duplicate percentages or counters that would become
outdated. Each tracking section identifies issues known on October 9, 2026,
and work that still needs a detailed implementation issue.

## 🔗 Dependencies: delivery order and prerequisites

v0.2.0 provides data, the Excel host foundation, and tests for all three
models. Decay precedes Rates and Stable in the delivery sequence, but this
does not create a mathematical dependency of Rates on the Decay formula.
v0.6.0 requires all three models and an explicit composition contract.
v0.7.0 adds coordinated scenarios and applicable constraints; v0.8.0 completes
the user workflow. v0.9.0 consolidates the evidence before v1.0.0 acceptance.

The methodological amendments [#32–#37](../methodology/MODEL_CONTRACTS.md)
must precede the affected implementations. Umbrella issue
[#26](https://github.com/danielep71/VBA-IRRBB-Toolkit/issues/26), assigned
to v0.7.0, does not defer the earlier models' prerequisites to that version.

## 🧪 Common completion requirements

- **Synthetic data and provenance:** examples and evidence committed to the
  repository must be reproducible and contain no confidential real data.
- **Independent references:** expected values must not come from the model
  being tested. Record sources, versions, seeds, and tolerances.
- **Excel evidence:** when VBA is introduced, compilation, tests, and
  scenarios must run on the agreed host and be tied to the exact commit.
  Static CI and LibreOffice do not replace these checks.
- **Errors and results:** an error must not produce a PASS. Previous,
  incomplete, or outdated results must not appear current after inputs change.
  Preserve both the primary error and any cleanup error.
- **Closeout and publication:** closing a milestone does not independently
  authorize main integration, a tag, a release, distribution, or a visibility
  change. Follow [RELEASING.md](../../RELEASING.md) and the owner's decisions.

## 📚 Authoritative decision records

These guides explain the roadmap; they do not amend equations, thresholds,
APIs, or support decisions. If a discrepancy arises, record it in an issue
and update the guide and authoritative contract together through a PR.

| Topic | Reference |
| --- | --- |
| Architecture and source boundaries | [Repository structure](../REPOSITORY_STRUCTURE.md) |
| Data and information timing | [Data contract](../methodology/DATA_CONTRACT.md) |
| Model specifications and prerequisites | [Model contracts](../methodology/MODEL_CONTRACTS.md) |
| Sources and verification status | [Sources](../methodology/SOURCES.md) |
| Numerical cases and validation | [Test cases](../methodology/TEST_CASES.md), [Validation plan](../methodology/VALIDATION_PLAN.md) |
| Excel host evidence | [Excel evidence](../EXCEL_EVIDENCE.md) |
| Workflow and issue closure | [Governance](../GOVERNANCE.md), [Contributing](../../CONTRIBUTING.md) |

The v1.0.0 scope is the behavioral application for non-maturity deposits
(NMDs). A full banking-book EVE/NII engine or a vendor-specific integration
requires separately approved scope. These guides do not set delivery dates.
