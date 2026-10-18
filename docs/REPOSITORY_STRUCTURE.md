<div align="center">

# 🗂️ Repository Structure and Application Profile

### Where every artifact belongs, and how the model core stays independent of Excel

[![Profile](https://img.shields.io/badge/Profile-application-217346?style=flat-square)](#profile-decision)
[![Model](https://img.shields.io/badge/Model-source--first-0969da?style=flat-square)](#repository-layout)
[![Boundary](https://img.shields.io/badge/Boundary-facade_over_core-6f42c1?style=flat-square)](#public-api-boundary)
[![Status](https://img.shields.io/badge/Status-Accepted_(%232)-217346?style=flat-square)](#profile-decision)

<br>

**One home per responsibility · The core never touches Excel · Units explicit at every boundary**

</div>

---

This document is authoritative for **the project profile, where every durable
artifact belongs, the boundary between the model core, the public facade and the
Excel host, and the parameter and units contract at that boundary**. Import order
and host support are owned by [`INSTALLATION.md`](../INSTALLATION.md); how VBA is
written by [`VBA_HOUSE_STYLE.md`](VBA_HOUSE_STYLE.md); model definitions by
[`methodology/`](methodology/README.md).

> [!IMPORTANT]
> This structure was **accepted by the owner on 2026-10-08 in issue #2** as the
> architecture decision: the `application` profile, the intended workflow, the
> canonical units and the parameter-set fields. Rules marked *enforced* below
> are checked by `python tools/check.py`; the rest are enforced by review.

<a id="profile-decision"></a>

## 🧭 Profile decision

**The IRRBB Toolkit uses the `application` profile.** Decided by the owner on
2026-10-08 in issue #2.

| Question | Answer for the IRRBB Toolkit |
| --- | --- |
| Who calls it? | A model developer or validator working in the IRRBB workbook: loading a synthetic or permitted dataset, calibrating the behavioral models, backtesting them, running scenarios and exporting parameters. |
| What does it own? | The workbook: input, parameter and result sheets, the model engines, run control, and their packaging. |
| Lifecycle | The workbook is the deliverable; it is built from the exported source in this repository, never edited as the source of truth. |
| Supported environments | Excel for Windows, 64-bit target and 32-bit best effort; see [supported hosts](../INSTALLATION.md#supported-hosts) (decided in issue #3). |

**Why not `library`:** the user needs a guided workflow with persistent inputs,
parameters, diagnostics and stale-result protection. Those belong to a
workbook; a function library alone would push them onto every user.

**Why not `add-in`:** the workbook holds per-analysis state (data window,
segments, calibrated parameter sets). An add-in would separate code from that
state and complicate reproducibility.

**What stays library-like:** the model core is host-independent. It never
touches workbook objects, so every estimator can be tested from arrays without a
workbook and reused if a library deliverable is ever wanted.

`.github/label-policy.json` already selects the `application` profile for the
label catalog; this document makes it the architecture decision as well.

<a id="excel-workflow"></a>

## 🔄 Intended Excel workflow

The architecture is checked against this end-to-end workflow. Sheet names are
the proposal in issue #5; data fields come from issue #4.

| Step | User action | Workbook layer (`src/workbook/`) | Facade and core |
| ---: | --- | --- | --- |
| 1 | Configure the run: segments, currency, data window, model versions | `Config`, `Segments` sheets | Validates and freezes run settings |
| 2 | Load a synthetic or permitted dataset | Reads the input range into arrays | Data validation, segmentation, reconciliation (`DataSummary`) |
| 3 | Calibrate the rate pass-through model | Writes parameters and diagnostics | ECM / OLS estimation and specification tests |
| 4 | Calibrate the stable-amount model | Writes parameters and diagnostics | Cash-out measurement and logit estimation |
| 5 | Calibrate the decay model | Writes the profile and diagnostics | Log-change dynamics, quantiles, mean life |
| 6 | Backtest out of sample | `Backtesting` sheet | Frozen-coefficient forecasts, no look-ahead |
| 7 | Run scenarios | `Scenarios` sheet | Rate-dependent responses on a frozen parameter set |
| 8 | Review checks and results | `Results`, `Checks` sheets | Status of every run; stale results withdrawn |
| 9 | Export a parameter set | Export adapter | Reads a versioned parameter set; no recalculation |

Every step can fail without leaving earlier results looking current: a failed
or invalidated run marks its outputs stale ([`VBA_HOUSE_STYLE.md`](VBA_HOUSE_STYLE.md)).

<a id="repository-layout"></a>

## 📁 Repository layout

```text
src/        production VBA source: the only input to the workbook
tests/      regression modules, synthetic fixtures and expected results
examples/   runnable examples of the supported API
docs/       contracts, architecture and methodology
tools/      static checks and evidence tooling; later, build tooling
.github/    workflows, label catalog and scripts
```

| Directory | Owns | Must not own |
| --- | --- | --- |
| `src/` | Production components that go into the workbook, and the workbook template | Tests, examples, built workbooks |
| `tests/` | Test modules, synthetic fixtures, reviewed expected values | Production entry points, run output, real data |
| `examples/` | Examples that use only the public API | Tests, real data |
| `docs/` | Contracts, architecture and [methodology](methodology/README.md) | Copies of root documents |
| `tools/` | Deterministic checks, build and evidence scripts | Model or calculation logic |

Each directory explains itself in a `README.md` until real material arrives.
Subdirectories are created with their first real file, never empty.

### Production source

| Location | Holds | Visibility |
| --- | --- | --- |
| `src/core/` | Model engines and numerics: estimation, transformations, diagnostics, profiles, data validation | In-project only. Every standard module declares `Option Private Module` (*enforced*). |
| `src/modules/` | Public facade: supported `IRRBB_*` entry points, constants and enums | Supported API, listed in [`PUBLIC_API.txt`](PUBLIC_API.txt) (*enforced*) |
| `src/classes/` | Class modules, e.g. parameter-set or run-state objects | Status stated in each class header |
| `src/workbook/` | Document modules (`ThisWorkbook`, sheet modules), sheet adapters, export adapters, the macro-free template | Host glue only; no model logic |
| `src/forms/` | UserForms, `.frm` beside its `.frx`, only if a form is needed | — |

### Verification and examples

| Location | Holds |
| --- | --- |
| `tests/modules/` | Regression modules and the harness entry point (issue #8) |
| `tests/fixtures/` | Synthetic panels, rate series and scenario definitions |
| `tests/expected/` | Expected values with their independent provenance and tolerances (issue #9) |
| `examples/modules/` | Example modules calling the public facade |

<a id="dependency-direction"></a>

## 🔀 Dependency direction

```text
Excel sheets / buttons / workbook events
          │
          ▼
src/workbook  ──▶  src/modules (facade)  ──▶  src/core
                         │                       ▲
src/classes  ◀───────────┘                       │
tests/       ────────────────────────────────────┘ (may test core directly)
examples/    ──▶  src/modules only
```

- **`src/core` depends on nothing outside `src/core`.** It takes arrays, scalars
  and dates and returns results or raises errors. It never reads or writes a
  workbook, sheet or range, never shows UI and never reads `Application` state
  (*enforced*: `tools/check_source.py` rejects Excel object-model and UI
  identifiers and COM activation/attachment in `src/core/` code; comments and
  strings are ignored).
- **`src/modules` validates at the public boundary** (units, ranges, array
  shape, ordering, missing values) and delegates to the core. It holds no model
  formulas of its own.
- **`src/workbook` adapts**: it reads sheets into arrays, calls the facade and
  writes results. It does not compute, and it never silently transforms a
  statistical or regulatory assumption.
- **Export adapters** for downstream ALM platforms live in `src/workbook/` and
  read a versioned parameter set; they are implemented only after an export
  contract is agreed.
- `tests` may call core procedures directly; `examples` may not.

Because VBA identifiers are case-insensitive, the reserved host names
(`Range`, `Cells`, `Application`, `MsgBox` and the others in
`tools/check_source.py`) cannot be used even as variable names in `src/core`.
`CreateObject` and `GetObject` are also reserved: core must not activate or
attach to external COM objects, including non-Excel objects or objects selected
by variable or concatenated ProgIDs. Such automation belongs in host adapters.
This is a static reserved-name check, not a proof of transitive independence.

<a id="public-api-boundary"></a>

## 🧩 Public API boundary

A VBA `Public` declaration is not automatically supported API.

- **Supported:** declarations in `src/modules/` listed in
  [`PUBLIC_API.txt`](PUBLIC_API.txt). Changing one is a contract change under
  [`CONTRIBUTING.md`](../CONTRIBUTING.md#compatibility-and-model-contracts).
- **In-project only:** everything in `src/core/`, kept off the external surface
  by `Option Private Module`, and anything in `src/modules/` not listed in the
  manifest.
- **Never supported:** test and example modules. They are not part of the
  production workbook.

The manifest lists the data-import facade `IRRBB_Import`, each declaration with its normalized signature.

<a id="parameter-and-units-boundary"></a>

## 📐 Parameter and units boundary

Units and versions are part of every public contract. The exact data
dictionary is fixed in issue #4; the boundary rules below hold regardless.

### Canonical units at the facade

| Quantity | Canonical form inside the facade and core | Where conversion happens |
| --- | --- | --- |
| Interest rates and spreads | Decimal per annum (`0.0125` = 1.25 %) | Workbook adapter, when a sheet shows percent or basis points |
| Rate shocks | Decimal, signed: positive = rates up | Workbook adapter |
| Dates | VBA `Date` values, observation dates at month end | Workbook adapter; the core parses only the ISO dates of the data files (`CORE_Parse`), never locale-dependent text |
| Balances | Currency units as `Double`, with the ISO currency code carried alongside | No FX conversion in the core |
| Shares and probabilities | Decimal in `[0, 1]` | Workbook adapter, when shown as percent |
| Horizons and maturities | Whole months | Workbook adapter, for other display units |
| Arrays | One observation per element, in ascending date order, with stated lower bound | Facade validates; core assumes |

A value that crosses the facade in any other unit is a defect, not a
convention. Conversions are explicit, named and tested.

### Versioned parameter sets

A calibration produces an **immutable parameter set**. Results, scenarios,
backtests and exports refer to it by ID; nothing recalibrates silently.

| Field | Content |
| --- | --- |
| Identity | Parameter-set ID, creation timestamp |
| Model | Model family (rates, stable, decay) and model version from [`methodology/`](methodology/README.md) |
| Scope | Segment, currency |
| Data | Data window (first and last observation date), source-data fingerprint |
| Code | Toolkit commit SHA that produced it |
| Values | Each parameter with its unit, plus the diagnostics summary |
| Status | `draft` or `superseded`; acceptance for use happens outside the toolkit |

Changing a model version, a data window or the code produces a new parameter
set, never an edit of an existing one. Export adapters read a parameter set as
stored; they do not recalculate or round beyond the agreed export format.

<a id="non-goals"></a>

## 🚫 Intentional non-goals

The toolkit deliberately does **not**:

- compute full IRRBB measures (EVE, NII, the supervisory outlier test) or
  repricing gaps; it produces behavioral assumptions that such systems consume;
- approve models or parameter sets: acceptance, regulatory compliance and
  model-risk sign-off happen outside the toolkit;
- store or ship real customer, bank or vendor data, at any repository
  visibility ([`CONTRIBUTING.md`](../CONTRIBUTING.md#data-confidentiality-and-provenance));
- depend at runtime on Python, R, external services or non-default Excel
  references; independent reference calculations may use them outside the
  workbook (issue #3);
- push data to a downstream ALM platform automatically: export produces a file
  for controlled hand-over, once the export contract exists;
- ship as an add-in or support Excel for Mac or the web.

<a id="placement-rules"></a>

## 🚦 Placement rules

1. VBA components (`.bas`, `.cls`, `.frm`, `.frx`) live only in
   `src/core/`, `src/modules/`, `src/classes/`, `src/workbook/`, `src/forms/`,
   `tests/` or `examples/`. *(enforced)*
2. Standard modules carry their role prefix: `CORE_` in `src/core/`, `IRRBB_` in
   `src/modules/`, `TEST_` in `tests/`. *(enforced)*
3. Every standard module in `src/core/` declares `Option Private Module`.
   *(enforced)*
4. Code in `src/core/` uses no Excel object-model or UI identifier. *(enforced)*
5. Test and example modules never go into the production import set.
6. Fixtures are synthetic; never commit real balances, rates, accounts or
   segments.
7. Workbooks are built from source and never committed, except at an exact
   path re-included in `.gitignore`. The one such path is
   `src/workbook/IRRBB_Template.xlsx`, the macro-free template of issue #5,
   generated by `tools/workbook/build_template.py` and laid out in
   [`WORKBOOK_CONTRACT.md`](WORKBOOK_CONTRACT.md); it may contain no VBA project
   and no document properties. *(enforced)*
8. A new location becomes contractual only when this document, the directory
   README, `INSTALLATION.md` and the checks are updated together.

## 📚 Directory guides

- [`../src/README.md`](../src/README.md)
- [`../tests/README.md`](../tests/README.md)
- [`../examples/README.md`](../examples/README.md)
- [`methodology/README.md`](methodology/README.md)
- [`../tools/README.md`](../tools/README.md)

---

**Structure principle:** one home per responsibility, the core never touches
Excel, units are explicit at every boundary, and only the facade is supported
API.
