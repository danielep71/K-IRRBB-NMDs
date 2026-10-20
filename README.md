<div align="center">

# 🏦 K-IRRBB

### IRRBB behavioral models in Excel/VBA | Interest Rate Risk in the Banking Book: modeling, calibration and independent testing

[![Excel VBA](https://img.shields.io/badge/Excel_VBA-Windows-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)](INSTALLATION.md)
[![Status](https://img.shields.io/badge/Status-v0.2.0%20in%20progress-6e7781?style=for-the-badge)](#status)
[![Development](https://img.shields.io/badge/Development-release%2Fv0.2.0-6f42c1?style=for-the-badge)](https://github.com/danielep71/K-IRRBB_behavioral_models/tree/release/v0.2.0)
[![License](https://img.shields.io/badge/License-MPL--2.0-2ea44f?style=for-the-badge)](LICENSE)

<br>

**Source-first · Host-independent calculation core · Synthetic data only · Auditable model results**

---

<p align="center">
  <img width="100%"
       alt="K-IRRBB: decay reference case DEC-E2E-01 on a synthetic retail transactional series, its minimum probable amount, the REG-1 core-deposit caps and the out-of-sample backtest"
       src="docs/assets/IRRBB_social_preview.png">
</p>

---

</div>

---

## ✨ Purpose

**K-IRRBB** is an Excel/VBA **application project** intended to develop transparent, independently testable tools for behavioral assumptions in Interest Rate Risk in the Banking Book (IRRBB). It addresses the modeling of non-maturity deposits (NMDs) and, where applicable, demand-side balances, with an emphasis on reproducible calculations, model risk management and ALM parameter export.

The planned scope has three distinct model families:

| Model | Question answered | Intended techniques |
| --- | --- | --- |
| **Interest-rate model** | How do administered customer rates respond to market rates? | Long-run relationship, error-correction dynamics (ECM), asymmetric pass-through, OLS and specification tests |
| **Stable amount model** | What portion of demand balances is expected to persist over a chosen horizon? | Cash-out / survival estimation, binomial-logit GLM, account or cluster predictors, rate-dependent scenarios |
| **Decay model** | How does the stable component amortize over time? | Log balance-change dynamics, volatility, minimum-probable-amount profiles, mean life and maturity constraints |

Planned cross-cutting functionality includes data validation, calibration, historical out-of-sample backtesting, scenario sensitivity, diagnostics, versioned parameter sets and controlled export for downstream ALM platforms.

> [!IMPORTANT]
> The models above are **design objectives**, not implemented or independently validated features. Methodological choices, econometric assumptions, statistical tests and regulatory constraints require explicit documentation and numerical validation before use. In particular, a statistical decay horizon does not by itself establish regulatory compliance of an IRRBB repricing profile.

<a id="status"></a>
## 🧭 Status

**v0.2.0 – Workbook, data and Excel tests (in progress).** The [v0.1.0 foundation](docs/Milestones/v0.1.0.md), closed on 2026-10-09, established source layout, documentation, issue governance, CI and static source checks; v0.2.0 builds the first workbook, data import and Excel test harness. Neither delivers pricing, econometric fitting, a production workbook or certified IRRBB outputs.

- **Decay model (draft, not validated, milestone v0.3.0):** core, public facade `IRRBB_Decay`, harness cases and example, reviewed on `release/v0.3.0` and present here since this branch was levelled with `main`; built to [`MODEL_CONTRACTS.md`](docs/methodology/MODEL_CONTRACTS.md#decay-model) and under review in [#36](https://github.com/danielep71/K-IRRBB_behavioral_models/issues/36); no Excel evidence is recorded yet. Rate and stable-amount models are not built.
- Application profile: an Excel workbook will be the end-user deliverable, assembled from reviewed exported VBA source.
- Production calculations live in `src/core/` without direct Excel object dependencies; at present the data import's parsing, validation, summaries and fingerprint, and the draft decay calculations.
- `src/modules/` exposes the documented public facade (the data import and the draft decay model, listed in [`docs/PUBLIC_API.txt`](docs/PUBLIC_API.txt)); `src/workbook/` contains host glue only.
- `tests/` holds synthetic fixtures, expected results and the VBA regression harness; evidence from a real Excel run must be separately recorded.
- No client data, confidential third-party material, vendor source code or customer-specific models are allowed in Git; the repository and its history are public.

## 📁 Repository navigation

| Resource | Purpose |
| --- | --- |
| [Structure](docs/REPOSITORY_STRUCTURE.md) | Application profile, ownership and dependency boundaries |
| [Milestone guides](docs/Milestones/README.md) | Detailed guides to each milestone: purpose, deliverables, dependencies and completion evidence |
| [Methodology roadmap](docs/methodology/README.md) | Model definitions, reference and calibration obligations |
| [Validation plan](docs/methodology/VALIDATION_PLAN.md) | Backtesting, numerical evidence and model-risk controls |
| [Installation](INSTALLATION.md) | Development prerequisites and local static checks |
| [Contribution guide](CONTRIBUTING.md) | Review requirements, branches, tests and privacy |
| [Governance](docs/GOVERNANCE.md) | Metadata, priority and completion criteria |
| [VBA house style](docs/VBA_HOUSE_STYLE.md) | Module contracts, error handling and export discipline |
| [Excel evidence](docs/EXCEL_EVIDENCE.md) | Real-host evidence, independent of static checks |
| [Tooling](tools/README.md) | Repository checks and CI limits |
| [Labels](docs/LABELS.md) | Issue-label catalog and priority rules |
| [Releasing](RELEASING.md) | Integration into `main` and the release sequence |
| [Security](SECURITY.md) | Vulnerability and information-handling policy |
| [Code of conduct](CODE_OF_CONDUCT.md) | Participant behavior and conduct reporting |
| [Changelog](CHANGELOG.md) | Unreleased work and eventual versions |

## 🛠️ Getting started

```shell
git clone https://github.com/danielep71/K-IRRBB_behavioral_models.git
cd K-IRRBB_behavioral_models
git switch release/v0.2.0
python tools/check.py
```

Git and Python 3.10+ are needed for the static checks. A Windows 64-bit Excel host is needed for workbook import, compile and the regression harness. **Passing Python checks does not imply VBA compilation or model validation.**

## 🔒 Data, confidentiality and provenance

This project is **generic and independent**. Only synthetic or redistributable reference datasets and expressly permitted source material may be committed. Do not copy third-party proprietary methodologies verbatim, internal bank models, customer balances, reports, restricted vendor manuals or training material, or tools into this repository. No visibility setting or private fork is permission to store any of these. The Mozilla Public License 2.0 covers original repository content; it does not grant rights over third-party intellectual property, and it is reviewed before any external distribution.

Contributor rules are in [Data, confidentiality and provenance](CONTRIBUTING.md#data-confidentiality-and-provenance); the distribution gate is in [Governance](docs/GOVERNANCE.md#licensing-and-distribution).

## 👤 Maintainer

**Daniele Penza** — [@danielep71](https://github.com/danielep71)

## 📄 License

[Mozilla Public License 2.0](LICENSE) (MPL-2.0), copyright 2026 Daniele Penza.
