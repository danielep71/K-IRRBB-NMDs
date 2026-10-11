<div align="center">

# 🧮 Numerical Test Cases

### Registered reference cases for every model, with independent provenance and tolerances

[![Status](https://img.shields.io/badge/Cases-registered%2C_not_built-d97706?style=flat-square)](#register)
[![Rule](https://img.shields.io/badge/Rule-no_self--validation-6f42c1?style=flat-square)](#provenance)

</div>

---

This register is authoritative for **which numerical cases test each model, where
their expected values come from, and how closely the toolkit must match them**.
The models are specified in [`MODEL_CONTRACTS.md`](MODEL_CONTRACTS.md); the
evidence protocol is in [`VALIDATION_PLAN.md`](VALIDATION_PLAN.md).

> [!NOTE]
> The model cases below are **registered, not built**. The data contract cases
> are *built*: their fixtures and expected files exist and are checked on every
> run of `python tools/check.py`. Cases marked *Built (VBA)* have a VBA
> implementation run by `TEST_Harness.RunTests` in Excel; their literal
> expectations are checked without Excel by `tools/test_vba_harness.py`. No
> Excel run is recorded yet. A case moves to *built* when its fixture and
> expected file are committed with their provenance; the harness registry in
> `tests/modules/TEST_Harness.bas` lists every ID here, in this order.

<a id="provenance"></a>

## 🔎 Provenance rules

- Expected values are produced **outside the toolkit**: by hand, from a
  published worked example, or by an independent package run offline
  ([runtime dependencies](../../INSTALLATION.md#runtime-dependencies)).
- Each expected file records the producing tool and its version, the script or
  steps used, the date, and the person who produced it. A script that produces
  reference values is committed next to the expected file.
- Fixtures are synthetic. A fixture generated from a known process records the
  process, its parameters and its seed.
- An expected value is never copied from toolkit output. If the toolkit and the
  reference disagree, the disagreement is investigated, not the reference
  edited.

<a id="tolerances"></a>

## 📏 Tolerances

| Kind | Rule |
| --- | --- |
| `exact` | Identical after the stated rounding; used for counts, codes and hand-constructed identities |
| `abs ε` | $\lvert x - x_{ref} \rvert \le \varepsilon$ |
| `rel ε` | $\lvert x - x_{ref} \rvert \le \varepsilon \cdot \max(1, \lvert x_{ref} \rvert)$ |

Tolerances reflect numerical method differences only (for example QR versus
normal equations, or iteration stopping rules), never sampling error. A test of
whether an estimator recovers a known data-generating process is a separate,
statistical check, not a tolerance test.

<a id="register"></a>

## 📋 Register

### Harness

| ID | Checks | Reference | Tolerance | Status |
| --- | --- | --- | --- | --- |
| `HARN-OUT-01` | Case outcome rules (`NOT RUN`, `ERROR`, `FAIL` with no assertion, `PASS` only when every assertion passes) and the `abs`, `rel` and `exact` tolerance kinds | Rules on this page; expectations re-derived by [`test_vba_harness.py`](../../tools/test_vba_harness.py) | `exact` | Built (VBA) |

### Workbook

| ID | Checks | Reference | Tolerance | Status |
| --- | --- | --- | --- | --- |
| `WB-LAYOUT-01` | Sheet order, names and CodeNames; every defined name refers to its address; no other names | [`WORKBOOK_CONTRACT.md`](../WORKBOOK_CONTRACT.md); literals checked against [`build_template.py`](../../tools/workbook/build_template.py) by [`test_vba_harness.py`](../../tools/test_vba_harness.py) | `exact` | Built (VBA) |

### Data contract

| ID | Checks | Reference | Tolerance | Status |
| --- | --- | --- | --- | --- |
| `DATA-VALID-01` | Totals per date, segment and currency; warnings `W01`–`W05`; the market-rate fixture is accepted | Hand computation, [`accounts_summary.json`](../../tests/expected/data_contract/accounts_summary.json); in VBA through the generated `TEST_Fixtures` module | `abs 0.005` balances; `exact` counts and codes | Built (Python, VBA) |
| `DATA-INVALID-01` | First error per line, `E02`–`E11`; in VBA also `E01` (header, field count), cross-row precedence, case-sensitive keys, market-rate rules and `E12` for a dropped month end | Hand construction, [`accounts_invalid_findings.json`](../../tests/expected/data_contract/accounts_invalid_findings.json); VBA literals re-derived by [`test_vba_harness.py`](../../tools/test_vba_harness.py) | `exact` | Built (Python, VBA) |
| `DATA-PIT-01` | Transition outcome per account month (observed, verified closure, gap, censored, end of sample), `W05` censoring, `close_date` visible at each origin | Hand assignment, [`availability_outcomes.json`](../../tests/expected/data_contract/availability_outcomes.json) | `exact` | Built |
| `DATA-PIT-02` | Leakage: altering, adding or removing information learned after an origin leaves every predictor at that origin unchanged; a missing month end gives `E12` | Property test on both panels at every origin, with an unmasked negative control | `exact` | Built |
| `DATA-PARSE-01` | ISO dates (`CORE_ParseIsoDate`): missing versus invalid, calendar and leap-year validity, years 1900–9999, no time part; month ends (`CORE_IsMonthEnd`) | [File format](DATA_CONTRACT.md#file-format); expectations re-derived from the Python reading in [`test_data_contract.py`](../../tools/test_data_contract.py) | `exact` | Built (VBA) |
| `DATA-PARSE-02` | Balances (4 places) and rates (8 places, −0.05 to 0.25) (`CORE_ParseDecimal`, `CORE_ParseRate`): missing versus invalid, locale-like forms rejected | As `DATA-PARSE-01` | `rel 1e-15` balances; `exact` rates and codes | Built (VBA) |
| `DATA-CODES-01` | Segment codes and the currency allowlist reject every other value, including case and spacing variants | As `DATA-PARSE-01` | `exact` | Built (VBA) |

### Rate pass-through

| ID | Checks | Reference | Tolerance | Status |
| --- | --- | --- | --- | --- |
| `RATE-OLS-01` | Long-run OLS coefficients, residuals, $R^2$ on a 60-month synthetic pair | Independent OLS routine (Python or R), offline | `rel 1e-9` | Registered |
| `RATE-ADF-01` | ADF statistic at a fixed lag order, levels and differences | Independent ADF routine, offline | `rel 1e-8` | Registered |
| `RATE-EG-01` | Residual cointegration statistic and its critical value | Independent routine; critical values from `ECO-3` | `rel 1e-8` statistic; `rel 1e-6` critical value | Registered |
| `RATE-ECM-01` | `ECM` and `DIFF` coefficients for fixed $p, q$ | Independent OLS on the same regressors, offline | `rel 1e-9` | Registered |
| `RATE-ASYM-01` | Split of $\Delta r^m$ into positive and negative parts | Hand construction | `exact` | Registered |
| `RATE-HAC-01` | HAC standard errors, Bartlett kernel, fixed lag | Independent HAC routine (`ECO-4`), offline | `rel 1e-8` | Registered |
| `RATE-CONS-01` | Candidates with $\beta > 1$ or $\lambda \ge 0$ are rejected, not clipped | Hand-constructed series | `exact` (rejection) | Registered |
| `RATE-SIM-01` | Three-step scenario recursion | Hand computation | `abs 1e-12` | Registered |

### Stable amount

| ID | Checks | Reference | Tolerance | Status |
| --- | --- | --- | --- | --- |
| `STAB-CASH-01` | Transitions, survival and cash-out shares on the data-contract panel, including gap, closure, disappearance and migration | Hand computation from `accounts.csv` | `abs 1e-12` | Registered |
| `STAB-LOGIT-01` | Weighted fractional-logit coefficients and quasi-log-likelihood | Independent GLM routine (binomial family, logit link, frequency weights), offline | `rel 1e-6` | Registered |
| `STAB-ROBUST-01` | Robust covariance clustered by account | Independent routine, offline | `rel 1e-6` | Registered |
| `STAB-SEP-01` | A separated sample stops with an error and stores no parameters | Hand-constructed separated sample | `exact` (error) | Registered |
| `STAB-CONV-01` | Non-convergence within the iteration limit stops with an error | Hand-constructed sample | `exact` (error) | Registered |
| `STAB-LEAK-01` | No predictor at $t$ uses data dated after $t$ | Property test: altering later data leaves $x_{i,t}$ unchanged | `exact` | Registered |

### Decay

| ID | Checks | Reference | Tolerance | Status |
| --- | --- | --- | --- | --- |
| `DEC-LOG-01` | Log changes, $\hat\mu$, $\hat\sigma$ with and without break dummies | Hand computation on a short series | `abs 1e-12` | Registered |
| `DEC-MPA-01` | $MPA(h)$ for $h = 0, \dots, 120$ | Independent closed-form calculation, offline | `rel 1e-10` | Registered |
| `DEC-PROF-01` | Runoff, remainder, notional conservation, monotonicity | Hand computation | `abs 1e-8` currency units | Registered |
| `DEC-LIFE-01` | Mean life of the profile and of the core | Hand computation | `abs 1e-9` months | Registered |
| `DEC-CAP-01` | Overlay: core-share scaling and cutoff reduction for the maturity cap | Hand computation with stated caps | `abs 1e-8` | Registered |

### Common

| ID | Checks | Reference | Tolerance | Status |
| --- | --- | --- | --- | --- |
| `OOS-FREEZE-01` | Coefficients unchanged across forecast origins in the test window | Property test | `exact` | Registered |
| `OOS-LEAK-01` | A forecast from origin $t$ is unchanged when data after $t$ are altered | Property test | `exact` | Registered |
| `OOS-SCORE-01` | Bias, MAE, RMSE and exceedance rate on a small hand case | Hand computation | `abs 1e-12` | Registered |

## ➕ Adding a case

1. Register the ID, what it checks, its reference and tolerance here.
2. Commit the synthetic fixture under `tests/fixtures/` and the expected file
   under `tests/expected/`, with the provenance fields above.
3. Mark the case *built* in the same pull request.
