<div align="center">

# 🧾 Data Contract

### Input data, segmentation and lineage for the IRRBB behavioral models

[![Status](https://img.shields.io/badge/Status-Accepted_(%234)-217346?style=flat-square)](#decisions)
[![Data](https://img.shields.io/badge/Data-synthetic_only-217346?style=flat-square)](../../CONTRIBUTING.md#data-confidentiality-and-provenance)
[![Fixtures](https://img.shields.io/badge/Fixtures-reference--checked-0969da?style=flat-square)](#synthetic-fixtures)

<br>

**One row per account per month end · Stocks, not flows · All-or-nothing import · Every run traceable to its data**

</div>

---

This document is authoritative for **the input datasets, their fields, units and
validation, segmentation and reconciliation, and data lineage**. Units at the
public boundary follow
[`REPOSITORY_STRUCTURE.md`](../REPOSITORY_STRUCTURE.md#parameter-and-units-boundary);
model definitions are in [`README.md`](README.md); validation obligations in
[`VALIDATION_PLAN.md`](VALIDATION_PLAN.md).

> [!IMPORTANT]
> This contract was **accepted by the owner on 2026-10-08 in issue #4**,
> including the [decisions](#decisions) below; see the
> [acceptance record](https://github.com/danielep71/VBA-IRRBB-Toolkit/issues/4#issuecomment-6068526038).
> The point-in-time, closure and censoring amendment was decided by the owner
> on 2026-10-09 in [#32](https://github.com/danielep71/VBA-IRRBB-Toolkit/issues/32)
> ([decisions 6–8](#decisions)). Other audit amendments remain tracked in
> [#26](https://github.com/danielep71/VBA-IRRBB-Toolkit/issues/26). Only synthetic data is ever
> committed or attached. Real data is loaded locally
> and never enters Git, issues or pull requests.

<a id="datasets"></a>

## 🗃️ Datasets

| Dataset | Granularity | Used by |
| --- | --- | --- |
| **Account panel** | One row per account per month end at which the account exists | Stable-amount, decay and rate models; segmentation; reconciliation |
| **Market rates** | One row per rate series per month end | Rate pass-through model and scenarios |

Customer-rate averages, segment balances, cash-outs and log balance changes
are **derived** by the toolkit from these inputs, never supplied as inputs.

<a id="file-format"></a>

## 📄 File format

Both datasets use the same plain, locale-independent CSV subset:

- ASCII text; LF or CRLF line endings; no byte-order mark;
- comma separator, no quoting, no embedded commas or line breaks;
- the first line is the exact header below, fields in the listed order;
- dates as `YYYY-MM-DD` with a year from 1900 to 9999 (the Excel date system);
  decimals with `.` and no thousands separator or exponent, at least one digit
  before the `.` and at least one after it;
- an empty field means *missing*. No sentinel values (`0`, `-999`, `N/A`).

Parsing never uses Excel's locale-dependent conversions (`CDate`, `CDbl` on
text, or sheet auto-formatting), so the same file gives the same result on
every host.

<a id="account-panel"></a>

## 🏦 Account panel

Header:
`as_of_date,account_id,segment,currency,balance,customer_rate,indexed,open_date,close_date`

| Field | Type and format | Unit | Required | Rules |
| --- | --- | --- | :---: | --- |
| `as_of_date` | ISO date | — | ✅ | Calendar month end |
| `account_id` | 1–32 characters `A–Z a–z 0–9 _ -` | — | ✅ | Opaque pseudonymous key; never a real account number |
| `segment` | Code | — | ✅ | One of the [segment codes](#segmentation) |
| `currency` | ISO 4217 code, 3 upper-case letters | — | ✅ | Constant for an account |
| `balance` | Decimal, up to 4 places | Currency units | ✅ | End-of-month **stock**, ≥ 0 |
| `customer_rate` | Decimal, up to 8 places | Decimal per annum | — | Rate paid on the account at month end; −0.05 to 0.25 |
| `indexed` | `0` or `1` | — | ✅ | `1` = contractually indexed to a market rate |
| `open_date` | ISO date | — | — | Not after `as_of_date` |
| `close_date` | ISO date | — | — | Not before `open_date`; no row dated on or after it (the account has no month-end row in its closure month) |

The key is (`as_of_date`, `account_id`). A missing `customer_rate` excludes
that account-month from rate-model inputs only. Missing `open_date` means
opened before the data window or unknown.

<a id="market-rates"></a>

## 📈 Market rates

Currency validation requires a versioned, explicit run allowlist of ISO 4217
codes; matching three uppercase letters alone is insufficient. The synthetic
reference checker uses `EUR` and `USD`. Extending that allowlist requires reviewed
configuration and provenance; it does not claim to maintain the full ISO register.
Balances must convert to finite VBA Double values; overflow is `E04`. Aggregate
overflow must also fail the future importer rather than produce an infinite result.
Known non-empty opening/closure dates must be consistent throughout an account's
history (`E11`); an omitted repeat does not erase a known date. A row dated on or
after any known closure is `E10`, regardless of file ordering. These consistency checks do
not make future closure information eligible as a predictor.

Header: `as_of_date,rate_id,currency,tenor_months,rate`

| Field | Type and format | Unit | Required | Rules |
| --- | --- | --- | :---: | --- |
| `as_of_date` | ISO date | — | ✅ | Calendar month end |
| `rate_id` | 1–32 characters `A–Z a–z 0–9 _ -` | — | ✅ | Series identifier, e.g. `EUR_1M` |
| `currency` | ISO 4217 code | — | ✅ | Constant for a series |
| `tenor_months` | Positive integer | Months | ✅ | Constant for a series |
| `rate` | Decimal, up to 8 places | Decimal per annum | ✅ | −0.05 to 0.25 |

The key is (`as_of_date`, `rate_id`). Which series a model uses is a run
setting, recorded in the parameter set.

<a id="segmentation"></a>

## 🧩 Segmentation

| Code | Segment |
| --- | --- |
| `RET_TX` | Retail, transactional |
| `RET_NTX` | Retail, non-transactional |
| `WHS_NFC` | Wholesale, non-financial corporate |

These follow the usual supervisory split of non-maturity deposits by
counterparty and transactional use. Adding a segment needs an issue and a
change to this table and the fixtures.

**Assignment is deterministic and supplied, not inferred:** each row carries
its segment, and the toolkit never re-segments an account. An account whose
segment changes between months is a **migration** (warning `W03`): each month
counts in the segment stated for that month, and the model specification
(issue #9) decides how a transition across segments enters estimation.

<a id="reconciliation"></a>

## ⚖️ Reconciliation

After import, the toolkit reports for every (`as_of_date`, `segment`,
`currency`) the number of accounts and the total balance, and checks that:

- totals retain input precision; round only for presentation, after reconciliation;
- the segment totals for each date and currency add up to the panel total for
  that date and currency; and
- if the user supplies control totals, the counts match exactly and balances
  within 0.005 currency units.

Balances in different currencies are never added together. A reconciliation
difference fails the import.

<a id="validation"></a>

## 🚦 Validation and invalid input

**The import is all-or-nothing.** Any error rejects the whole dataset: no
partial dataset is kept, and earlier results stay marked stale. Warnings do not
stop the import; they are reported and recorded in the lineage.

Each row reports its **first** failing rule, checked in the order below, so the
same file always gives the same findings in the same order.

| Code | Error | Rows checked |
| --- | --- | --- |
| `E01` | Header or number of fields does not match the contract | File, then each row |
| `E03` | A required field is empty | Each row |
| `E02` | A date is not a valid ISO date, or `as_of_date` is not a month end | Each row |
| `E04` | A value cannot be parsed or is out of range (ID characters, balance or rate format, rate bounds, `indexed` not `0`/`1`) | Each row |
| `E06` | Unknown segment code | Each row |
| `E07` | Invalid currency code | Each row |
| `E09` | Negative balance | Each row |
| `E11` | Inconsistent known history dates, `open_date` after `as_of_date`, or `close_date` before `open_date` | Each affected row |
| `E10` | Row dated on or after the account's `close_date` | Each row |
| `E05` | Duplicate (`as_of_date`, `account_id`) | Second and later occurrences |
| `E08` | Account currency or market-series currency/tenor differs from its earlier rows | Each later row |
| `E12` | **Missing extract**: a month end inside the data window has no account rows | Once per missing month end |

| Code | Warning | Reported at |
| --- | --- | --- |
| `W01` | **Gap**: account missing in a month between its first and last observation | The missing month |
| `W02` | Zero balance (dormant account) | The row |
| `W03` | **Migration**: segment differs from the previous observation | The row |
| `W04` | **Outlier candidate**: balance more than 10× or less than 1/10 of the previous month, both positive | The later month |
| `W05` | **Censored disappearance**: account stops before the window end without a closure recorded in the following month (no `close_date`, or one later than the next month end) | Its last observation |
| `W06` | Row outside the data window; excluded and counted | The row |

<a id="outcomes"></a>

### Missing observations, openings, closures and censoring

A transition runs from an observed month end $t$ of an account to the next
month end $t+1$. Its **outcome** is the first of these that applies:

| Outcome | Condition | In estimation |
| --- | --- | --- |
| **End of sample** | $t$ is the last month end of the data window | No outcome: right-censored |
| **Observed** | The account has a row at $t+1$ | Transition with the observed balance |
| **Verified closure** | No row at $t+1$, and the account's `close_date` lies in $(t, t+1]$ | Full cash-out: balance $0$ at $t+1$ |
| **Gap** (`W01`) | No row at $t+1$, but a later row exists | Excluded; the gap is never filled |
| **Censored disappearance** (`W05`) | No row at $t+1$, no later row, and no closure in $(t, t+1]$ | No outcome: right-censored |

- Only a recorded `close_date` is a **verified closure**. Disappearance alone
  is never treated as cash-out: the account leaves the risk set after its last
  observation. A `close_date` later than $t+1$ for an account with no further
  rows does not make the missing months a closure; the account is censored.
- An account **opens** at its first observation. Its first month has no
  predecessor and enters no transition as an outcome.
- A **missing extract** (a month end in the window with no account rows) is a
  data-delivery fault, not a mass closure or a panel-wide gap: the import is
  rejected (`E12`).
- A **migration** (`W03`) does not change the outcome: the transition belongs to
  the segment stated at $t$, which is known at $t$.
- **Eligibility:** a transition enters an account-level estimation sample only
  with outcome *observed* or *verified closure*. Model-specific conditions (for
  example a positive balance at $t$) are set in
  [`MODEL_CONTRACTS.md`](MODEL_CONTRACTS.md).
- **Lineage** records, per month end, segment and currency, the number of
  verified closures, censored disappearances with their last balance, gaps and
  end-of-sample accounts, and the number of `close_date` values masked at each
  forecast origin ([point-in-time availability](#look-ahead)).

### Structural breaks and outliers

The import never alters, removes or winsorizes a value. Outlier candidates are
flagged (`W04`); structural breaks are declared in the run configuration with a
date, segment and reason. How flagged values and breaks enter estimation is a
model setting (issue #9), recorded in the parameter set, never an import side
effect.

<a id="look-ahead"></a>

### Point-in-time availability and look-ahead prevention

Every run has a data window (first and last month end). Rows outside it are
excluded (`W06`). The window is part of the lineage and the parameter set.

Each row is an **observation** of the account at its `as_of_date`: segment,
currency, balance, customer rate, `indexed` and `open_date` are known at that
date (`open_date` cannot be later, `E11`). `close_date` is the exception: an
extract may back-fill it on earlier rows, so it can be **learned after** the
row's `as_of_date`.

The **point-in-time view at a forecast origin** $t$ is what every predictor,
backtest and forecast at $t$ may use:

1. only rows with `as_of_date` $\le t$;
2. on those rows, `close_date` is shown only if it is $\le t$; a later
   `close_date` is **masked** (treated as empty) in the view.

A masked `close_date` is outcome information: it decides a later transition's
outcome ([above](#outcomes)) and never enters a predictor at $t$. Consequently
changing, adding or removing any information learned after $t$ (later rows,
later closure dates) cannot change any predictor at $t$. The reference checker
tests exactly this. No extract-vintage column is required; the rows are
accepted as supplied.

**Timely closure reporting (source requirement).** Without a vintage field the
view cannot tell a closure known when it happened from one the source reported
months later. The contract therefore **assumes a closure is known at the month
end of the month in which it occurs**: a `close_date` $\le t$ is treated as known
at $t$. A per-account predictor cannot expose it (an account with a row at $t$
cannot be closed on or before $t$, `E10`), but a fit or backtest at $t$ uses it to label
earlier transitions as closures. A source that reports closures later than the
month end of their occurrence breaks this assumption and **must not be loaded
as-is**; supporting it needs an availability field and a reopened decision 6.
The import cannot detect late reporting, so the user attests timely reporting
for each import and the attestation is part of the [lineage](#lineage).

<a id="dataset-size"></a>

## 📦 Dataset size

A worksheet holds at most 1,048,576 rows, while an account panel easily
exceeds that (50,000 accounts × 120 months = 6 million rows).

**Decided:** the account panel is read from the CSV file by VBA's own file
input, streamed line by line into arrays and aggregates in memory. Worksheets
hold only summaries, parameters and results, never the raw panel. No external
engine is involved ([runtime dependencies](../../INSTALLATION.md#runtime-dependencies)).

Reading a user-selected input file is a file-system access that
[`SECURITY.md`](../../SECURITY.md#security-scope) requires to be agreed in an
issue: read-only, the file the user picks, nothing written next to it. This
issue (#4) is that agreement; `SECURITY.md` is updated in the same change as
the importer.
Memory use and import time for large panels are measured when the importer
exists; nothing about performance is claimed here.

<a id="lineage"></a>

## 🔗 Lineage and reproducible import

Every successful import records, in the workbook and in each parameter set
built from it:

| Item | Content |
| --- | --- |
| Source | File name (no directory path), size in bytes, SHA-256 of the file bytes |
| Counts | Rows read, accepted, excluded by window; accounts; warnings by code |
| Settings | Data window, segment list, rate series selected |
| Attestation | Closure dates reported by the month end of the closure ([timely closure reporting](#look-ahead)) |
| Code | Toolkit commit SHA |
| Time | Import timestamp with UTC offset |

**Reproducibility:** the same file, settings and commit give the same dataset,
the same findings in the same order and the same fingerprint, on any supported
host. The fingerprint is computed by the toolkit itself, since no external tool
or reference may be used at runtime.

<a id="synthetic-fixtures"></a>

## 🧪 Synthetic fixtures

| File | Content |
| --- | --- |
| [`tests/fixtures/data_contract/accounts.csv`](../../tests/fixtures/data_contract/accounts.csv) | 10 accounts, 6 month ends (Jan–Jun 2025), 51 rows: two currencies, all three segments, an opening mid-window, a closure, a gap, a dormant account, a migration, a disappearance and an outlier |
| [`tests/fixtures/data_contract/accounts_invalid.csv`](../../tests/fixtures/data_contract/accounts_invalid.csv) | One invalid row per error rule `E02`–`E11` |
| [`tests/fixtures/data_contract/market_rates.csv`](../../tests/fixtures/data_contract/market_rates.csv) | `EUR_1M` and `USD_1M`, 6 month ends |
| [`tests/fixtures/data_contract/availability.csv`](../../tests/fixtures/data_contract/availability.csv) | 7 accounts, Jan–Jun 2025: every [transition outcome](#outcomes), back-filled closure dates before and after the window end, and a migration |
| [`tests/expected/data_contract/availability_outcomes.json`](../../tests/expected/data_contract/availability_outcomes.json) | Hand-assigned outcome for every transition, warnings, and the `close_date` values visible at each origin |
| [`tests/expected/data_contract/accounts_summary.json`](../../tests/expected/data_contract/accounts_summary.json) | Hand-computed totals per date, segment and currency, and every expected warning |
| [`tests/expected/data_contract/accounts_invalid_findings.json`](../../tests/expected/data_contract/accounts_invalid_findings.json) | Expected first error per invalid line |

The fixtures are hand-constructed; the expected values were computed by hand
from the rows. `tools/test_data_contract.py`, run by `python tools/check.py`,
re-reads the fixtures with an independent Python reading of this contract and
fails if fixtures, expectations and rules disagree. It also runs the
leakage tests: altering, adding or removing information learned after an
origin leaves every predictor at that origin unchanged, and dropping a month
end from the window gives `E12`. The VBA importer, when it exists, must
reproduce the same expected files.

<a id="decisions"></a>

## ✅ Decisions

Decisions 1–5 were accepted by the owner on 2026-10-08 in issue #4;
decisions 6–8 on 2026-10-09 in issue #32.

| # | Decision | Outcome |
| ---: | --- | --- |
| 1 | Input route for the account panel | CSV file streamed by VBA file input, with the read-only file access added to `SECURITY.md` |
| 2 | Segment list | `RET_TX`, `RET_NTX`, `WHS_NFC` |
| 3 | Negative balances | Invalid (`E09`): overdrafts are assets and out of scope |
| 4 | Outlier threshold for `W04` | Factor 10 month on month |
| 5 | Data fingerprint | SHA-256 implemented in VBA, tested against published test vectors |
| 6 | Back-filled `close_date` (#32, 2026-10-09) | Accepted as supplied; masked in the point-in-time view at origins before it; closures assumed reported by the month end in which they occur, attested per import |
| 7 | Disappearance without a verified closure (#32, 2026-10-09) | Right-censored (`W05`), never cash-out |
| 8 | Month end with no rows inside the window (#32, 2026-10-09) | Import rejected (`E12`) |

---

**Data principle:** keep the input as it was given, reject what breaks the
contract, flag what looks odd, and record exactly what each result was built
from.
