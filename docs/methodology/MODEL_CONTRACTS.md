<div align="center">

# 📐 Model Contracts

### What the rate, stable-amount and decay models estimate, and how each is independently validated

[![Status](https://img.shields.io/badge/Status-Accepted_(%239)-217346?style=flat-square)](#decisions)
[![Sources](https://img.shields.io/badge/Sources-partly_verified-d97706?style=flat-square)](SOURCES.md)
[![Units](https://img.shields.io/badge/Units-decimal_%C2%B7_months-0969da?style=flat-square)](../REPOSITORY_STRUCTURE.md#parameter-and-units-boundary)

<br>

**Fit and regulation kept apart · Frozen coefficients out of sample · No model validates itself**

</div>

---

This document is authoritative for **the specification of each behavioral
model: inputs, equations, estimation, constraints, diagnostics, outputs, and
how fit is separated from the regulatory overlay**. Inputs come from the
[data contract](DATA_CONTRACT.md); units follow the
[parameter and units boundary](../REPOSITORY_STRUCTURE.md#parameter-and-units-boundary);
external sources are cited by ID from the [source register](SOURCES.md); the
numerical cases that will test each model are registered in
[`TEST_CASES.md`](TEST_CASES.md); validation and backtesting obligations are in
[`VALIDATION_PLAN.md`](VALIDATION_PLAN.md).

> [!IMPORTANT]
> These contracts were **accepted by the owner on 2026-10-08 in issue #9**,
> including the [decisions](#decisions) below; see the
> [acceptance record](https://github.com/danielep71/VBA-IRRBB-Toolkit/issues/9#issuecomment-6068527520). Regulatory figures are labeled with their source; those resting on an
> unverified source in the [register](SOURCES.md) are provisional.

### Implementation hold points from the repository audit

The accepted baseline is not numerical validation. The owner-approved roadmap
of 2026-10-09 splits [#26](https://github.com/danielep71/VBA-IRRBB-Toolkit/issues/26)
into focused prerequisites. Each affected implementation must wait for its
reviewed contract amendment; the parent's v0.7.0 milestone does not defer the
earlier gates:

- ~~field availability at each forecast origin, and verified closure versus missing
  data/right censoring for W05 (#32, v0.2.0)~~ — decided on 2026-10-09 in #32:
  [point-in-time view](DATA_CONTRACT.md#look-ahead) and
  [transition outcomes](DATA_CONTRACT.md#outcomes);
- full ECM dynamic stability with dependent-variable lags, root convention and
  tolerance, beyond the beta/lambda bounds below (#33, v0.4.0);
- account-cohort versus aggregate compounding and the exact evolution of survival
  weights (50/50 balances, survival .9/.1, two months: 41 versus 25), together
  with reproducible fractional-logit conventions (#34, v0.5.0);
- deterministic model/lag ties, comparable estimation samples, and complete HAC
  settings and ECM/DIFF reporting (#33, v0.4.0);
- outstanding official-source verification and regulatory applicability (#35,
  v0.7.0, before each affected overlay);
- independent numerical benchmarks delivered with each model: Decay #36
  (v0.3.0), Rates #33 (v0.4.0), Stable #34 (v0.5.0);
- the Stable-Decay integration contract, reconciling existing-account survival
  with aggregate balances without double-counting runoff (#37, before v0.6.0
  integration).

These are explicit implementation blockers. Do not silently choose one
interpretation or report these models as validated.

Notation: $t$ is a month end, $B$ a balance in currency units, $r$ a rate as a
decimal per annum, $h$ a horizon in whole months. All models run per
**segment × currency**.

<a id="common"></a>

## 🧭 Common rules

### Segment granularity

Every model is estimated separately for each segment (`RET_TX`, `RET_NTX`,
`WHS_NFC`) and currency. Balances in different currencies are never pooled. A
segment and currency with fewer observations than the minimum below is not
estimated; the run reports it instead.

| Model | Minimum sample (proposed) |
| --- | --- |
| Rate pass-through | 60 consecutive monthly observations |
| Stable amount | 1,000 account transitions from at least 30 accounts |
| Decay | 36 consecutive monthly log changes |

### Estimation window and out-of-sample freeze

A run declares an **estimation window** $[T_s, T_e]$ and an optional **test
window** $(T_e, T_f]$, both inside the data window.

- Coefficients are estimated once, on the estimation window only, and then
  **frozen**. Nothing is re-estimated inside the test window; a recalibration
  is a separate, later run.
- Each month end $t$ in the test window is a **forecast origin**. A forecast
  from $t$ uses only observations dated on or before $t$.
- Rate and stable forecasts are scored at horizons $h = 1, \dots, 12$ by mean
  error (bias), mean absolute error and root mean squared error, per horizon.
- Decay profiles are scored by the **exceedance rate**: the share of
  (origin, horizon) pairs where the realized balance falls below the profile,
  compared with the stated tail probability.

### Confidence bands

| Model | Band | Not included |
| --- | --- | --- |
| Rate pass-through | Analytic $h$-step forecast interval from the fitted dynamics, assuming independent normal errors | Parameter uncertainty |
| Stable amount | Delta-method interval on the linear predictor using the robust covariance, mapped through the logistic function | Model-selection uncertainty |
| Decay | The minimum probable amount is itself the lower quantile; the median path is reported alongside | Parameter uncertainty in $\mu$ and $\sigma$ |

Each band states its coverage level and what it leaves out.

### Scenarios

A scenario is a market-rate path per currency, defined relative to the rate
observed at the forecast origin $r^m_{t_0}$:

| Scenario | Path for $k = 1, \dots, H$ |
| --- | --- |
| Base | $r^m_{t_0}$, held flat |
| Parallel up / down | $r^m_{t_0} \pm \Delta$ from the first month, held |
| Ramp up / down | $r^m_{t_0} \pm \Delta \cdot \min(k, 12)/12$ |

Shock sizes $\Delta$ are run settings, recorded with the results. The
supervisory scenarios and the post-shock floor of `REG-3` are applied as an
[overlay](#regulatory-overlay), not as part of the fit. Each model below states
how a scenario reaches it.

<a id="rate-model"></a>

## 📈 Rate pass-through model

**Question:** how does the administered customer rate respond to the market
rate?

### Inputs

- $r^c_t$: balance-weighted average `customer_rate` of the **non-indexed**
  accounts in the segment and currency, weights $B_{i,t}$; account-months with
  a missing rate are left out; a month with no weight is missing.
- $r^m_t$: the market-rate series selected for the run, same currency.

Indexed accounts follow their contractual formula and are outside this model.

### Specification

1. **Unit roots** (`ECO-2`): augmented Dickey–Fuller tests on $r^c_t$ and
   $r^m_t$ in levels and first differences, with a constant, lag order chosen
   by AIC up to 6, at the 5 % level.
2. **Long-run relation**, by OLS:

   $$r^c_t = \alpha + \beta\, r^m_t + u_t$$

3. **Cointegration** (`ECO-1`, `ECO-3`): ADF test on $\hat u_t$ without a
   constant, against Engle–Granger critical values, at the 5 % level.
4. **Short-run dynamics**, by OLS, one of these candidate forms:

   | Form | Equation | Allowed when |
   | --- | --- | --- |
   | `ECM` | $\Delta r^c_t = \sum_{i=0}^{p} \delta_i \Delta r^m_{t-i} + \sum_{j=1}^{q} \phi_j \Delta r^c_{t-j} + \lambda \hat u_{t-1} + \varepsilon_t$ | Cointegration found: a unit root in $\hat u_t$ is rejected |
   | `ECM-ASYM` | as `ECM`, with $\delta_i$ split into $\delta_i^{+}$ on $\max(\Delta r^m_{t-i}, 0)$ and $\delta_i^{-}$ on $\min(\Delta r^m_{t-i}, 0)$ | As `ECM`, and a Wald test rejects $\delta^{+} = \delta^{-}$ at 5 % |
   | `DIFF` | as `ECM` without the $\lambda \hat u_{t-1}$ term | No cointegration found |

5. **Lag selection:** $p, q \in \{0, \dots, 3\}$, chosen by the Bayesian
   information criterion among candidates that meet every constraint below.

### Constraints

A candidate that violates a constraint is **rejected, never clipped**:

- $0 \le \beta \le 1$;
- $-1 < \lambda < 0$ (adjustment towards the long-run relation);
- $\beta = 1$ is used only if estimated, never imposed. A test of $\beta = 1$ is
  reported with HAC standard errors (`ECO-4`).

If no candidate is valid, the model is not estimated and the run says so.

### Diagnostics reported

Coefficients with HAC standard errors (`ECO-4`), $R^2$, residual
autocorrelation up to 12 lags, the unit-root and cointegration statistics with
their critical values, and the lag search table.

### Outputs and scenarios

The parameter set holds $\alpha$, $\beta$, the chosen form, all short-run
coefficients and $\hat\sigma_\varepsilon$. A scenario is applied by iterating
the fitted equation forward from the last observed state, with
$u_t = r^c_t - \alpha - \beta r^m_t$. A customer-rate floor, if configured
(default none), is a run setting applied to the simulated path and reported
separately from the fit.

<a id="stable-model"></a>

## 🏦 Stable-amount model

**Question:** what share of an existing balance is still there one month
later, and how does that depend on observable drivers?

### Cash-out and survival

For each account $i$ and consecutive month ends $t \to t+1$ inside the
estimation window with $B_{i,t} > 0$:

- the transition's [outcome](DATA_CONTRACT.md#outcomes) is *observed*
  ($B_{i,t+1}$ is the next observation) or a *verified closure* in
  $(t, t+1]$ ($B_{i,t+1} = 0$);
- a gap (`W01`), a censored disappearance (`W05`) and the last month end of the
  window are right-censored or excluded and give no cash-out label;
- a migration (`W03`) stays in the segment of month $t$ and is counted.

$$s_{i,t} = \frac{\min(B_{i,t+1},\, B_{i,t})}{B_{i,t}} \in [0, 1], \qquad c_{i,t} = 1 - s_{i,t}$$

$s$ is the **one-month survival share** and $c$ the **cash-out share**. Growth
in an existing balance is not survival; new accounts enter only from their
second month.

### Specification

Fractional-response logit (`ECO-5`):

$$\mathbb{E}[s_{i,t} \mid x_{i,t}] = \Lambda(x_{i,t}'\theta), \qquad \Lambda(z) = \frac{1}{1 + e^{-z}}$$

estimated by Bernoulli quasi-maximum likelihood with **weights $B_{i,t}$**, so
that the fit reproduces balance survival rather than account counts. Weights
are divided by their mean before estimation.

### Predictors

Every predictor is **known at $t$**: it is computed from the
[point-in-time view](DATA_CONTRACT.md#look-ahead) at $t$, so nothing dated or
learned after $t$, including a back-filled `close_date`, enters $x_{i,t}$.

| Predictor | Definition |
| --- | --- |
| Constant | — |
| Rate spread | $r^c_{i,t} - r^m_t$ (missing customer rate: indicator plus zero) |
| Market-rate change | $r^m_t - r^m_{t-1}$ |
| Account age | Years from `open_date` to $t$; unknown age as an indicator |
| Size | $\ln B_{i,t}$ |
| Indexed | `indexed` flag |

**Selection:** backward elimination at the 5 % level using robust Wald tests;
the constant is always kept. The elimination trace is part of the diagnostics.

### Estimation, convergence and failure

- Newton–Raphson on the quasi-log-likelihood, starting from $\theta = 0$.
- Converged when $\max_k |\Delta\theta_k| < 10^{-8}$ and the relative change
  in quasi-log-likelihood is below $10^{-10}$; at most 100 iterations.
- Standard errors: robust sandwich covariance, clustered by account.
- **Failure is an error, never a result:** no convergence within the limit, a
  singular Hessian, or signs of separation (any $|\theta_k| > 50$, or fitted
  values within $10^{-10}$ of 0 or 1) stop the estimation and no parameters
  are stored.

### Outputs and scenarios

The parameter set holds $\theta$, its robust covariance, the selected
predictors and the trace. The segment's expected survival in month $k$ is the
balance-weighted mean of $\Lambda(x'\theta)$ over its accounts, with rate
predictors following the scenario path, account age increasing by one month
per step, and other attributes held at the origin. The stable amount over $H$
months is

$$SA(H) = B_{t_0} \prod_{k=1}^{H} \bar s_k$$

The sign of the spread effect is reported as a diagnostic (a higher market
rate relative to the customer rate is expected to lower survival); it is not
imposed.

<a id="decay-model"></a>

## 📉 Decay model

**Question:** how fast could the segment's balance run off, at a prudential
confidence level?

### Specification

1. **Balance unit:** the aggregate balance $B_t = \sum_i B_{i,t}$ of the
   segment and currency, including new accounts. Censored disappearances
   (`W05`) reduce this total as supplied; their balance is reported in the
   lineage per month end, and whether the decay series adjusts for it is
   decided in #36.
2. **Log changes** for consecutive month ends with $B > 0$:

   $$g_t = \ln B_t - \ln B_{t-1}$$

3. **Structural breaks** declared in the run configuration enter as impulse
   dummies $D_{k,t}$, estimated by OLS:

   $$g_t = \mu + \sum_k \kappa_k D_{k,t} + e_t$$

   $\hat\sigma$ is the residual standard deviation with $n - 1 - K$ degrees of
   freedom for $K$ dummies.
4. **Prudential drift:** $\mu^{*} = \min(\hat\mu, 0)$, so balance growth is
   never credited.
5. **Minimum probable amount** at confidence $c$ (proposed $c = 0.99$), with
   $z = \Phi^{-1}(1 - c)$:

   $$MPA(h) = B_{t_0} \exp\!\left(\mu^{*} h + z\, \hat\sigma \sqrt{h}\right), \qquad h = 0, 1, \dots, H$$

   Because $\mu^{*} \le 0$ and $z < 0$, $MPA(h)$ is non-increasing in $h$; the
   toolkit still checks it.

### Profile, core and mean life

With cutoff $H$ (proposed 120 months), monthly runoff and the remainder are

$$o_h = MPA(h-1) - MPA(h), \quad h = 1, \dots, H, \qquad \text{remainder } MPA(H) \text{ placed at month } H$$

so that $\sum_{h=1}^{H} o_h + MPA(H) = B_{t_0}$ (**notional conserved**).

- **Non-core** share: $1 - MPA(1) / B_{t_0}$, the part expected to leave in the
  first month.
- **Core** amount: $MPA(1)$.
- **Mean life** of the whole profile, in months:

  $$L = \frac{\sum_{h=1}^{H} h\, o_h + H \cdot MPA(H)}{B_{t_0}}$$

- **Mean life of the core**: the same sum from $h = 2$, divided by $MPA(1)$.

### Diagnostics reported

$\hat\mu$, $\hat\sigma$, break coefficients, residual autocorrelation and
normality tests, and the empirical $(1-c)$ quantile of the residuals next to
the normal one. A normality rejection is reported; the normal quantile remains
the contract until the owner decides otherwise.

### Scenarios

The decay model has no rate input. Rate scenarios reach the behavioral
profile only through the stable-amount and rate models; this limitation is
stated with every decay result.

<a id="regulatory-overlay"></a>

## ⚖️ Econometric fit vs regulatory overlay

The toolkit keeps two separate layers and always reports both:

| Layer | Content | Source |
| --- | --- | --- |
| **Fit** | Estimated parameters, profiles and bands from the data, as specified above | The data and the methods cited |
| **Overlay** | Constraints on what may be used for regulatory repricing assumptions | `REG-1`, `REG-2`, `REG-3` |

The overlay never changes a fitted parameter. It produces a separate,
constrained profile, labeled with the source ID and the verification status of
each constraint applied:

1. **Non-core in the overnight bucket**: the overlay profile places the
   non-core amount in the overnight bucket, not in month 1 (`REG-1` ¶114,
   verified).
2. **Core-share cap** per segment: if $MPA(1) / B_{t_0}$ exceeds the cap
   $k$, the core profile is scaled by $k\, B_{t_0} / MPA(1)$ for every
   $h \ge 1$ and the difference moves to the overnight bucket, so the profile
   stays non-increasing and still sums to $B_{t_0}$. Caps per `REG-1` ¶113,
   Table 2 (verified): retail transactional 90 %, retail non-transactional
   70 %, wholesale 50 %.
3. **Average-maturity cap** per segment: if the core mean life exceeds the cap,
   the cutoff is reduced to the largest $H' \le H$ that meets it, with the
   remainder placed at $H'$. Caps per `REG-1` ¶115, Table 2 (verified):
   5 years, 4.5 years, 4 years. An EU 5-year cap on the average repricing
   maturity of non-maturity deposits is attributed to `REG-2` and stays
   **provisional** until `REG-2` is verified; `REG-3` contains no such cap.
4. **Supervisory scenarios and floor** for scenario runs, per `REG-3`
   (verified): six scenarios for the economic value of equity (parallel up,
   parallel down, steepener, flattener, short rates up, short rates down) and
   two for net interest income (parallel up, parallel down); a post-shock lower
   bound of $F(m) = \min(-150 + 3m,\ 0)$ basis points at maturity $m$ in
   years, replaced by the observed rate where that is lower. Shock sizes per
   currency are not yet verified, so supervisory shock runs stay provisional.

The `REG-1` categories map to the data-contract segments as follows:
retail transactional to `RET_TX`, retail non-transactional to `RET_NTX`, and
wholesale to `WHS_NFC`. `WHS_NFC` covers only the non-financial part of `REG-1`
wholesale, and `REG-1` treats qualifying small businesses (managed as retail,
total liabilities below EUR 1 million) as retail; segments are supplied with the
data ([data contract](DATA_CONTRACT.md#segmentation)), so the data owner applies
that rule.

An overlay resting on an unverified source is shown as **provisional** and may
not be presented as compliant. A statistical decay horizon is never evidence
that a repricing assumption is permitted.

<a id="validation"></a>

## 🧪 Independent validation

Every model is tested against reference values that the toolkit did not
produce, registered in [`TEST_CASES.md`](TEST_CASES.md) with their provenance
and tolerance. Backtesting follows the
[out-of-sample freeze](#estimation-window-and-out-of-sample-freeze) above and
[`VALIDATION_PLAN.md`](VALIDATION_PLAN.md).

<a id="decisions"></a>

## ✅ Decisions

Accepted by the owner on 2026-10-08 in issue #9.

| # | Decision | Outcome |
| ---: | --- | --- |
| 1 | Model granularity | Separate model per segment × currency, all three models |
| 2 | Minimum samples | 60 months (rate), 1,000 transitions from 30 accounts (stable), 36 log changes (decay) |
| 3 | Rate model lag limit and selection | $p, q \le 3$, BIC among valid candidates; asymmetry only if a Wald test rejects symmetry at 5 % |
| 4 | Stable model predictors and selection | Spread, market-rate change, age, size, indexed; backward elimination at 5 % with robust tests |
| 5 | Migration in the stable model | Transition counted in the segment of the starting month |
| 6 | Decay confidence level | 99 % |
| 7 | Decay drift | $\mu^{*} = \min(\hat\mu, 0)$ |
| 8 | Profile cutoff | 120 months |
| 9 | Meeting the average-maturity cap | Reduce the cutoff until the cap is met |
| 10 | Default scenario shocks | Base, parallel and 12-month ramp, ±200 bp, until the `REG-3` shock sizes per currency are verified |

---

**Model principle:** state the equation, the data it sees and the constraint it
must meet; reject what fails, and keep what the data say apart from what a
regulation allows.
