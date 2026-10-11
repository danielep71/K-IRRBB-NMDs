# Behavioral IRRBB methodology roadmap

This is a model-development roadmap, **not a validated specification**. Do not assume that a generic econometric fit satisfies prudential constraints or model risk standards.

## Common data contract

Define segments, currencies, observation frequency, account identifiers, balance signs, rate units, missing-period handling, portfolio migrations, structural breaks, indexed/non-indexed products, winsorization and time alignment. Keep transformations deterministic and independently testable.

The contract accepted in issue #4 is [`DATA_CONTRACT.md`](DATA_CONTRACT.md): field dictionary, segmentation, reconciliation, validation codes, dataset size, lineage and synthetic fixtures.

## Model families

The specifications accepted in issue #9 are in [`MODEL_CONTRACTS.md`](MODEL_CONTRACTS.md), with the numerical reference cases registered in [`TEST_CASES.md`](TEST_CASES.md). The requirements below are what those contracts must satisfy.

### Rate / pass-through

Specify long-run equilibrium and short-run ECM independently. Record the market rate, spread factors, transformations, unit roots, cointegration or other evidence supporting ECM, identification constraints, statistical significance, dynamic stability, up/down asymmetry and out-of-sample performance. Do not assume beta = 1 without a documented test.

### Stable amount

Define exactly how account-level cash-out, stable amount and missing subsequent observations are measured. Specify any binomial-logit interpretation and its weighting/aggregation, predictor availability at forecast date, selection procedure, convergence and separation handling. Stable shares and scenario responses must be bounded and economically interpretable.

### Decay

Define the balance unit (aggregate vs per-account), logarithmic changes, structural dummy treatment, residual diagnostics, prudential quantiles, mean-life calculation and cutoff schedule. Ensure positive, monotonic profiles and conservation of notional. Separate estimated behavioral duration from regulatory constraints on repricing assumptions.

## Regulatory overlay

Maintain an explicit reference register for the applicable EBA Guidelines, EU RTS and supervisory expectations before implementing any rule: [`SOURCES.md`](SOURCES.md). `REG-1` and `REG-3` are verified for the provisions listed in its verification record; `REG-2` and the econometric sources are not, and anything resting on them stays provisional. Separate **internal measurement assumptions** from **standardized method assumptions**. Regulatory rule versions and sources must be attached to test evidence, not inferred from the econometric output.

## Reproducibility

Record input provenance, data window, feature pipeline version, model formula, optimization settings, rounding rules, resulting coefficients, diagnostics, run timestamps and independent benchmark outputs.

The initial milestone establishes these contracts. No model is implemented at v0.1.0.
