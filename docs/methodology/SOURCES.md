<div align="center">

# 📚 Source Register

### Every external source the model contracts rely on, and whether it has been checked

[![Status](https://img.shields.io/badge/Verified-REG--1_%C2%B7_REG--3-217346?style=flat-square)](#register)
[![Rule](https://img.shields.io/badge/Rule-cite%2C_never_copy-6f42c1?style=flat-square)](../../CONTRIBUTING.md#data-confidentiality-and-provenance)

</div>

---

This register is authoritative for **which external sources the toolkit relies
on and whether each has been verified against its official text**. Model
contracts in [`MODEL_CONTRACTS.md`](MODEL_CONTRACTS.md) cite sources by ID.
Sources are cited by reference only; their text is never copied into the
repository.

> [!WARNING]
> `REG-1` and `REG-3` were verified on 2026-10-08 by the owner's research
> against the official BIS and EUR-Lex texts, for the provisions listed in
> [Verification record](#verification-record). `REG-2` and the econometric
> sources are not verified. A regulatory rule may be implemented only from a
> verified provision, recorded with its exact location and version.

<a id="register"></a>

## Register

### Regulatory and supervisory

| ID | Source | Relied on for | Verified |
| --- | --- | --- | :---: |
| `REG-1` | Basel Committee on Banking Supervision, *Interest rate risk in the banking book*, Standards, April 2016 (BCBS d368); consolidated in the Basel Framework, SRP31 (version effective 1 January 2026, published 16 July 2024) | Non-maturity deposit categories; the standardized framework's caps on the proportion and average maturity of core deposits per category; non-core deposits in the overnight bucket | ☑ |
| `REG-2` | European Banking Authority, *Guidelines on IRRBB and CSRBB*, EBA/GL/2022/14 (exact title and dates not yet verified) | Expectations on behavioral assumptions for non-maturity deposits; the EU 5-year cap on their average repricing maturity, as reported by secondary sources | ☐ |
| `REG-3` | Commission Delegated Regulation (EU) 2024/856 of 1 December 2023 supplementing Directive 2013/36/EU of the European Parliament and of the Council with regard to regulatory technical standards specifying the supervisory shock scenarios, the common modelling and parametric assumptions and what constitutes a large decline. OJ L, 2024/856, 24.4.2024 | Supervisory shock scenarios for EVE and NII; the post-shock interest-rate floor | ☑ |

### Econometric and statistical

| ID | Source | Relied on for | Verified |
| --- | --- | --- | :---: |
| `ECO-1` | Engle, R. F. and Granger, C. W. J. (1987), "Co-integration and error correction: representation, estimation, and testing", *Econometrica* 55(2), 251–276 | Two-step estimation of the long-run relation and the error-correction model | ☐ |
| `ECO-2` | Dickey, D. A. and Fuller, W. A. (1979), "Distribution of the estimators for autoregressive time series with a unit root", *Journal of the American Statistical Association* 74, 427–431 | Unit-root tests on the rate series | ☐ |
| `ECO-3` | MacKinnon, J. G. (2010), "Critical values for cointegration tests", Queen's Economics Department Working Paper 1227 | Critical values for the residual-based cointegration test | ☐ |
| `ECO-4` | Newey, W. K. and West, K. D. (1987), "A simple, positive semi-definite, heteroskedasticity and autocorrelation consistent covariance matrix", *Econometrica* 55(3), 703–708 | Standard errors robust to autocorrelation | ☐ |
| `ECO-5` | Papke, L. E. and Wooldridge, J. M. (1996), "Econometric methods for fractional response variables with an application to 401(k) plan participation rates", *Journal of Applied Econometrics* 11(6), 619–632 | Fractional-response logit for monthly survival shares | ☐ |

<a id="verification-record"></a>

## ✅ Verification record

Verified on 2026-10-08 by the owner's research, using only the publishers' own
websites (bis.org, eur-lex.europa.eu). Page numbers are printed pages.

| Source | Provision | Verified content |
| --- | --- | --- |
| `REG-1` | ¶111, pp. 25–26; SRP31.108 | NMDs are segmented into retail and wholesale; categories *Retail/transactional*, *Retail/non-transactional*, *Wholesale*. Retail covers individuals, plus small businesses managed as retail with aggregate liabilities below EUR 1 million; transactional accounts have regular transactions or bear no interest; wholesale covers legal entities, sole proprietorships and partnerships |
| `REG-1` | ¶113, Table 2, p. 26; SRP31.110–31.112, Table 4 | Cap on the proportion of core deposits: 90 %, 70 %, 50 % |
| `REG-1` | ¶114 | Non-core deposits are placed in the overnight bucket |
| `REG-1` | ¶115, Table 2, p. 26; SRP31.111–31.112, Table 4 | Cap on the average maturity of core deposits: 5, 4.5 and 4 years |
| `REG-3` | Title, OJ header p. 1; Article 6, p. 5 | Number 2024/856, adopted 1 December 2023, OJ L 24.4.2024; in force on the twentieth day after publication (14 May 2024) |
| `REG-3` | Article 1(1)–(2), p. 2 | EVE scenarios: parallel up, parallel down, steepener, flattener, short rates up, short rates down; NII scenarios: parallel up, parallel down |
| `REG-3` | Article 3(7), p. 4; Article 4(1), p. 5 | Post-shock lower bound $\min(-150 + 3m,\ 0)$ basis points at maturity $m$ in years (rising by 3 basis points per year), replaced by the observed rate where lower; extended to NII by Article 4(1) |
| `REG-3` | Articles 1–6 and Annex, pp. 2–7 | **No** cap on the average repricing maturity of non-maturity deposits |

Not yet verified: `REG-2` in full (the official EBA texts could not be opened),
including the EU 5-year cap on NMD repricing maturity, its scope and the
guidelines' exact title and application dates; and the `REG-3` shock sizes per
currency.

The toolkit's own session could not open these sites (network policy), so the
record above rests on the owner's research and was not re-checked here.

## Verifying an entry

1. Obtain the official text from its publisher.
2. Record the exact title, version or publication date, and the article,
   paragraph or table relied on.
3. Check every figure or rule the contracts attribute to it.
4. Mark the entry verified in a reviewed change that names the reviewer, and
   correct the contracts where they differ.

A source that cannot be obtained stays unverified, and anything resting on it
stays a proposal.
