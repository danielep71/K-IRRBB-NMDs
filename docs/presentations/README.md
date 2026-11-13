# 🎓 Presentations

Teaching material on the subject of K-IRRBB. A presentation explains the rules
and the modeling practice; it does not certify any model or replace the
[methodology](../methodology/README.md), which holds the model contracts, the
test cases and the regulatory sources the toolkit follows.

<a id="modeling-non-maturity-deposits"></a>

## 🏦 Modeling non-maturity deposits

[`Modeling_Non-Maturity_Deposits.pptx`](Modeling_Non-Maturity_Deposits.pptx)
is a 70-slide lecture, with speaker notes, on how banks model non-maturity
deposits for interest rate risk in the banking book:

- the Basel standard, the CRD and CRR, the EU technical standards and the EBA
  guidelines, including the 5-year cap on the average repricing date;
- volume, rate and migration models, replicating portfolios, and the path from
  the model to EVE and NII;
- validation and governance, the EU standardized approach and the supervisory
  outlier tests;
- the evidence from the EBA's data and the 2023 bank failures;
- hedging and the EU carve-out from IAS 39.

The core path takes 90 minutes; the deep-dive slides, marked in the notes, add
about an hour. The last slides hold exercises, with solutions in the notes,
and the references.

- **Status:** the regulatory statements describe the texts as of November 2026,
  as the references slide says; the rules on shock sizes and hedge accounting
  were still moving then. The worked examples use illustrative figures, not
  the toolkit's test cases, and nothing in the lecture validates a model.
- **PDF:** [`Modeling_Non-Maturity_Deposits.pdf`](Modeling_Non-Maturity_Deposits.pdf)
  is the same deck for reading in the browser, without the notes. It is
  exported with LibreOffice and can differ slightly from PowerPoint.
- **Fonts:** the deck uses IBM Plex Sans and Source Serif 4, both under the SIL
  Open Font License. They are not embedded in the PPTX; without them installed,
  PowerPoint substitutes other fonts. The PDF embeds subsets of both.

Regenerate the PDF after changing the deck, with both fonts installed:

```sh
soffice --headless --convert-to pdf --outdir docs/presentations \
  docs/presentations/Modeling_Non-Maturity_Deposits.pptx
```
