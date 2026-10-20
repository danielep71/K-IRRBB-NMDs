# 🖼️ Documentation images

The images in this directory illustrate the IRRBB toolkit. They do not certify
any model or replace the [methodology](../methodology/README.md), which holds
the model contracts, the test cases and the regulatory sources.

<a id="social-preview"></a>

## 🪧 Social preview

`IRRBB_social_preview.html` is the source for `IRRBB_social_preview.png`, the
repository's social preview image and the picture at the top of the root
README (1280 x 640 px, under GitHub's 1 MB limit). It is a promotional image,
laid out like the social previews of VBA-SACCR-Toolkit and MKT-FTRB-SA.

- **Figures:** the decay card and the curve show the reference case
  `DEC-E2E-01` (`tests/expected/decay/DEC-E2E-01.json`, on `release/v0.3.0`
  and `main`): the synthetic retail transactional series `RET_TX`, estimated
  from 2012-01-31 to 2022-12-31, at 99% confidence, in EUR millions. Balance at
  t0 396.1 (`b0`), MPA after 12 and 120 months 364.3 and 304.1 (`mpa12`,
  `mpa120`), mean life 100.9 months (`mean_life`), non-core share 2.4%
  (`noncore_share`), and the core mean life after the REG-1 overlay, 59.5
  months (`ov_core_mean_life`), with the cutoff reduced to 66 months
  (`ov_cutoff`). The expected values come from numpy and scipy, independently
  of the VBA code, and have not yet been reviewed by the owner (#36). The decay
  model is a draft; the card says so.
- **Curve:** MPA(h) / B(t0) = exp(mu* h + z sigma sqrt(h)) for h = 0 to 120
  months, with the case's mu* = 0, sigma = 0.010374 and z = -2.326
  (`docs/methodology/MODEL_CONTRACTS.md`, decay profile), plotted from 1.00
  to 0.70; the dashed line marks the REG-1 cutoff of 66 months.
- **Caps:** the core-deposit caps of REG-1 (BCBS d368, Table 2; SRP31.110 to
  31.112): retail transactional 90% and 5 years, retail non-transactional 70%
  and 4.5 years, wholesale 50% and 4 years (`docs/methodology/SOURCES.md`).
- **Backtest:** the out-of-sample score of the same case, 2023-01 to 2025-12
  with frozen parameters: 1 exceedance in 354 (origin, horizon) pairs, a rate
  of 0.28% against the 1% tail probability (`bt_exceed`, `bt_pairs`,
  `bt_rate`).
- **Models:** the three model families of the root README; only decay exists,
  as a draft.
- **Name:** the title is K-IRRBB, in the naming of the sibling toolkits
  K-SACCR and K-FRTB-SA; the subject, IRRBB, is spelled out beneath it.
- **Status:** the pill names the version in progress, "In progress 0.2.0";
  update it with each release.

Regenerate by opening the HTML in a Chromium browser with a 1280 x 640 viewport
at device scale factor 1, after the Google Fonts (Barlow Condensed, Barlow,
JetBrains Mono) have loaded, and capturing the full page, for example:

```sh
chromium --headless --hide-scrollbars --force-device-scale-factor=1 \
  --window-size=1280,640 --virtual-time-budget=8000 \
  --screenshot=docs/assets/IRRBB_social_preview.png docs/assets/IRRBB_social_preview.html
```

Then upload the PNG in **Settings > General > Social preview**; GitHub has no
API for this setting.
