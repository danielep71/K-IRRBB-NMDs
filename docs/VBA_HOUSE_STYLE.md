# VBA house style

Adapted to K-IRRBB from the maintainer's SACCR source-first conventions.

- Use `Option Explicit` on every module; `Option Private Module` in internal `src/core/` standard modules.
- Core code never uses the Excel object model or UI. Host names such as `Range`, `Cells`, `Application` and `MsgBox` are reserved in `src/core/`, even as variable names; see [`REPOSITORY_STRUCTURE.md`](REPOSITORY_STRUCTURE.md#dependency-direction).
- Pass rates as decimals, dates as `Date`, horizons in months; convert other units only in workbook adapters ([parameter and units boundary](REPOSITORY_STRUCTURE.md#parameter-and-units-boundary)).
- Match exported file stems to `Attribute VB_Name`. Export to Git with LF stored and CRLF checkout under `.gitattributes`.
- Place pure numeric/statistical computations in `CORE_*` modules, supported public entry points in `IRRBB_*`, tests in `TEST_*`.
- Public functions document contract, units, allowed input domain, missing-data semantics, errors and numerical tolerance.
- Use explicit parameter directions, `LongPtr` / `PtrSafe` for supported 64-bit declarations, no silent `Variant` return for calculation failures.
- Calculation procedures must be deterministic. They must not depend on workbook selection, active sheet, locale-specific string parsing or external mutable Excel state.
- Separate primary errors from cleanup errors. On failure, invalidate results rather than leaving stale workbook outputs apparently usable.
- Array lower bounds, row ordering, observation dates, scenario shock sign, basis points versus decimal rates and confidence quantiles are part of the contract.
- Public API changes require a versioned manifest update and affected tests.
- No pasted third-party proprietary source code or actual customer datasets.
