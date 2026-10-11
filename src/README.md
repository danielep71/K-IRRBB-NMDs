# Production VBA source

Source-first application profile, adapted from VBA-SACCR-Toolkit.

- `core/`: host-independent numerical, econometric and validation functions (`CORE_*`, `Option Private Module`). `tools/check_source.py` rejects Excel object-model and UI identifiers here.
- `modules/`: explicitly supported public facade (`IRRBB_*`).
- `workbook/`: host adapters and exported Excel document modules.
- `classes/`, `forms/`: introduce only when needed.

`core/` holds the first modules: `CORE_Parse` (locale-independent dates, decimals and rates of the data contract) and `CORE_Codes` (segment codes and the currency allowlist), plus the draft decay model `CORE_Math`, `CORE_Decay`, `CORE_Overlay` with the facade `modules/IRRBB_Decay`, under review in #36 and without Excel evidence. Layout, dependency direction and the parameter and units boundary are defined in [`docs/REPOSITORY_STRUCTURE.md`](../docs/REPOSITORY_STRUCTURE.md).
