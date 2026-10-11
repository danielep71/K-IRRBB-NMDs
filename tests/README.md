# Tests

Synthetic fixtures, reviewed expected values and the Excel regression harness live here. Group cases by model (Rates, Stable, Decay), transformation, validation and application recovery. Distinguish static checks from execution in real Excel.

`modules/` holds the VBA harness (#8). `TEST_Harness.RunTests` runs every case in the [register](../docs/methodology/TEST_CASES.md#register), in register order, and writes one row per case and one per assertion to the `Checks` sheet. A case is `PASS` only when it made at least one assertion and all passed; an unexpected error is `ERROR`; a registered case without a VBA implementation is `NOT RUN`, never `PASS`. Earlier results are cleared before a run starts; if writing the results fails, the status is `FAILED` with the primary and the cleanup error (`TEST_RunTestsWithWriteFault` forces that path, and `tools/devbuild/lo_smoke.py` runs it). `TEST_CoreCases` implements `HARN-OUT-01`, `DATA-PARSE-01`, `DATA-PARSE-02` and `DATA-CODES-01`; their expected values are literals that `tools/test_vba_harness.py` re-derives from the contract without Excel, and it keeps the harness registry equal to the register. Build the workbook to run it with `tools/devbuild/build_dev_workbook.py`.

Synthetic fixtures for the data contract are in `fixtures/data_contract/`, with hand-computed expectations in `expected/data_contract/`; `tools/test_data_contract.py` keeps them consistent with [`docs/methodology/DATA_CONTRACT.md`](../docs/methodology/DATA_CONTRACT.md).

No Excel test run is recorded yet, and no numerical model test passes are claimed. Do not commit actual customer records or proprietary benchmarks.
