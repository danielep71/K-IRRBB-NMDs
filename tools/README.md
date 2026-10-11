# Repository tooling

The tools are adapted from the shared source-first conventions and portable source gates of VBA-SACCR-Toolkit; they do not perform IRRBB calculations.

```shell
python tools/check.py
python tools/check.py --ci
python tools/check.py --ci --base <base-sha>
```

The gate runner checks:
- Python-tool regression self-tests;
- committed / working-tree whitespace;
- source placement, LF storage and VBA export headers;
- data-contract fixtures in `tests/` agree with their hand-computed expectations and with `docs/methodology/DATA_CONTRACT.md` (`test_data_contract.py`);
- the VBA harness: every literal expectation in `tests/modules/TEST_CoreCases.bas` agrees with the contract and the harness rules, and the harness registry equals the register in `docs/methodology/TEST_CASES.md` (`test_vba_harness.py`; VBA is read, not run);
- the development-workbook builder: `vbaProject.bin` read back with an independent [MS-OVBA]/[MS-CFB] decoder (`test_devbuild.py`; the workbook itself only when `openpyxl` is installed);
- `src/core/` host independence: no Excel object-model or UI identifiers outside comments and strings;
- VBA jump targets and conditional compilation across supported Windows bitness contexts;
- explicit public API declarations compared with `docs/PUBLIC_API.txt`;
- changelog/version consistency.

CI runs these checks on GitHub's hosted Linux runner. **VBA compile, Excel workbook execution, econometric calibration and numerical result validation are not performed by these gates.** Record Excel host evidence separately under `docs/EXCEL_EVIDENCE.md`.

`devbuild/` builds the macro-enabled development workbook used for Excel compile and harness evidence; see [`devbuild/README.md`](devbuild/README.md).

Label reconciliation scripts and workflows mirror the SACCR label policy; merging them into the trusted default branch is needed before automatic label synchronization begins.

Ruff and mypy settings in `pyproject.toml` are optional development targets, not
CI checks or verified compliance. The required CI command remains the portable
standard-library gate runner above.
