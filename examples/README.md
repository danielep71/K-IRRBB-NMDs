# Examples

Only small, synthetic and redistributable examples.

- `modules/EX_DecayExample.bas` — fits the decay model, applies the `REG-1` overlay and scores the test window through `IRRBB_Decay` only, on a synthetic series read from a sheet named `DecayExample`. The template-based development workbook does not have that sheet yet (#36). Because the sheet is outside the [workbook contract](../docs/WORKBOOK_CONTRACT.md#sheets), the example addresses it by tab name and literal cells, the contract's one stated exception; it never touches a contract sheet or name.

Each future example will identify its model version, input conventions, expected outputs and how to reproduce them using supported APIs.
