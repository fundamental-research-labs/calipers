# officejs (blank-init Office.js)

243 cases that start from a **blank XLSX** (byte copy of
[`../roundtrip/empty/init.xlsx`](../roundtrip/empty/init.xlsx)) and run one
`await Excel.run(...)` script. This is added coverage next to the 15
[`scratch/`](../scratch/) cases; those keep their committed goldens.
All 243 cases have committed `excel-run` goldens and `config.json`
budgets (`maxPeakMemoryBytes`, `maxDurationMs`).

Committed `golden.xlsx` files come from Windows + Excel
(`excel-run-pass`). To regenerate:

```bat
calipers excel-run-pass verification\cases\officejs
```

Committed `config.json` budgets (`maxPeakMemoryBytes`, `maxDurationMs`)
come from Windows `measure-budgets` (measured Excel peak working set and
wall time, times a 1.5 margin). To regenerate:

```bat
calipers measure-budgets --engine excel --suite officejs --force
```

Regenerate the case dirs from the spec table (does not write goldens or configs):

```bash
go run ./scripts/gen-officejs-cases
```

Each case is one primary feature: tables, pivot tables, charts, spill
functions, styles, small and large grid writes, range navigators,
worksheets, names, conditional formatting, validation, autofilter,
freeze, comments, hyperlinks, sort, insert/delete, and leftover
collection lookups (`getCount`, `getItem` / `getItemAt` / `get*OrNullObject`,
`getRange`).

The 50 `api_function_*` cases exercise `workbook.functions` arithmetic,
trigonometry, aggregation, and text methods. Each writes primary and boundary
results (including worksheet errors) into `D1:F3`. Their source is
`scripts/gen-officejs-cases/specs_functions.go`. Goldens come from
`excel-run-pass`; budgets come from `measure-budgets`.
