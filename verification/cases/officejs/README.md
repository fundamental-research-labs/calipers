# officejs (blank-init Office.js)

135 cases that start from a **blank XLSX** (byte copy of
[`../roundtrip/empty/init.xlsx`](../roundtrip/empty/init.xlsx)) and run one
`await Excel.run(...)` script. This is added coverage next to the 15
[`scratch/`](../scratch/) cases; those keep their committed goldens.
The original 100 cases have committed `excel-run` goldens; 35 newer
cases are pending Windows `excel-run-pass` (`verify` skips a missing
golden).

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
freeze, comments, hyperlinks, sort, and insert/delete.
