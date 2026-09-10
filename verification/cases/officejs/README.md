# officejs (blank-init Office.js)

About 100 cases that start from a **blank XLSX** (byte copy of
[`../roundtrip/empty/init.xlsx`](../roundtrip/empty/init.xlsx)) and run one
`await Excel.run(...)` script. This is added coverage next to the 15
[`scratch/`](../scratch/) cases; those keep their committed goldens.

There is **no** `golden.xlsx` in this suite yet. Generate them on
Windows + Excel:

```bat
calipers excel-run-pass verification\cases\officejs
```

Optional per-case `config.json` (`maxPeakMemoryBytes`, `maxDurationMs`) is
also Windows-only, after goldens:

```bat
calipers measure-budgets --engine excel --suite officejs
```

Regenerate the case dirs from the spec table (does not write goldens or configs):

```bash
go run ./scripts/gen-officejs-cases
```

Each case is one primary feature: tables, pivot tables, charts, spill
functions, styles, small and large grid writes, worksheets, names,
conditional formatting, validation, autofilter, comments, hyperlinks,
sort, and insert/delete.
