# Generating goldens for new cases

Goldens must be produced on **Windows + licensed desktop Excel**. Do not generate them on macOS.

Build the tool from the repo root:

```bat
go build -o calipers.exe ./cmd/calipers
```

## Load+save cases (`excel-save`)

`roundtrip/custom_view_printer_settings` covers mog #360. That engine fix lives in mog PR #362 (opened first); this calipers case is for verifying that PR, not a parallel mog patch.

These have `init.xlsx` and no `script.js`:

```bat
calipers excel-save verification\cases\roundtrip\custom_view_printer_settings\init.xlsx verification\cases\roundtrip\custom_view_printer_settings\golden.xlsx
calipers excel-save verification\cases\roundtrip\theme_linked_colors\init.xlsx verification\cases\roundtrip\theme_linked_colors\golden.xlsx
```

Or regenerate every unscripted case that still needs a golden:

```bat
calipers excel-save-pass
```

`excel-save-pass` skips Office.js cases.

## Office.js cases (`excel-run`)

These have a non-empty `script.js`:

```bat
calipers excel-run verification\cases\default\names_add_defined_names_order\init.xlsx verification\cases\default\names_add_defined_names_order\script.js verification\cases\default\names_add_defined_names_order\golden.xlsx
calipers excel-run verification\cases\default\add_sheet_sparse_ids\init.xlsx verification\cases\default\add_sheet_sparse_ids\script.js verification\cases\default\add_sheet_sparse_ids\golden.xlsx
calipers excel-run verification\cases\default\chart_titles\init.xlsx verification\cases\default\chart_titles\script.js verification\cases\default\chart_titles\golden.xlsx
calipers excel-run verification\cases\default\multi_row_formulas\init.xlsx verification\cases\default\multi_row_formulas\script.js verification\cases\default\multi_row_formulas\golden.xlsx
```

Each successful run also writes `golden.xlsx.meta.json` beside the golden.

`officejs/` cases start from `roundtrip/empty` and have committed
`excel-run` goldens (`host=excel-win`, `script=script.js`). On Windows,
regenerate every pending scripted golden:

```bat
calipers excel-run-pass
calipers excel-run-pass verification\cases\officejs
```

`excel-run-pass` skips load+save cases and cases that already have a golden.
Off Windows it errors once there is work to do (same as `excel-run`).

## Optional peak-memory / duration budgets

A case directory may include `config.json` next to `init.xlsx` / `script.js`
/ `golden.xlsx`:

```json
{
  "maxPeakMemoryBytes": 268435456,
  "maxDurationMs": 8000
}
```

Missing `config.json` is valid (no budget). Do not invent numbers: measure
on Windows (Excel and/or an engine binary) and write the sidecar with a
margin:

```bat
calipers measure-budgets --engine excel --suite officejs
calipers measure-budgets --engine excel --margin 1.5 --force
```

`--engine PATH` uses the same `save` / `run` argv as `verify`. This helper
is not run in Linux CI.

## Verify with committed goldens

Goldens for the six cases above plus roundtrip/default/scratch/officejs are committed (`host=excel-win`). `calipers verify --engine <bin>` skips any case that has no `golden.xlsx`.

PASS/FAIL is **semantic** compare (values, types, formulas, styles, sheets, names, merges, freeze, `date1904`). A failing run prints each case’s diffs and a `Differences:` recap. Charts, printer settings, `sheetId`, and workbook XML child order are **not** in that gate. For those, also run:

```bat
calipers verify --engine <mog-bin> --package --case roundtrip\custom_view_printer_settings
calipers verify --engine <mog-bin> --package --case default\chart_titles
calipers verify --engine <mog-bin> --package --case default\names_add_defined_names_order
calipers verify --engine <mog-bin> --package --case default\add_sheet_sparse_ids
```

`--package` is diagnostic only and does not change PASS/FAIL. An engine export crash is still `ERROR` (that is how #360 fails today without a printer-settings compare).
