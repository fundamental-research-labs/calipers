# Cache-free recalculation

This suite exercises **load XLSX without formula-result caches → full recalculation
→ save → compare with a Windows Excel golden**. It uses the same discovery,
semantic comparison, per-case budgets, and benchmarking as the other suites.

The inherited `config.json` requires `recalculate: true`. On Excel the host calls
[`Application.CalculateFullRebuild`](https://learn.microsoft.com/en-us/office/vba/api/excel.application.calculatefullrebuild)
between opening and saving. External engines receive `save --recalculate`.
The policy applies to `excel-save-pass`, `verify`, `measure-budgets`, and `bench`.
It is independent of the optional CLI `verify --recalculate` override.
Scripted cases cannot set this config field. A host lacking the operation errors.

Inputs are synthetic, deterministic OOXML with no formula `<v>`/`<is>` caches or
calculation chain. Manual calculation mode makes explicit recalculation necessary.
Unlike the roundtrip corpus, these inputs must **not** be rewritten by Excel;
that would populate their caches. There are no comparison exceptions, including
for volatile functions: `OFFSET` results must compare normally.

## Cases for [Mog #401](https://github.com/fundamental-research-labs/mog/issues/401)

All inputs have sheet `S`, period labels `FQ,FQ,FY,FQ,FY` in `A2:E2`, and this block
in `A10:E12`:

| Key | Q1 | Q2 | Q3 | Q4 |
| --- | -- | -- | -- | -- |
| k2 | 10 | 20 | 30 | 40 |
| k3 | 50 | 60 | 70 | 80 |

Names: `PeriodRange=S!$A$2:$E$2`, `SingleCellBase=S!$A$2`,
`PeriodRow=S!$2:$2`, `LookupBlock=S!$A$10:$E$12`, `DataStart=S!$B$11`.
`SingleCellBase` avoids the report's ambiguous `C1` identifier.

| Case | Result cells | Expected results | Purpose |
| --- | --- | --- | --- |
| offset_controls | H2:H8 | 3, 3, 4, 30, FY, 2, 140 | Literal OFFSET; independent VLOOKUP, INDEX, COUNTIF |
| offset_named_bases | H2:H6 | 3, 3, 3, 140, 130 | Bounded, single-cell, whole-row names; width, height, shifted origin |
| offset_index_bases | H2:H5 | 3, 3, 3, 140 | INDEX reference bases; omitted arguments; two-dimensional size |
| offset_period_counts | A20:D20 | 1, 2, 1, 3 | COUNTIF/OFFSET/INDEX/COLUMN composition |
| offset_lookup | H2:H4 | 4, 4, 30 | MATCH header slice and IFERROR/VLOOKUP composition |

Exact formulas and expected scalar results live in
[`scripts/gen-recalculate-cases`](../../../scripts/gen-recalculate-cases/main.go).
The expectations are fixture assertions, **not Excel goldens**.

Regenerate only the inputs (preserves budgets and goldens):

```sh
go run ./scripts/gen-recalculate-cases
go test ./scripts/gen-recalculate-cases ./internal/cases
```

## Windows handoff

Requires Go 1.25+, desktop Excel, and an engine binary implementing the Calipers
`save --recalculate <input> <output>` protocol. Close other Excel workbooks for
uncontaminated peak-memory measurements. No Office.js add-in is required.

From the Calipers checkout, with the Mog adapter built and `MOG_BIN` set:

```powershell
powershell -ExecutionPolicy Bypass -File scripts/capture-recalculate.ps1 -Engine C:\path\to\calipers-mog.exe
```

The script generates five goldens and metadata with `recalculate: true`, validates
all 23 formula results, measures peak memory and duration into case `config.json`
files, verifies Excel against the goldens, and verifies Mog against those same
files. It requires **5 Excel passes**, then **1 Mog pass / 4 semantic failures /
0 errors**, including a passing control case. It retains logs and exports in
`tmp/recalculate/`; missing goldens or skips cannot masquerade as success.
It regenerates the selected suite's goldens and budgets on reruns.

Commit the five goldens, their `.meta.json` sidecars, and five measured configs on
this branch. Share `tmp/recalculate/` when returning the evidence. Budgets remain
pending until measured on Windows, following the existing `measure-budgets` model;
no invented limits are committed. As with other suites, `verify` gates semantics;
`bench` records measurements, and `measure-budgets` writes limits.

Individual commands, from the repository root:

```powershell
go build -o calipers.exe ./cmd/calipers
.\calipers.exe excel-save-pass verification/cases/recalculate
go run ./scripts/gen-recalculate-cases -check-goldens
.\calipers.exe measure-budgets --engine excel --suite recalculate --force
.\calipers.exe verify --engine excel --suite recalculate
.\calipers.exe verify --engine C:\path\to\calipers-mog.exe --suite recalculate
```

The last command is expected to exit 1 until Mog is fixed. `excel-save` alone does
not resolve case configs; use the suite-aware `excel-save-pass` above.
