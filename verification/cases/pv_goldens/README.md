# PV golden recalculation round trips

107 cases exercise **load a golden with formula-result caches removed → full
recalculation → save cached results → compare against the supplied golden**.
They use the existing Calipers case loader, semantic comparator, budget capture,
and benchmark runner. No Office.js script or add-in is needed.

The suite's `config.json` inherits `"recalculate": true` into every case. Excel
uses COM `CalculateFullRebuild`; an external engine receives
`save --recalculate <init.xlsx> <output.xlsx>`. No memory or duration limits are
invented here: individual case configs await measurement on Windows.

## Fixtures and provenance

The [source fixtures](https://github.com/lyfegame/shortcut/tree/e351394e607b7b8b9e8791893abc052a0ead89f3/packages/spreadsheet-verification/tests/fixtures/pv_goldens)
contain 66 `pv23`, 20 `pv26`, and 21 `pv26_2` goldens. Cases retain the source IDs
(e.g. `pv_goldens/pv23_001`). Only the source `golden.xlsx` files are imported;
the source `input.xlsx` files are never used.

Each committed `golden.xlsx` is byte-identical to its source Git LFS object.
[`manifest.json`](manifest.json) pins the repository, commit, relative source
paths, and SHA-256 hashes (the source LFS object IDs). The upstream fixture README
describes anonymization of identifying text and metadata; these files are used
as supplied. We do not invent Excel version/build or capture sidecars for them.
These imported references are an exception to the usual Calipers workflow of
creating goldens locally through Windows Excel.

[`scripts/gen-pv-golden-inits`](../../../scripts/gen-pv-golden-inits/main.go)
derives each `init.xlsx` solely from its committed golden. It deletes worksheet
cell `<v>` and `<is>` elements for formula cells, including shared formulas, and
for every cell within an array/spill or data-table formula's `ref` range. This
also clears cached followers that have no `<f>` of their own. Formulas, cell
attributes, constants, formatting, workbook calculation settings, calculation
chains, and every other ZIP part remain unchanged. External-link, chart, and
pivot caches are not worksheet formula results and remain intact. Explicit full
recalculation is requested by the suite config, without rewriting the workbook
settings or saving the init through Excel.

From the Calipers repository root, regenerate or check without Excel/network:

```sh
go run ./scripts/gen-pv-golden-inits
go run ./scripts/gen-pv-golden-inits -check
go test ./scripts/gen-pv-golden-inits ./internal/cases
```

The generator touches only `init.xlsx`; it preserves goldens, the manifest,
and measured configs. `-check` verifies all source hashes and compares every ZIP
payload in each init against the cache-stripped golden, independent of compressor
version. Tests cover scalar, shared, array/spill, and data-table caches; literal
preservation; all 107 discovered cases; and the actual semantic reader.

## Windows handoff: capture budgets, check Excel

Requires Go 1.25+ and desktop Microsoft Excel on Windows. Close other Excel
workbooks for accurate memory measurements. From this Calipers branch:

```powershell
powershell -ExecutionPolicy Bypass -File scripts/capture-pv-goldens.ps1
```

This builds Calipers, checks fixture integrity, measures all 107 cases with Excel
into per-case `config.json` files, and verifies Excel exports against the supplied
goldens. It checks integrity again afterward. Reruns remeasure all budgets using
`--force`. Logs and recalculated workbooks are retained under `tmp/pv_goldens/`.
No Mog binary is needed for this step.

Equivalent individual commands:

```powershell
go build -o calipers.exe ./cmd/calipers
go run ./scripts/gen-pv-golden-inits -check
.\calipers.exe measure-budgets --engine excel --suite pv_goldens --force
.\calipers.exe verify --engine excel --suite pv_goldens --out-dir tmp/pv_goldens/excel
```

Expect `measure-budgets: 107 wrote, 0 failed, 0 skipped` and, if Excel reproduces
every supplied reference, `verify: 107 pass, 0 fail, 0 error`. The supplied
workbooks have not yet been checked with Excel by this change. Anonymization,
volatile or external data, and Excel-version differences may cause discrepancies;
there are no comparison exceptions or disabled cases hiding them. If comparison
fails, keep the references and share the verification log and affected exports.
Budget capture still completes independently of semantic mismatches.

**Do not run `excel-save-pass` on this suite** (or an unfiltered pass over the
whole corpus): it regenerates/overwrites goldens. These references must remain
unchanged. Never save `init.xlsx` through Excel either: that would repopulate its
caches. Use the integrity check after the Windows run.

Commit and push the 107 measured configs on this branch:

```powershell
git add -- 'verification/cases/pv_goldens/*/config.json'
git commit -m "Capture Excel budgets for PV golden round trips"
git push
```

Keep `tmp/pv_goldens/` separately as the evidence bundle. Later, after those
configs are pushed, run Mog from the planned new Mog branch:

```sh
calipers verify --engine /path/to/calipers-mog --suite pv_goldens
```

The engine must implement `save --recalculate`. Calipers compares workbook
semantics (cached values/types, formulas, styles, sheets, names, merges, freeze,
and `date1904`), not ZIP byte equality. The existing `verify` command gates
semantics; `measure-budgets` writes limits and `bench` records performance.
