# Excel-golden verification

Goal: keep a spreadsheet engine’s workbook behavior (XLSX I/O now, Office.js later) aligned with **desktop Microsoft Excel**. Oracle is Excel itself, not LibreOffice or another library.

This repo ships the golden generator (`calipers excel-save` / `excel-run`), the case corpus, and `calipers verify` (engine load → optional Office.js → export → semantic comparison against the golden).

## First test: open + save

1. Open an `.xlsx`, save it again — once in **Excel**, once in the engine under test.
2. Compare the two outputs (not the original vs the engine). Excel rewrite is the golden.

```
calipers excel-save <init.xlsx> <golden.xlsx>   # Windows + Excel COM only
# example: verification/cases/roundtrip/simple/init.xlsx
```

Goldens are generated on **Windows Excel via COM**. Do not generate goldens on Mac — Excel for Mac serializes OOXML differently. CI never runs Excel; it compares engine output to committed goldens.

`excel-save` also writes `<golden.xlsx>.meta.json` (`host=excel-win`, Excel version/build, OS, tool name, input basename, optional `script` identity, UTC time). Load+save goldens omit `script`. Refuse to mix `excel-win` and `excel-mac` goldens.

## What to compare

Excel-saved vs engine-saved is **never** byte-identical (timestamps, `calcId`, style indexes, relationship ids, extra Excel parts).

| Layer | Meaning | Current status |
|-------|---------|------|
| **Semantic** | Resolved cell values, types, formulas, styles, sheets, names, merges, freeze, `date1904` | Current `verify` PASS/FAIL |
| **Package** | ZIP-part contents after selected metadata exclusions and XML line-ending normalization | Optional `verify --package` diagnostic (does not change PASS/FAIL) |

FAIL lines are semantic diffs (`values: Sheet1!C1 expected 15 got 0`), not ZIP-part counts. A failing walk also prints a `Differences:` recap listing every FAIL/ERROR case. XML attribute/element ordering, numeric encodings of the same binary64 value, shared-string indexes, relationship IDs, style IDs, theme display names, Excel ST_Percentage tint text (`0.2` vs `0.19998779259620961`), empty border sides, and default row/col layout are not the gate. Font name and size are the gate. `verify --package` still reports those as differing ZIP parts.

A package match does not prove that formulas were evaluated. By default, `verify` uses the host's existing `save`/`run` behavior. Opt-in `verify --engine PATH --recalculate` requests full recalculation before export by passing `save --recalculate <in> <out>` or `run --recalculate <in> <script> <out>`. The selected external engine must support this flag. The selected policy is printed before verification. Engines retain their original argv when the flag is absent. `--engine` is required (binary path or `excel`).

`--recalculate` is rejected with the Excel host, which opens and saves without explicitly controlling calculation. Neither policy controls time, randomness, or environment-dependent results. Volatile formulas and their transitive dependents still require controlled assertions; fresh recalculation alone cannot match their previously captured golden caches. Goldens remain unchanged by `verify`.

Always ignore: ZIP mtimes, `docProps` creator/dates, `Application`/`AppVersion`, `workbookPr@calcId`, `xl/calcChain.xml`, printer settings.

Never ignore: `date1904`, sheet order, values, types, formulas, styles, merges, names, freeze.

## Hosts (reuse, don’t fork)

```
IWorkbookHost.OpenSave(in, out)     # excel-save: Excel-on-Windows COM
IWorkbookHost.RunScript(in, js, out)  # excel-run: same Office.js in Excel via sideloaded add-in
```

Excel does **not** execute Office.js through COM. `excel-run` sideloads a local add-in (WebView2) that `Excel.run`s the corpus script. AppSource install is not required. Cases with a missing or empty `script.js` skip script execution (load+save / `excel-save` only).

## Cases (`cases/`)

Cases are grouped into **suites** (test categories): `cases/<suite>/<case>/` with required `init.xlsx`, optional `script.js`, optional `config.json` (peak-memory / duration budgets), and a dedicated `golden.xlsx` destination (not mixed into the inits). Case directories are feature names. Missing or empty `script.js` means load+save only. Missing `config.json` is valid (no budget). **203 cases** (83 roundtrip, 5 default, 15 scratch, 100 officejs). Eight hostiles live in `_disabled/` and are not loaded. Six XLSX roundtrip/lost-info cases (custom views, theme colors, names.add order, sparse sheet IDs, chart titles, multi-row formulas) have committed Excel-win goldens. `verify` skips a case that has no `golden.xlsx`.

| Suite | Role |
|-------|------|
| `roundtrip/` | Load+save semantic comparison (Excel rewrite vs engine export). |
| `default/` | Untriaged until categorized (Office.js `simple_set_a1`). |
| `scratch/` | Office.js from an empty init (one feature per script). Committed `excel-run` goldens. |
| `officejs/` | ~100 blank-init Office.js cases (copy of `roundtrip/empty`). Committed `excel-run` goldens. |
| `_disabled/` | Hostile / later cases. Kept on disk; not a suite (names starting with `_` are skipped). |

Committed Office.js goldens: `default/simple_set_a1` (copy of `roundtrip/simple` init; Office.js sets A1 to `calipers`; golden from `excel-run`), the 15 `scratch/` cases (copy of `roundtrip/empty`; each runs one `Excel.run` feature), and `officejs/` (the larger blank-init set: tables, pivots, charts, spill, styles, small/large writes, …). `officejs/` also has committed `config.json` budgets from Windows `measure-budgets`. `roundtrip/simple` stays load+save. Default-pass load+save goldens (`golden.xlsx` + `.meta.json`, `host=excel-win`) are committed next to each unscripted case. Office.js goldens record `script=script.js`; `excel-save-pass` still skips scripted cases (including scratch and officejs). Hostiles in `_disabled/` are not in that set. Provenance: [`cases/README.md`](cases/README.md).

`calipers verify` walks committed-golden cases (roundtrip, default, scratch, and officejs). `--suite scratch` runs only the empty-init Office.js cases; `--suite officejs` runs the larger blank-init set; `--suite roundtrip` runs one directory; `--case roundtrip/simple` runs one test. Case ids are `suite/name`.

Sources (do not vendor FUSE/SpreadsheetBench — 16k unlabeled real-world files):

- Authored fixtures in this corpus (empty/simple layout, types, styles, unicode, multi-sheet, formula categories)
- [SheetJS test_files](http://oss.sheetjs.com/test_files/) (Apache 2.0)
- [Apache POI test-data/spreadsheet](https://github.com/apache/poi/tree/trunk/test-data/spreadsheet) (Apache 2.0)

Skip `.xlsm` for v1: COM `SaveAs` format 51 writes non-macro xlsx. `_disabled/password` is OLE-encrypted (not a zip); `_disabled/xmlbomb` is a parser stress file. Hostiles in `_disabled/` are not loaded.

## Later tests

- Same Office.js script in the engine and Excel → compare workbooks (`excel-run` is the Excel side).
- Formula eval vs Excel cached results.
- Live dual-host runs (Excel + engine, no golden) on a Windows box.

## Commands

```bash
# Build
go build -o calipers ./cmd/calipers

# Windows + Excel
./calipers excel-save verification/cases/roundtrip/simple/init.xlsx golden.xlsx
./calipers excel-save-pass
./calipers excel-run verification/cases/default/simple_set_a1/init.xlsx \
  verification/cases/default/simple_set_a1/script.js golden.xlsx

# Off Windows (expected)
./calipers excel-save …   # errors: requires Windows + Excel (COM)
./calipers excel-run …    # same error class
./calipers version

# Engine vs committed goldens
./calipers verify --engine /path/to/engine
./calipers verify --engine /path/to/engine --suite roundtrip
./calipers verify --engine /path/to/engine --suite scratch
./calipers verify --engine /path/to/engine --case roundtrip/simple
./calipers verify --engine excel --case roundtrip/simple
```
