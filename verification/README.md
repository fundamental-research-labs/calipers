# Excel-golden verification

Goal: keep a spreadsheet engine’s workbook behavior (XLSX I/O now, Office.js later) aligned with **desktop Microsoft Excel**. Oracle is Excel itself, not LibreOffice or another library.

This repo ships the golden generator (`calipers excel-save`) and the init corpus. Engine-vs-Excel semantic/package compare is a later step.

## First test: open + save

1. Open an `.xlsx`, save it again — once in **Excel**, once in the engine under test.
2. Compare the two outputs (not the original vs the engine). Excel rewrite is the golden.

```
calipers excel-save <init.xlsx> <golden.xlsx>   # Windows + Excel COM only
```

Goldens are generated on **Windows Excel via COM**. Do not generate goldens on Mac — Excel for Mac serializes OOXML differently. CI never runs Excel; it compares engine output to committed goldens.

`excel-save` also writes `<golden.xlsx>.meta.json` (`host=excel-win`, Excel version/build, OS, tool name, input basename, UTC time). Refuse to mix `excel-win` and `excel-mac` goldens.

## What to compare

Excel-saved vs engine-saved is **never** byte-identical (timestamps, `calcId`, style indexes, relationship ids, extra Excel parts).

| Layer | Meaning | Gate |
|-------|---------|------|
| **Semantic** | Cell values, types, formulas, styles, sheets, names, merges, freeze, `date1904` | Pass/fail once a case is green |
| **Package** | Canonical OOXML after stripping volatile bits | Report now; tighten later |

Always ignore: ZIP mtimes, `docProps` creator/dates, `Application`/`AppVersion`, `workbookPr@calcId`, `xl/calcChain.xml`, printer settings.

Never ignore: `date1904`, sheet order, values, formulas, styles, merges, names.

## Hosts (reuse, don’t fork)

```
IWorkbookHost.OpenSave(in, out)     # v1: Excel-on-Windows; later the engine CLI
IWorkbookHost.RunScript(in, js, out)  # later: same Office.js in both hosts
```

Excel does **not** execute Office.js through COM. A later sidecar add-in is required to run the same script inside Excel. Until then, Excel `RunScript` is unsupported.

## Init corpus (`init_files/`)

Inputs only — goldens are produced on Windows, not stored here yet. **91 files** (`tier_a_` 57, `tier_b_` 26, `tier_c_` 8).

| Prefix | Role |
|--------|------|
| `tier_a_` | Engine claims support (values, formulas, styles, sheets, merges, names, freeze, unicode, themes, number formats). First goldens. |
| `tier_b_` | Known remaining work (charts, pivot, tables/autofilter, CF, validation, comments, drawings, hyperlinks, protection, print). Goldens still useful: Excel keeps these, an engine may drop them. |
| `tier_c_` | Later / hostile (strict OOXML, password, XML bomb, huge stress, ATP, OLE embed). **Do not** run in the default golden pass. |

Sources (do not vendor FUSE/SpreadsheetBench — 16k unlabeled real-world files):

- Authored fixtures in this corpus (empty/simple layout, types, styles, unicode, multi-sheet, formula categories)
- [SheetJS test_files](http://oss.sheetjs.com/test_files/) (Apache 2.0)
- [Apache POI test-data/spreadsheet](https://github.com/apache/poi/tree/trunk/test-data/spreadsheet) (Apache 2.0)

Skip `.xlsm` for v1: COM `SaveAs` format 51 writes non-macro xlsx. `tier_c_password.xlsx` is OLE-encrypted (not a zip); `tier_c_xmlbomb.xlsx` is a parser stress file. A default golden pass skips those hostile files.

## Later tests

- Same Office.js script in the engine and Excel → compare workbooks (needs Excel add-in).
- Formula eval vs Excel cached results.
- Live dual-host runs (Excel + engine, no golden) on a Windows box.

## Commands

```bash
# Build
go build -o calipers ./cmd/calipers

# Windows + Excel
./calipers excel-save verification/init_files/tier_a_simple.xlsx golden.xlsx

# Off Windows (expected)
./calipers excel-save …   # errors: requires Windows + Excel (COM)
./calipers version
```
