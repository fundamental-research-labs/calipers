# Excel-golden verification

Goal: keep a spreadsheet engine’s workbook behavior (XLSX I/O now, Office.js later) aligned with **desktop Microsoft Excel**. Oracle is Excel itself, not LibreOffice or another library.

This repo ships the golden generator (`calipers excel-save` / `excel-run`), the case corpus, and `calipers verify` (engine load → optional Office.js → export → package comparison against the golden).

## First test: open + save

1. Open an `.xlsx`, save it again — once in **Excel**, once in the engine under test.
2. Compare the two outputs (not the original vs the engine). Excel rewrite is the golden.

```
calipers excel-save <init.xlsx> <golden.xlsx>   # Windows + Excel COM only
# example: verification/cases/roundtrip/tier_a_simple/init.xlsx
```

Goldens are generated on **Windows Excel via COM**. Do not generate goldens on Mac — Excel for Mac serializes OOXML differently. CI never runs Excel; it compares engine output to committed goldens.

`excel-save` also writes `<golden.xlsx>.meta.json` (`host=excel-win`, Excel version/build, OS, tool name, input basename, optional `script` identity, UTC time). Load+save goldens omit `script`. Refuse to mix `excel-win` and `excel-mac` goldens.

## What to compare

Excel-saved vs engine-saved is **never** byte-identical (timestamps, `calcId`, style indexes, relationship ids, extra Excel parts).

| Layer | Meaning | Current status |
|-------|---------|------|
| **Semantic** | Resolved cell values, types, formulas, styles, sheets, names, merges, freeze, `date1904` | Not implemented |
| **Package** | ZIP-part contents after selected metadata exclusions and XML line-ending normalization | Current `verify` PASS/FAIL |

Each difference is one ZIP part, ordered by part name; the first part is not a severity ranking. XML attribute/element ordering, numeric encodings, shared strings, relationship IDs, style IDs, and shared formulas are not resolved or canonicalized. These can produce failures for equivalent workbook content. All character data inside XML is retained, including whitespace-only text and indentation: without content-model information, removing whitespace between tags can erase actual spreadsheet strings.

A package match does not prove that formulas were evaluated. By default, `verify` uses the host's existing `save`/`run` behavior; Mog preserves imported caches. Opt-in `verify --recalculate` requests full recalculation before export by passing `save --recalculate <in> <out>` or `run --recalculate <in> <script> <out>`. The selected external engine must support this flag; Mog recalculates after script execution for `run`. The selected policy is printed before verification. Generic engines retain their original argv when the flag is absent.

`--recalculate` is rejected with the Excel host, which opens and saves without explicitly controlling calculation. Neither policy controls time, randomness, or environment-dependent results. Volatile formulas and their transitive dependents still require controlled assertions; fresh recalculation alone cannot match their previously captured golden caches. Goldens remain unchanged by `verify`.

Always ignore: ZIP mtimes, `docProps` creator/dates, `Application`/`AppVersion`, `workbookPr@calcId`, `xl/calcChain.xml`, printer settings.

Never ignore: `date1904`, sheet order, values, formulas, styles, merges, names.

## Hosts (reuse, don’t fork)

```
IWorkbookHost.OpenSave(in, out)     # excel-save: Excel-on-Windows COM
IWorkbookHost.RunScript(in, js, out)  # excel-run: same Office.js in Excel via sideloaded add-in
```

Excel does **not** execute Office.js through COM. `excel-run` sideloads a local add-in (WebView2) that `Excel.run`s the corpus script. AppSource install is not required. Cases with a missing or empty `script.js` skip script execution (load+save / `excel-save` only).

## Cases (`cases/`)

Cases are grouped into **suites** (test categories): `cases/<suite>/tier_{a|b|c}_<feature>/` with required `init.xlsx`, optional `script.js`, and a dedicated `golden.xlsx` destination (not mixed into the inits). Missing or empty `script.js` means load+save only. **90 cases** (`tier_a_` 57, `tier_b_` 25, `tier_c_` 8).

| Suite | Role |
|-------|------|
| `roundtrip/` | Load+save package comparison (Excel rewrite vs engine export). |
| `default/` | Untriaged until categorized (Office.js `tier_a_simple_set_a1`, `tier_c` hostiles). |

One scripted case so far: `default/tier_a_simple_set_a1` (copy of `roundtrip/tier_a_simple` init; Office.js sets A1 to `calipers`; golden from `excel-run`). `roundtrip/tier_a_simple` stays load+save. Default-pass load+save goldens (`golden.xlsx` + `.meta.json`, `host=excel-win`) are committed next to each unscripted `tier_a`/`tier_b` case. The Office.js golden is also committed (`script=script.js`); `excel-save-pass` still skips it. Not in that set: `tier_c` hostiles. Provenance: [`cases/README.md`](cases/README.md).

`calipers verify` walks every suite. `--suite roundtrip` runs one directory; `--case roundtrip/tier_a_simple` runs one test. Case ids are `suite/name`.

| Prefix | Role |
|--------|------|
| `tier_a_` | Engine claims support (values, formulas, styles, sheets, merges, names, freeze, unicode, themes, number formats). First goldens. |
| `tier_b_` | Known remaining work (charts, pivot, tables/autofilter, CF, validation, comments, drawings, hyperlinks, protection, print). Goldens still useful: Excel keeps these, an engine may drop them. |
| `tier_c_` | Later / hostile (strict OOXML, password, XML bomb, huge stress, ATP, OLE embed). **Do not** run in the default golden pass. |

Sources (do not vendor FUSE/SpreadsheetBench — 16k unlabeled real-world files):

- Authored fixtures in this corpus (empty/simple layout, types, styles, unicode, multi-sheet, formula categories)
- [SheetJS test_files](http://oss.sheetjs.com/test_files/) (Apache 2.0)
- [Apache POI test-data/spreadsheet](https://github.com/apache/poi/tree/trunk/test-data/spreadsheet) (Apache 2.0)

Skip `.xlsm` for v1: COM `SaveAs` format 51 writes non-macro xlsx. `tier_c_password` is OLE-encrypted (not a zip); `tier_c_xmlbomb` is a parser stress file. A default golden pass skips `tier_c` hostiles.

## Later tests

- Same Office.js script in the engine and Excel → compare workbooks (`excel-run` is the Excel side).
- Formula eval vs Excel cached results.
- Live dual-host runs (Excel + engine, no golden) on a Windows box.

## Commands

```bash
# Build
go build -o calipers ./cmd/calipers

# Windows + Excel
./calipers excel-save verification/cases/roundtrip/tier_a_simple/init.xlsx golden.xlsx
./calipers excel-save-pass
./calipers excel-run verification/cases/default/tier_a_simple_set_a1/init.xlsx \
  verification/cases/default/tier_a_simple_set_a1/script.js golden.xlsx

# Off Windows (expected)
./calipers excel-save …   # errors: requires Windows + Excel (COM)
./calipers excel-run …    # same error class
./calipers version

# Engine vs committed goldens
./calipers verify --engine /path/to/engine
./calipers verify --engine /path/to/engine --suite roundtrip
./calipers verify --engine /path/to/engine --case roundtrip/tier_a_simple
./calipers verify --engine excel --case roundtrip/tier_a_simple
```
