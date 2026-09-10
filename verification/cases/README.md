# Verification cases

Cases live in a **suite** directory (`roundtrip/`, `default/`, `scratch/`, or `officejs/`) as feature-named directories. Directories whose names start with `_` are not suites.

| Suite | Role |
|-------|------|
| [`roundtrip/`](roundtrip/) | Load+save semantic comparison |
| [`default/`](default/) | Untriaged (Office.js) |
| [`scratch/`](scratch/) | Office.js from an empty init (one feature per script; committed `excel-run` goldens) |
| [`officejs/`](officejs/) | ~100 blank-init Office.js cases (copy of `roundtrip/empty`; **no goldens yet** — Windows `excel-run-pass`) |
| [`_disabled/`](_disabled/) | Hostile / later cases (not loaded) |

Committed Excel-win goldens for XLSX roundtrip/lost-info cases (regenerate with [`GOLDENS.md`](../GOLDENS.md)):

| Case | Issue | Semantic gate | Golden command |
|------|-------|---------------|----------------|
| `roundtrip/custom_view_printer_settings/` | mog #360 custom-view pageSetup / printer rels | export crash = ERROR; print parts not compared | `excel-save` |
| `roundtrip/theme_linked_colors/` | mog #329 theme+tint colors | styles (theme slots) | `excel-save` |
| `default/names_add_defined_names_order/` | mog #332 `<definedNames>` before `<calcPr>`/`<extLst>` | name exists, not XML order — use `--package` | `excel-run` |
| `default/add_sheet_sparse_ids/` | mog #334 unique `sheetId` after add | sheet names, not `sheetId` — use `--package` | `excel-run` |
| `default/chart_titles/` | mog #338 `<c:overlay val="0"/>` on authored titles | charts out of v1 — use `--package` | `excel-run` |
| `default/multi_row_formulas/` | mog #328 formula strings on later rows | formulas | `excel-run` |

Each case directory:

| File | Role |
|------|------|
| `init.xlsx` | Required input workbook |
| `script.js` | Optional Office.js. Missing or empty → load+save only (skip script execution) |
| `golden.xlsx` | Excel-win oracle. Committed for the default open+save pass (unscripted cases), `simple_set_a1` (`excel-run`), and `scratch/` (`excel-run`). **Not** present for `officejs/` until Windows `excel-run-pass`. |
| `golden.xlsx.meta.json` | Sidecar (`host=excel-win`, Excel version/build, init; load+save omits `script`; `excel-run` records `script.js`) |
| `config.json` | Optional. `maxPeakMemoryBytes` and `maxDurationMs` budgets. Missing file is valid. Written by Windows `measure-budgets`. |

Do not mix goldens into a flat init dump. Do not plant a dummy `script.js` on load+save cases.

Office.js corpus: `default/simple_set_a1/` copies the `roundtrip/simple` init and sets A1 to `calipers` (`excel-run` golden is committed). `roundtrip/simple/` stays load+save only. `scratch/` cases copy `roundtrip/empty/init.xlsx` and each run one Office.js feature (tables, charts, pivot, spill, CF, basic Excel). Scratch goldens are committed from `excel-run`; `excel-save-pass` still skips them because they are scripted. The default `verify` walk includes scratch once those goldens exist.

`officejs/` is a larger blank-init Office.js suite (tables, pivots, charts, spill, styles, small and large writes, worksheets, names, CF, validation, autofilter, comments, hyperlinks, sort, insert). Scripts are generated from `scripts/gen-officejs-cases`. Goldens and budget numbers are a Windows follow-up (`excel-run-pass`, `measure-budgets`). Default `verify` skips them until goldens exist.

Original source names are preserved in the tables below.

## Scratch (Office.js from empty)

Each case is `scratch/<feature>/` with a copy of `roundtrip/empty/init.xlsx` and a non-empty `script.js` (`await Excel.run(...)`). Seed cells in the same script are setup for that one feature.

| Case | Feature |
|------|---------|
| `text/` | Text cell values |
| `numbers/` | Numeric cell values |
| `column_row_sizes/` | `columnWidth` / `rowHeight` |
| `freeze_panes/` | Freeze first row |
| `spill/` | Spilling `SEQUENCE` formula |
| `named_range/` | Workbook named range |
| `merge/` | Merge cells |
| `number_format/` | Number format |
| `add_sheet/` | Add a worksheet |
| `table/` | `tables.add` |
| `chart/` | `charts.add` |
| `conditional_formatting/` | `conditionalFormats.add` |
| `pivot_table/` | `pivotTables.add` (row + data hierarchy) |
| `data_validation/` | List data validation |
| `autofilter/` | AutoFilter |

## officejs (blank-init Office.js, pending goldens)

Each case is `officejs/<feature>/` with a copy of `roundtrip/empty/init.xlsx` and a non-empty `script.js`. Regenerated with `go run ./scripts/gen-officejs-cases`. Families: range values/formulas (small and 200–500 row writes), font/fill/border/alignment/number formats, tables (style, rows, totals, sort, columns), pivot (row/column/filter/data), charts (column/bar/line/pie/area/scatter/doughnut + title/legend/axes), spill (`SEQUENCE`, `FILTER`, `UNIQUE`, `SORT`, `SORTBY`, `XLOOKUP`, `XMATCH`, `TRANSPOSE`, `VSTACK`, `TAKE`/`DROP`, `CHOOSEROWS`), worksheets, names, CF, validation, autofilter, sort, insert/delete/hide, comments, hyperlinks, merge.

See [`officejs/README.md`](officejs/README.md). Windows: `calipers excel-run-pass verification\cases\officejs`.

## Authored fixtures

Workbooks assembled for basic layout and formula-category coverage.

| Case | Coverage |
|------|----------|
| `empty/` | Empty workbook |
| `simple/` | Minimal values (load+save) |
| `simple_set_a1/` | Copy of `simple` init + Office.js that sets A1 |
| `custom_view_printer_settings/` | Custom sheet views with nested pageSetup + printer rels (mog #360) |
| `theme_linked_colors/` | Font color `theme=4 tint=0.2` (mog #329) |
| `names_add_defined_names_order/` | `names.add` on an Excel-shaped workbook (mog #332) |
| `add_sheet_sparse_ids/` | Add a sheet after importing `sheetId="2"` (mog #334) |
| `chart_titles/` | `charts.add` with chart + axis titles (mog #338) |
| `multi_row_formulas/` | Multi-row `Range.formulas` matrix (mog #328) |
| `types/` | Cell types |
| `styled/` | Basic styles |
| `formulas/` | Formulas |
| `unicode/` | Unicode text |
| `multi_sheet/` | Multiple sheets |
| `bg_color_ranges/` | Background color ranges |
| `lbo_model/` | LBO-style model |
| `_disabled/stress/` | Large stress workbook |
| `formulas_math_basic/` | Math (basic) |
| `formulas_math_trig/` | Math (trig) |
| `formulas_logical/` | Logical |
| `formulas_text/` | Text |
| `formulas_datetime/` | Date and time |
| `formulas_statistical/` | Statistical |
| `formulas_lookup/` | Lookup and reference |
| `formulas_conditional/` | Conditional aggregates |
| `formulas_dynamic_arrays/` | Dynamic arrays |
| `formulas_financial/` | Financial |
| `formulas_database/` | Database |
| `formulas_engineering/` | Engineering |
| `formulas_information/` | Information |
| `formulas_nested/` | Nested formulas |
| `i18n_hebrew/` | Hebrew / i18n |
| `legacy_indexed_colors/` | Legacy indexed colors |

## SheetJS test_files (Apache 2.0)

http://oss.sheetjs.com/test_files/ — feature-named workbooks. GitHub `SheetJS/test_files` is disabled; the HTTP tree still serves them. Mirror: https://git.sheetjs.com.

| Case | Original |
|------|----------|
| `formula_stress_test/` | `formula_stress_test.xlsx` (Excel 2011 functions, arrays, errors) |
| `merge_cells/` | `merge_cells.xlsx` |
| `named_ranges/` | `named_ranges_2011.xlsx` |
| `rich_text/` | `rich_text_stress.xlsx` |
| `date_cells/` | `xlsx-stream-d-date-cell.xlsx` |
| `lonumbers_2010/` | `LONumbers-2010.xlsx` |
| `lonumbers_2011/` | `LONumbers-2011.xlsx` |
| `rk_number/` | `RkNumber.xlsx` |
| `large_strings/` | `large_strings.xlsx` |
| `custom_properties/` | `custom_properties.xlsx` |
| `mixed_sheets/` | `mixed_sheets.xlsx` |
| `autofilter/` | `AutoFilter.xlsx` |
| `comments/` | `comments_stress_test.xlsx` |
| `pivot_named_range/` | `pivot_table_named_range.xlsx` |

More there: `number_format.xlsm` (re-save as xlsx on Windows; skip `.xlsm` for COM format 51), `pivot_table_test.xlsm`, `hyperlink_stress_test_2011.xlsx` (not always present as xlsx), `BlankSheetTypes.xlsm`.

## Apache POI (Apache 2.0)

https://github.com/apache/poi/tree/trunk/test-data/spreadsheet  
Raw: `https://raw.githubusercontent.com/apache/poi/trunk/test-data/spreadsheet/<name>`

| Case | Original |
|------|----------|
| `sample_ss/` | `SampleSS.xlsx` |
| `booleans/` | `Booleans.xlsx` |
| `date_formats/` | `DateFormatTests.xlsx` |
| `number_formats/` | `NumberFormatTests.xlsx` |
| `formatting/` | `Formatting.xlsx` |
| `widths_heights/` | `WidthsAndHeights.xlsx` |
| `themes/` | `Themes.xlsx` |
| `shared_formulas/` | `shared_formulas.xlsx` |
| `hidden_sheet/` | `TwoSheetsOneHidden.xlsx` |
| `unicode_sheet_name/` | `unicodeSheetName.xlsx` |
| `rtl/` | `right-to-left.xlsx` |
| `inline_strings/` | `InlineStrings.xlsx` |
| `styles/` | `styles.xlsx` |
| `xlookup/` | `xlookup.xlsx` |
| `formula_eval/` | `formula-eval.xlsx` |
| `font_theme_colours/` | `50784-font_theme_colours.xlsx` |
| `border_colours/` | `50846-border_colours.xlsx` |
| `row_col_groups/` | `GroupTest.xlsx` |
| `sheet_tab_colors/` | `SheetTabColors.xlsx` |
| `simple_multi_cell/` | `SimpleMultiCell.xlsx` |
| `with_chart/` | `WithChart.xlsx` |
| `three_charts/` | `WithThreeCharts.xlsx` |
| `scatter_chart/` | `SimpleScatterChart.xlsx` |
| `pivot/` | `ExcelPivotTableSample.xlsx` |
| `excel_tables/` | `ExcelTables.xlsx` |
| `tables/` | `Tables.xlsx` |
| `with_table/` | `WithTable.xlsx` |
| `table_50867/` | `50867_with_table.xlsx` |
| `conditional_formatting/` | `WithConditionalFormatting.xlsx` |
| `cf_samples/` | `ConditionalFormattingSamples.xlsx` |
| `data_validation/` | `DataValidationEvaluations.xlsx` |
| `simple_comments/` | `SimpleWithComments.xlsx` |
| `comment_test/` | `commentTest.xlsx` |
| `drawing/` | `WithDrawing.xlsx` |
| `picture/` | `picture.xlsx` |
| `hyperlinks/` | `sharedhyperlink.xlsx` |
| `sheet_protection/` | `sheetProtection_allLocked.xlsx` |
| `workbook_protection/` | `workbookProtection_worksheet_protected.xlsx` |
| `print_repeating/` | `RepeatingRowsCols.xlsx` |
| `header_footer/` | `HeaderFooterComplexFormats.xlsx` |
| `structured_refs/` | `StructuredReferences.xlsx` |
| `textbox/` | `WithTextBox.xlsx` |
| `_disabled/strict_ooxml/` | `SampleSS.strict.xlsx` |
| `_disabled/password/` | `protected_passtika.xlsx` (OLE-encrypted, not a zip) |
| `_disabled/analysis_toolpak/` | `atp.xlsx` |
| `_disabled/xmlbomb/` | `poc-xmlbomb.xlsx` |
| `_disabled/corrupted/` | `xlsx-corrupted.xlsx` |
| `_disabled/embedded_ole/` | `WithEmbeded.xlsx` |
| `_disabled/custom_xml/` | `CustomXMLMappings.xlsx` |

Hundreds more in that folder (bug-repro `NNNNN.xlsx`, fuzz crashes). Prefer named files over numbered bugs unless chasing a specific OOXML quirk.

## Where to get more

| Source | License | Why |
|--------|---------|-----|
| [Apache OpenOffice filter testdocs](https://www.openoffice.org/sc/testdocs/) | Apache | Feature matrix built in Excel 2007 (cells, formats, number formats, CF, formulas, names, validation, hyperlinks, protection, autofilter, pivot, comments, charts). Take the **XML (12) / `.xlsx`** column. |
| [LibreOffice `sc/qa/unit/data/xlsx/`](https://wiki.documentfoundation.org/Development/Calc_Import_Unit_Tests) | MPL-2.0 | Import/export fixtures named by feature (`condFormat_*.xlsx`, themes, sparklines, pivot). Binaries often live in a large tarball, not all in git. |
| Apache POI `test-data/spreadsheet/` (same tree as above) | Apache 2.0 | Remaining named files: `Themes2.xlsx`, `NewStyleConditionalFormattings.xlsx`, `MatrixFormulaEvalTestData.xlsx`, `WithTwoCharts.xlsx`, `workbookProtection_*.xlsx`, `DataValidations-49244.xlsx`, … |
| SheetJS HTTP tree (same as above) | Apache 2.0 | Remaining xlsx; convert `.xlsm` via Windows Excel Save As. |
| [ClosedXML examples](https://github.com/ClosedXML/ClosedXML) | MIT | Generate feature workbooks from C# samples. |
| Generate with Excel COM | ours | Best for license + 1:1 mapping to engine parity docs. `calipers excel-save` on Windows. |

**Avoid as committed goldens:** FUSE (~10.5k OOXML, Zenodo CC-BY) and SpreadsheetBench (~5.4k) — unlabeled real-world crawl, PII risk, too large. Fine later as soak tests, not this directory.

Skip `.xlsm` until the COM host can SaveAs macro-enabled (format 52). `_disabled/xmlbomb` is not loaded.
