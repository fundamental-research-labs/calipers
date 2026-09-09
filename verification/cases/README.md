# Verification cases

Cases live in a **suite** directory (`roundtrip/` or `default/` until further triage), then `tier_{a|b|c}_<feature>/`:

| Suite | Role |
|-------|------|
| [`roundtrip/`](roundtrip/) | Load+save package comparison |
| [`default/`](default/) | Untriaged (Office.js and `tier_c` hostiles) |

Each case directory:

| File | Role |
|------|------|
| `init.xlsx` | Required input workbook |
| `script.js` | Optional Office.js. Missing or empty → load+save only (skip script execution) |
| `golden.xlsx` | Excel-win oracle. Committed for the default open+save pass (`tier_a`/`tier_b` without Office.js) and for `tier_a_simple_set_a1` (`excel-run`). |
| `golden.xlsx.meta.json` | Sidecar (`host=excel-win`, Excel version/build, init; load+save omits `script`; `excel-run` records `script.js`) |

Do not mix goldens into a flat init dump. Do not plant a dummy `script.js` on load+save cases.

Office.js corpus is starting tiny: `default/tier_a_simple_set_a1/` copies the `roundtrip/tier_a_simple` init and sets A1 to `calipers` (`excel-run` golden is committed). `roundtrip/tier_a_simple/` stays load+save only. More scripts later.

Original source names are preserved in the tables below.

Tiers: **A** = engine claims support (first goldens). **B** = remaining work (Excel keeps, an engine may drop). **C** = hostile / later; skip in the default pass.

## Authored fixtures

Workbooks assembled for basic layout and formula-category coverage.

| Case | Coverage |
|------|----------|
| `tier_a_empty/` | Empty workbook |
| `tier_a_simple/` | Minimal values (load+save) |
| `tier_a_simple_set_a1/` | Copy of `tier_a_simple` init + Office.js that sets A1 |
| `tier_a_types/` | Cell types |
| `tier_a_styled/` | Basic styles |
| `tier_a_formulas/` | Formulas |
| `tier_a_unicode/` | Unicode text |
| `tier_a_multi_sheet/` | Multiple sheets |
| `tier_a_bg_color_ranges/` | Background color ranges |
| `tier_a_lbo_model/` | LBO-style model |
| `tier_c_stress/` | Large stress workbook |
| `tier_a_formulas_math_basic/` | Math (basic) |
| `tier_a_formulas_math_trig/` | Math (trig) |
| `tier_a_formulas_logical/` | Logical |
| `tier_a_formulas_text/` | Text |
| `tier_a_formulas_datetime/` | Date and time |
| `tier_a_formulas_statistical/` | Statistical |
| `tier_a_formulas_lookup/` | Lookup and reference |
| `tier_a_formulas_conditional/` | Conditional aggregates |
| `tier_a_formulas_dynamic_arrays/` | Dynamic arrays |
| `tier_a_formulas_financial/` | Financial |
| `tier_a_formulas_database/` | Database |
| `tier_a_formulas_engineering/` | Engineering |
| `tier_a_formulas_information/` | Information |
| `tier_a_formulas_nested/` | Nested formulas |
| `tier_a_i18n_hebrew/` | Hebrew / i18n |
| `tier_a_legacy_indexed_colors/` | Legacy indexed colors |

## SheetJS test_files (Apache 2.0)

http://oss.sheetjs.com/test_files/ — feature-named workbooks. GitHub `SheetJS/test_files` is disabled; the HTTP tree still serves them. Mirror: https://git.sheetjs.com.

| Case | Original |
|------|----------|
| `tier_a_formula_stress_test/` | `formula_stress_test.xlsx` (Excel 2011 functions, arrays, errors) |
| `tier_a_merge_cells/` | `merge_cells.xlsx` |
| `tier_a_named_ranges/` | `named_ranges_2011.xlsx` |
| `tier_a_rich_text/` | `rich_text_stress.xlsx` |
| `tier_a_date_cells/` | `xlsx-stream-d-date-cell.xlsx` |
| `tier_a_lonumbers_2010/` | `LONumbers-2010.xlsx` |
| `tier_a_lonumbers_2011/` | `LONumbers-2011.xlsx` |
| `tier_a_rk_number/` | `RkNumber.xlsx` |
| `tier_a_large_strings/` | `large_strings.xlsx` |
| `tier_a_custom_properties/` | `custom_properties.xlsx` |
| `tier_a_mixed_sheets/` | `mixed_sheets.xlsx` |
| `tier_b_autofilter/` | `AutoFilter.xlsx` |
| `tier_b_comments/` | `comments_stress_test.xlsx` |
| `tier_b_pivot_named_range/` | `pivot_table_named_range.xlsx` |

More there: `number_format.xlsm` (re-save as xlsx on Windows; skip `.xlsm` for COM format 51), `pivot_table_test.xlsm`, `hyperlink_stress_test_2011.xlsx` (not always present as xlsx), `BlankSheetTypes.xlsm`.

## Apache POI (Apache 2.0)

https://github.com/apache/poi/tree/trunk/test-data/spreadsheet  
Raw: `https://raw.githubusercontent.com/apache/poi/trunk/test-data/spreadsheet/<name>`

| Case | Original |
|------|----------|
| `tier_a_sample_ss/` | `SampleSS.xlsx` |
| `tier_a_booleans/` | `Booleans.xlsx` |
| `tier_a_date_formats/` | `DateFormatTests.xlsx` |
| `tier_a_number_formats/` | `NumberFormatTests.xlsx` |
| `tier_a_formatting/` | `Formatting.xlsx` |
| `tier_a_widths_heights/` | `WidthsAndHeights.xlsx` |
| `tier_a_themes/` | `Themes.xlsx` |
| `tier_a_shared_formulas/` | `shared_formulas.xlsx` |
| `tier_a_hidden_sheet/` | `TwoSheetsOneHidden.xlsx` |
| `tier_a_unicode_sheet_name/` | `unicodeSheetName.xlsx` |
| `tier_a_rtl/` | `right-to-left.xlsx` |
| `tier_a_inline_strings/` | `InlineStrings.xlsx` |
| `tier_a_styles/` | `styles.xlsx` |
| `tier_a_xlookup/` | `xlookup.xlsx` |
| `tier_a_formula_eval/` | `formula-eval.xlsx` |
| `tier_a_font_theme_colours/` | `50784-font_theme_colours.xlsx` |
| `tier_a_border_colours/` | `50846-border_colours.xlsx` |
| `tier_a_row_col_groups/` | `GroupTest.xlsx` |
| `tier_a_sheet_tab_colors/` | `SheetTabColors.xlsx` |
| `tier_a_simple_multi_cell/` | `SimpleMultiCell.xlsx` |
| `tier_b_with_chart/` | `WithChart.xlsx` |
| `tier_b_three_charts/` | `WithThreeCharts.xlsx` |
| `tier_b_scatter_chart/` | `SimpleScatterChart.xlsx` |
| `tier_b_pivot/` | `ExcelPivotTableSample.xlsx` |
| `tier_b_excel_tables/` | `ExcelTables.xlsx` |
| `tier_b_tables/` | `Tables.xlsx` |
| `tier_b_with_table/` | `WithTable.xlsx` |
| `tier_b_table_50867/` | `50867_with_table.xlsx` |
| `tier_b_conditional_formatting/` | `WithConditionalFormatting.xlsx` |
| `tier_b_cf_samples/` | `ConditionalFormattingSamples.xlsx` |
| `tier_b_data_validation/` | `DataValidationEvaluations.xlsx` |
| `tier_b_simple_comments/` | `SimpleWithComments.xlsx` |
| `tier_b_comment_test/` | `commentTest.xlsx` |
| `tier_b_drawing/` | `WithDrawing.xlsx` |
| `tier_b_picture/` | `picture.xlsx` |
| `tier_b_hyperlinks/` | `sharedhyperlink.xlsx` |
| `tier_b_sheet_protection/` | `sheetProtection_allLocked.xlsx` |
| `tier_b_workbook_protection/` | `workbookProtection_worksheet_protected.xlsx` |
| `tier_b_print_repeating/` | `RepeatingRowsCols.xlsx` |
| `tier_b_header_footer/` | `HeaderFooterComplexFormats.xlsx` |
| `tier_b_structured_refs/` | `StructuredReferences.xlsx` |
| `tier_b_textbox/` | `WithTextBox.xlsx` |
| `tier_c_strict_ooxml/` | `SampleSS.strict.xlsx` |
| `tier_c_password/` | `protected_passtika.xlsx` (OLE-encrypted, not a zip) |
| `tier_c_analysis_toolpak/` | `atp.xlsx` |
| `tier_c_xmlbomb/` | `poc-xmlbomb.xlsx` |
| `tier_c_corrupted/` | `xlsx-corrupted.xlsx` |
| `tier_c_embedded_ole/` | `WithEmbeded.xlsx` |
| `tier_c_custom_xml/` | `CustomXMLMappings.xlsx` |

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

Skip `.xlsm` until the COM host can SaveAs macro-enabled (format 52). Do not run `tier_c_xmlbomb` in the default pass.
