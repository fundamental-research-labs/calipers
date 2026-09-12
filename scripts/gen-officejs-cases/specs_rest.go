package main

func restSpecs() []spec {
	return []spec{
		rawJS("sheet_add", `// officejs: add a worksheet.
await Excel.run(async (context) => {
  context.workbook.worksheets.add("Second");
  await context.sync();
});
`),
		sheetJS("sheet_rename", "officejs: rename the active worksheet.", `  sheet.getRange("A1").values = [["renamed"]];
  sheet.name = "Renamed";`),
		sheetJS("sheet_tab_color", "officejs: set worksheet tab color.", `  sheet.getRange("A1").values = [["tab"]];
  sheet.tabColor = "#FF0000";`),
		rawJS("sheet_hidden", `// officejs: add a sheet and hide it.
await Excel.run(async (context) => {
  const hidden = context.workbook.worksheets.add("Hidden");
  hidden.getRange("A1").values = [["secret"]];
  hidden.visibility = Excel.SheetVisibility.hidden;
  await context.sync();
});
`),
		rawJS("sheet_copy", `// officejs: copy the active worksheet to the end.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1").values = [["copied"]];
  sheet.copy(Excel.WorksheetPositionType.end);
  await context.sync();
});
`),
		sheetJS("sheet_gridlines", "officejs: hide worksheet gridlines.", `  sheet.getRange("A1").values = [["nogrid"]];
  sheet.showGridlines = false;`),
		sheetJS("sheet_position", "officejs: add a sheet and move it to position 0.", `  const extra = context.workbook.worksheets.add("Front");
  extra.position = 0;`),
		sheetJS("freeze_rows", "officejs: freeze the first row.", `  sheet.getRange("A1:C1").values = [["Name", "Qty", "Price"]];
  sheet.getRange("A2:C3").values = [["a", 1, 2], ["b", 3, 4]];
  sheet.freezePanes.freezeRows(1);`),
		sheetJS("freeze_columns", "officejs: freeze the first column.", `  sheet.getRange("A1:C2").values = [["K", "V1", "V2"], ["a", 1, 2]];
  sheet.freezePanes.freezeColumns(1);`),
		sheetJS("names_range", "officejs: workbook named range.", `  sheet.getRange("A1").values = [[42]];
  context.workbook.names.add("Answer", sheet.getRange("A1"));`),
		sheetJS("names_formula", "officejs: workbook named formula.", `  sheet.getRange("A1").values = [[10]];
  context.workbook.names.add("DoubleA1", "=2*Sheet1!$A$1");`),
		sheetJS("names_sheet_scoped", "officejs: worksheet-scoped named range.", `  sheet.getRange("B1").values = [[7]];
  sheet.names.add("Local", sheet.getRange("B1"));`),
		sheetJS("cf_cell_value", "officejs: cell-value conditional formatting.", `  sheet.getRange("A1:A5").values = [[1], [5], [10], [15], [20]];
  const cf = sheet.getRange("A1:A5").conditionalFormats.add(Excel.ConditionalFormatType.cellValue);
  cf.cellValue.rule = {
    formula1: "10",
    operator: Excel.ConditionalCellValueOperator.greaterThan,
  };
  cf.cellValue.format.fill.color = "yellow";`),
		sheetJS("cf_color_scale", "officejs: 3-color scale conditional formatting.", `  sheet.getRange("A1:A5").values = [[1], [5], [10], [15], [20]];
  const cf = sheet.getRange("A1:A5").conditionalFormats.add(Excel.ConditionalFormatType.colorScale);
  cf.colorScale.criteria = {
    minimum: { formula: null, type: Excel.ConditionalFormatColorCriterionType.lowestValue, color: "#F8696B" },
    midpoint: { formula: "50", type: Excel.ConditionalFormatColorCriterionType.percentile, color: "#FFEB84" },
    maximum: { formula: null, type: Excel.ConditionalFormatColorCriterionType.highestValue, color: "#63BE7B" },
  };`),
		sheetJS("cf_data_bar", "officejs: data-bar conditional formatting.", `  sheet.getRange("A1:A5").values = [[1], [5], [10], [15], [20]];
  sheet.getRange("A1:A5").conditionalFormats.add(Excel.ConditionalFormatType.dataBar);`),
		sheetJS("cf_icon_set", "officejs: icon-set conditional formatting.", `  sheet.getRange("A1:A5").values = [[1], [5], [10], [15], [20]];
  const cf = sheet.getRange("A1:A5").conditionalFormats.add(Excel.ConditionalFormatType.iconSet);
  cf.iconSet.style = Excel.IconSet.threeTrafficLights1;`),
		sheetJS("cf_preset", "officejs: preset-criteria conditional formatting.", `  sheet.getRange("A1:A6").values = [[1], [2], [3], [10], [11], [12]];
  const cf = sheet.getRange("A1:A6").conditionalFormats.add(Excel.ConditionalFormatType.presetCriteria);
  cf.preset.rule = { criterion: Excel.ConditionalFormatPresetCriterion.aboveAverage };
  cf.preset.format.fill.color = "lightblue";`),
		sheetJS("cf_text", "officejs: contains-text conditional formatting.", `  sheet.getRange("A1:A4").values = [["East"], ["West"], ["East"], ["North"]];
  const cf = sheet.getRange("A1:A4").conditionalFormats.add(Excel.ConditionalFormatType.containsText);
  cf.textComparison.rule = { operator: Excel.ConditionalTextOperator.contains, text: "East" };
  cf.textComparison.format.font.bold = true;`),
		sheetJS("validation_list", "officejs: list data validation.", `  const range = sheet.getRange("A1");
  range.dataValidation.rule = {
    list: { inCellDropDown: true, source: "Yes,No,Maybe" },
  };`),
		sheetJS("validation_whole", "officejs: whole-number data validation.", `  const range = sheet.getRange("A1:A5");
  range.dataValidation.rule = {
    wholeNumber: {
      formula1: "1",
      formula2: "10",
      operator: Excel.DataValidationOperator.between,
    },
  };`),
		sheetJS("autofilter_apply", "officejs: AutoFilter on a data range.", `  sheet.getRange("A1:B4").values = [
    ["Name", "Value"],
    ["a", 1],
    ["b", 2],
    ["c", 3],
  ];
  sheet.autoFilter.apply(sheet.getRange("A1:B4"));`),
		sheetJS("autofilter_values", "officejs: AutoFilter with a custom criterion.", `  sheet.getRange("A1:B5").values = [
    ["Name", "Value"],
    ["a", 5],
    ["b", 15],
    ["c", 25],
    ["d", 8],
  ];
  sheet.autoFilter.apply(sheet.getRange("A1:B5"), 1, {
    criterion1: ">10",
    filterOn: Excel.FilterOn.custom,
  });`),
		sheetJS("sort_asc", "officejs: sort a range ascending by the first column.", `  sheet.getRange("A1:B4").values = [
    [3, "c"],
    [1, "a"],
    [4, "d"],
    [2, "b"],
  ];
  sheet.getRange("A1:B4").sort.apply([{ key: 0, ascending: true }]);`),
		sheetJS("insert_rows", "officejs: insert rows shifting down.", `  sheet.getRange("A1:A3").values = [[1], [2], [3]];
  sheet.getRange("A2:A2").insert(Excel.InsertShiftDirection.down);
  sheet.getRange("A2").values = [[99]];`),
		sheetJS("insert_cols", "officejs: insert a column shifting right.", `  sheet.getRange("A1:B2").values = [[1, 2], [3, 4]];
  sheet.getRange("B1:B2").insert(Excel.InsertShiftDirection.right);
  sheet.getRange("B1:B2").values = [[8], [9]];`),
		sheetJS("delete_rows", "officejs: delete a row shifting up.", `  sheet.getRange("A1:A4").values = [[1], [2], [3], [4]];
  sheet.getRange("A2:A2").delete(Excel.DeleteShiftDirection.up);`),
		sheetJS("hide_rows_cols", "officejs: hide a row and a column.", `  sheet.getRange("A1:C3").values = [
    [1, 2, 3],
    [4, 5, 6],
    [7, 8, 9],
  ];
  sheet.getRange("A2").getEntireRow().rowHidden = true;
  sheet.getRange("B1").getEntireColumn().columnHidden = true;`),
		sheetJS("comment_add", "officejs: add a threaded comment on A1.", `  sheet.getRange("A1").values = [["noted"]];
  context.workbook.comments.add("Sheet1!A1", "hello from officejs");`),
		sheetJS("hyperlink_url", "officejs: web hyperlink on a cell.", `  sheet.getRange("A1").values = [["example"]];
  sheet.getRange("A1").hyperlink = {
    address: "https://example.com/",
    textToDisplay: "example",
    screenTip: "example.com",
  };`),
		sheetJS("merge_cells", "officejs: merge a 2x2 block.", `  sheet.getRange("A1").values = [["merged"]];
  sheet.getRange("A1:B2").merge();`),
		rawJS("sheet_activate", `// officejs: activate a newly added worksheet and write through it.
await Excel.run(async (context) => {
  const extra = context.workbook.worksheets.add("Second");
  extra.getRange("A1").values = [["second"]];
  extra.activate();
  const active = context.workbook.worksheets.getActiveWorksheet();
  active.getRange("B1").values = [["active"]];
  await context.sync();
});
`),
		rawJS("sheet_delete", `// officejs: add a worksheet then delete it.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1").values = [["kept"]];
  const doomed = context.workbook.worksheets.add("Doomed");
  doomed.getRange("A1").values = [["gone"]];
  doomed.delete();
  await context.sync();
});
`),
		rawJS("sheet_get_next", `// officejs: write via Worksheet.getNext.
await Excel.run(async (context) => {
  const first = context.workbook.worksheets.getActiveWorksheet();
  first.name = "First";
  first.getRange("A1").values = [["first"]];
  context.workbook.worksheets.add("Second");
  first.getNext().getRange("B1").values = [["via next"]];
  await context.sync();
});
`),
		rawJS("sheet_get_previous", `// officejs: write via Worksheet.getPrevious.
await Excel.run(async (context) => {
  const first = context.workbook.worksheets.getActiveWorksheet();
  first.name = "First";
  first.getRange("A1").values = [["first"]];
  const second = context.workbook.worksheets.add("Second");
  second.getPrevious().getRange("B1").values = [["via prev"]];
  await context.sync();
});
`),
		rawJS("sheet_get_first", `// officejs: write via worksheets.getFirst.
await Excel.run(async (context) => {
  context.workbook.worksheets.add("Second");
  context.workbook.worksheets.getFirst().getRange("A1").values = [["first"]];
  await context.sync();
});
`),
		rawJS("sheet_get_last", `// officejs: write via worksheets.getLast.
await Excel.run(async (context) => {
  context.workbook.worksheets.add("Second");
  context.workbook.worksheets.getLast().getRange("A1").values = [["last"]];
  await context.sync();
});
`),
		rawJS("sheet_get_item", `// officejs: write via worksheets.getItem.
await Excel.run(async (context) => {
  context.workbook.worksheets.add("Second");
  context.workbook.worksheets.getItem("Second").getRange("A1").values = [["via getItem"]];
  await context.sync();
});
`),
		sheetJS("names_formula_local", "officejs: workbook named formula via addFormulaLocal.", `  sheet.getRange("A1").values = [[10]];
  context.workbook.names.addFormulaLocal("DoubleA1", "=2*Sheet1!$A$1");
  sheet.getRange("B1").formulas = [["=DoubleA1"]];`),
		sheetJS("names_delete", "officejs: add a named range then NamedItem.delete.", `  sheet.getRange("A1").values = [[42]];
  const named = context.workbook.names.add("TempName", sheet.getRange("A1"));
  named.delete();
  sheet.getRange("B1").values = [["gone"]];`),
		sheetJS("freeze_at", "officejs: freeze panes at a cell.", `  sheet.getRange("A1:C3").values = [
    ["Name", "Qty", "Price"],
    ["a", 1, 2],
    ["b", 3, 4],
  ];
  sheet.freezePanes.freezeAt(sheet.getRange("B2"));`),
		sheetJS("freeze_unfreeze", "officejs: freeze a row then unfreeze.", `  sheet.getRange("A1:C2").values = [["K", "V1", "V2"], ["a", 1, 2]];
  sheet.freezePanes.freezeRows(1);
  sheet.freezePanes.unfreeze();`),
		sheetJS("autofilter_remove", "officejs: apply AutoFilter then remove it.", `  sheet.getRange("A1:B4").values = [
    ["Name", "Value"],
    ["a", 1],
    ["b", 2],
    ["c", 3],
  ];
  sheet.autoFilter.apply(sheet.getRange("A1:B4"));
  sheet.autoFilter.remove();`),
		sheetJS("autofilter_clear_criteria", "officejs: apply AutoFilter then clearCriteria.", `  sheet.getRange("A1:B5").values = [
    ["Name", "Value"],
    ["a", 5],
    ["b", 15],
    ["c", 25],
    ["d", 8],
  ];
  sheet.autoFilter.apply(sheet.getRange("A1:B5"), 1, {
    criterion1: ">10",
    filterOn: Excel.FilterOn.custom,
  });
  sheet.autoFilter.clearCriteria();`),
		sheetJS("autofilter_clear_column_criteria", "officejs: apply AutoFilter then clearColumnCriteria.", `  sheet.getRange("A1:B5").values = [
    ["Name", "Value"],
    ["a", 5],
    ["b", 15],
    ["c", 25],
    ["d", 8],
  ];
  sheet.autoFilter.apply(sheet.getRange("A1:B5"), 1, {
    criterion1: ">10",
    filterOn: Excel.FilterOn.custom,
  });
  sheet.autoFilter.clearColumnCriteria(1);`),
		sheetJS("autofilter_reapply", "officejs: change a filtered cell then reapply AutoFilter.", `  sheet.getRange("A1:B5").values = [
    ["Name", "Value"],
    ["a", 5],
    ["b", 15],
    ["c", 25],
    ["d", 8],
  ];
  sheet.autoFilter.apply(sheet.getRange("A1:B5"), 1, {
    criterion1: ">10",
    filterOn: Excel.FilterOn.custom,
  });
  sheet.getRange("B2").values = [[50]];
  sheet.autoFilter.reapply();`),
		sheetJS("validation_clear", "officejs: set list validation then dataValidation.clear.", `  const range = sheet.getRange("A1");
  range.dataValidation.rule = {
    list: { inCellDropDown: true, source: "Yes,No,Maybe" },
  };
  range.dataValidation.clear();`),
	}
}
