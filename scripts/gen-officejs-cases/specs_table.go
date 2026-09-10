package main

func tableSpecs() []spec {
	const sales = `  sheet.getRange("A1:C4").values = [
    ["Region", "Product", "Sales"],
    ["East", "A", 10],
    ["West", "A", 20],
    ["East", "B", 30],
  ];`

	return []spec{
		sheetJS("table_add", "officejs: create an Excel table from a small range.", sales+`
  sheet.tables.add("A1:C4", true);`),
		sheetJS("table_style", "officejs: table with a built-in table style.", sales+`
  const table = sheet.tables.add("A1:C4", true);
  table.style = "TableStyleMedium2";`),
		sheetJS("table_rows_add", "officejs: add a row to an existing table.", sales+`
  const table = sheet.tables.add("A1:C4", true);
  table.rows.add(null, [["West", "B", 40]]);`),
		sheetJS("table_totals", "officejs: show a totals row with SUM.", sales+`
  const table = sheet.tables.add("A1:C4", true);
  table.showTotals = true;
  table.columns.getItem("Sales").totalsRowFunction = Excel.AggregationFunction.sum;`),
		sheetJS("table_sort", "officejs: sort a table by the Sales column.", sales+`
  const table = sheet.tables.add("A1:C4", true);
  table.sort.apply([{ key: 2, ascending: true }]);`),
		sheetJS("table_column_add", "officejs: add a column to a table.", `  sheet.getRange("A1:B3").values = [
    ["Name", "Value"],
    ["a", 1],
    ["b", 2],
  ];
  const table = sheet.tables.add("A1:B3", true);
  table.columns.add(null, [["Extra"], [10], [20]]);`),
		sheetJS("table_banded", "officejs: banded rows and columns on a table.", sales+`
  const table = sheet.tables.add("A1:C4", true);
  table.showBandedRows = true;
  table.showBandedColumns = true;`),
		sheetJS("table_name", "officejs: name a table.", sales+`
  const table = sheet.tables.add("A1:C4", true);
  table.name = "SalesTable";`),
		sheetJS("table_highlight", "officejs: highlight first and last table columns.", sales+`
  const table = sheet.tables.add("A1:C4", true);
  table.highlightFirstColumn = true;
  table.highlightLastColumn = true;`),
		rawJS("pivot_row_data", `// officejs: pivot table with a row and a data hierarchy.
await Excel.run(async (context) => {
  const data = context.workbook.worksheets.getActiveWorksheet();
  data.getRange("A1:C6").values = [
    ["Region", "Product", "Sales"],
    ["East", "A", 10],
    ["West", "A", 20],
    ["East", "B", 30],
    ["West", "B", 40],
    ["East", "A", 15],
  ];
  await context.sync();
  const dest = context.workbook.worksheets.add("Pivot");
  await context.sync();
  const pivot = context.workbook.pivotTables.add(
    "SalesPivot",
    data.getRange("A1:C6"),
    dest.getRange("A1")
  );
  pivot.rowHierarchies.add(pivot.hierarchies.getItem("Region"));
  pivot.dataHierarchies.add(pivot.hierarchies.getItem("Sales"));
  await context.sync();
});
`),
		rawJS("pivot_row_col_data", `// officejs: pivot table with row, column, and data hierarchies.
await Excel.run(async (context) => {
  const data = context.workbook.worksheets.getActiveWorksheet();
  data.getRange("A1:C6").values = [
    ["Region", "Product", "Sales"],
    ["East", "A", 10],
    ["West", "A", 20],
    ["East", "B", 30],
    ["West", "B", 40],
    ["East", "A", 15],
  ];
  await context.sync();
  const dest = context.workbook.worksheets.add("Pivot");
  await context.sync();
  const pivot = context.workbook.pivotTables.add(
    "GridPivot",
    data.getRange("A1:C6"),
    dest.getRange("A1")
  );
  pivot.rowHierarchies.add(pivot.hierarchies.getItem("Region"));
  pivot.columnHierarchies.add(pivot.hierarchies.getItem("Product"));
  pivot.dataHierarchies.add(pivot.hierarchies.getItem("Sales"));
  await context.sync();
});
`),
		rawJS("pivot_filter", `// officejs: pivot table with a filter hierarchy.
await Excel.run(async (context) => {
  const data = context.workbook.worksheets.getActiveWorksheet();
  data.getRange("A1:C6").values = [
    ["Region", "Product", "Sales"],
    ["East", "A", 10],
    ["West", "A", 20],
    ["East", "B", 30],
    ["West", "B", 40],
    ["East", "A", 15],
  ];
  await context.sync();
  const dest = context.workbook.worksheets.add("Pivot");
  await context.sync();
  const pivot = context.workbook.pivotTables.add(
    "FilterPivot",
    data.getRange("A1:C6"),
    dest.getRange("A1")
  );
  pivot.rowHierarchies.add(pivot.hierarchies.getItem("Region"));
  pivot.filterHierarchies.add(pivot.hierarchies.getItem("Product"));
  pivot.dataHierarchies.add(pivot.hierarchies.getItem("Sales"));
  await context.sync();
});
`),
		rawJS("pivot_count", `// officejs: pivot data hierarchy summarized with COUNT.
await Excel.run(async (context) => {
  const data = context.workbook.worksheets.getActiveWorksheet();
  data.getRange("A1:B6").values = [
    ["Region", "Product"],
    ["East", "A"],
    ["West", "A"],
    ["East", "B"],
    ["West", "B"],
    ["East", "A"],
  ];
  await context.sync();
  const dest = context.workbook.worksheets.add("Pivot");
  await context.sync();
  const pivot = context.workbook.pivotTables.add(
    "CountPivot",
    data.getRange("A1:B6"),
    dest.getRange("A1")
  );
  pivot.rowHierarchies.add(pivot.hierarchies.getItem("Region"));
  const dh = pivot.dataHierarchies.add(pivot.hierarchies.getItem("Product"));
  dh.summarizeBy = Excel.AggregationFunction.count;
  await context.sync();
});
`),
		rawJS("pivot_two_data", `// officejs: pivot table with two data fields.
await Excel.run(async (context) => {
  const data = context.workbook.worksheets.getActiveWorksheet();
  data.getRange("A1:D5").values = [
    ["Region", "Product", "Sales", "Qty"],
    ["East", "A", 10, 2],
    ["West", "A", 20, 4],
    ["East", "B", 30, 3],
    ["West", "B", 40, 5],
  ];
  await context.sync();
  const dest = context.workbook.worksheets.add("Pivot");
  await context.sync();
  const pivot = context.workbook.pivotTables.add(
    "TwoDataPivot",
    data.getRange("A1:D5"),
    dest.getRange("A1")
  );
  pivot.rowHierarchies.add(pivot.hierarchies.getItem("Region"));
  pivot.dataHierarchies.add(pivot.hierarchies.getItem("Sales"));
  pivot.dataHierarchies.add(pivot.hierarchies.getItem("Qty"));
  await context.sync();
});
`),
		rawJS("pivot_name", `// officejs: name a pivot table.
await Excel.run(async (context) => {
  const data = context.workbook.worksheets.getActiveWorksheet();
  data.getRange("A1:C4").values = [
    ["Region", "Product", "Sales"],
    ["East", "A", 10],
    ["West", "A", 20],
    ["East", "B", 30],
  ];
  await context.sync();
  const dest = context.workbook.worksheets.add("Pivot");
  await context.sync();
  const pivot = context.workbook.pivotTables.add(
    "NamedPivot",
    data.getRange("A1:C4"),
    dest.getRange("A1")
  );
  pivot.rowHierarchies.add(pivot.hierarchies.getItem("Region"));
  pivot.dataHierarchies.add(pivot.hierarchies.getItem("Sales"));
  pivot.name = "NamedPivot";
  await context.sync();
});
`),
	}
}
