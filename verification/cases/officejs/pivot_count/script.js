// officejs: pivot data hierarchy summarized with COUNT.
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
