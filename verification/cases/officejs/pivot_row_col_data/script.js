// officejs: pivot table with row, column, and data hierarchies.
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
