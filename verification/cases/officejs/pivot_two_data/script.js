// officejs: pivot table with two data fields.
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
