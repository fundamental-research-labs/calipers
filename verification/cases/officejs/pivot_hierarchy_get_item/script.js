// officejs: bind PivotHierarchyCollection.getItem then add the hierarchy.
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
    "LookupPivot",
    data.getRange("A1:C6"),
    dest.getRange("A1")
  );
  const region = pivot.hierarchies.getItem("Region");
  const sales = pivot.hierarchies.getItem("Sales");
  pivot.rowHierarchies.add(region);
  pivot.dataHierarchies.add(sales);
  await context.sync();
});
