// officejs: show a totals row with SUM.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1:C4").values = [
    ["Region", "Product", "Sales"],
    ["East", "A", 10],
    ["West", "A", 20],
    ["East", "B", 30],
  ];
  const table = sheet.tables.add("A1:C4", true);
  table.showTotals = true;
  table.columns.getItem("Sales").totalsRowFunction = Excel.AggregationFunction.sum;
  await context.sync();
});
