// Scratch: add a chart from a small data range.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1:B3").values = [
    ["Name", "Value"],
    ["a", 1],
    ["b", 2],
  ];
  sheet.charts.add(Excel.ChartType.columnClustered, sheet.getRange("A1:B3"));
  await context.sync();
});
