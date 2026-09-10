// officejs: line chart.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1:B5").values = [
    ["Name", "Value"],
    ["a", 10],
    ["b", 20],
    ["c", 15],
    ["d", 25],
  ];
  sheet.charts.add(Excel.ChartType.line, sheet.getRange("A1:B5"));
  await context.sync();
});
