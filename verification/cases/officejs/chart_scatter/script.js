// officejs: XY scatter chart from numeric pairs.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1:B5").values = [
    ["X", "Y"],
    [1, 2],
    [2, 4],
    [3, 5],
    [4, 4],
  ];
  sheet.charts.add(Excel.ChartType.xyScatter, sheet.getRange("A1:B5"));
  await context.sync();
});
