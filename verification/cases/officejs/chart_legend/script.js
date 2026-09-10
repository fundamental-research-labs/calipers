// officejs: chart legend position.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1:B5").values = [
    ["Name", "Value"],
    ["a", 10],
    ["b", 20],
    ["c", 15],
    ["d", 25],
  ];
  const chart = sheet.charts.add(Excel.ChartType.columnClustered, sheet.getRange("A1:B5"));
  chart.legend.visible = true;
  chart.legend.position = Excel.ChartLegendPosition.bottom;
  await context.sync();
});
