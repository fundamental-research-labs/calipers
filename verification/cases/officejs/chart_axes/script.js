// officejs: chart with category and value axis titles.
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
  chart.axes.categoryAxis.title.text = "Name";
  chart.axes.valueAxis.title.text = "Value";
  await context.sync();
});
