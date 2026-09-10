// Author a column chart with a chart title and both axis titles.
// Catches mog #338: exported <c:title> must include <c:overlay val="0"/>.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1:B5").values = [
    ["Month", "Units"],
    ["M0", 10],
    ["M1", 13],
    ["M2", 16],
    ["M3", 19],
  ];
  const chart = sheet.charts.add(
    Excel.ChartType.columnClustered,
    sheet.getRange("A1:B5")
  );
  chart.title.text = "Monthly Units";
  chart.axes.categoryAxis.title.text = "Month";
  chart.axes.valueAxis.title.text = "Units Sold";
  await context.sync();
});
