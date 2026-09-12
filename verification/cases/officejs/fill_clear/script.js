// officejs: set a fill then format.fill.clear.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1").values = [["x"]];
  sheet.getRange("A1").format.fill.color = "#FFFF00";
  sheet.getRange("A1").format.fill.clear();
  await context.sync();
});
