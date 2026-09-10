// officejs: solid fill colors.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1").values = [["y"]];
  sheet.getRange("A1").format.fill.color = "#FFFF00";
  sheet.getRange("B1").values = [["o"]];
  sheet.getRange("B1").format.fill.color = "#FF9900";
  await context.sync();
});
