// officejs: font color.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1").values = [["red"]];
  sheet.getRange("A1").format.font.color = "#FF0000";
  sheet.getRange("B1").values = [["blue"]];
  sheet.getRange("B1").format.font.color = "#0000FF";
  await context.sync();
});
