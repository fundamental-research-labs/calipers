// officejs: set font name and size.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1").values = [["Hello"]];
  sheet.getRange("A1").format.font.name = "Calibri";
  sheet.getRange("A1").format.font.size = 16;
  await context.sync();
});
