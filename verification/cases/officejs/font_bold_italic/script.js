// officejs: bold and italic font.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1").values = [["Bold"]];
  sheet.getRange("A1").format.font.bold = true;
  sheet.getRange("B1").values = [["Italic"]];
  sheet.getRange("B1").format.font.italic = true;
  await context.sync();
});
