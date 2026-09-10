// officejs: worksheet-scoped named range.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("B1").values = [[7]];
  sheet.names.add("Local", sheet.getRange("B1"));
  await context.sync();
});
