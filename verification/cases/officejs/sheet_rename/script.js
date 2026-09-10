// officejs: rename the active worksheet.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1").values = [["renamed"]];
  sheet.name = "Renamed";
  await context.sync();
});
