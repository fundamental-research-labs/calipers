// officejs: copy the active worksheet to the end.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1").values = [["copied"]];
  sheet.copy(Excel.WorksheetPositionType.end);
  await context.sync();
});
