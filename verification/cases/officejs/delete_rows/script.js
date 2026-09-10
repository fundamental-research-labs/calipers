// officejs: delete a row shifting up.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1:A4").values = [[1], [2], [3], [4]];
  sheet.getRange("A2:A2").delete(Excel.DeleteShiftDirection.up);
  await context.sync();
});
