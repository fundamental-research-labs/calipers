// officejs: freeze the first column.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1:C2").values = [["K", "V1", "V2"], ["a", 1, 2]];
  sheet.freezePanes.freezeColumns(1);
  await context.sync();
});
