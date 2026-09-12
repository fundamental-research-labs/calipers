// officejs: freeze a row then unfreeze.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1:C2").values = [["K", "V1", "V2"], ["a", 1, 2]];
  sheet.freezePanes.freezeRows(1);
  sheet.freezePanes.unfreeze();
  await context.sync();
});
