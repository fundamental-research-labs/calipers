// officejs: insert rows shifting down.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1:A3").values = [[1], [2], [3]];
  sheet.getRange("A2:A2").insert(Excel.InsertShiftDirection.down);
  sheet.getRange("A2").values = [[99]];
  await context.sync();
});
