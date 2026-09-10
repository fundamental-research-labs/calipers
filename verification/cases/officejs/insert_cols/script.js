// officejs: insert a column shifting right.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1:B2").values = [[1, 2], [3, 4]];
  sheet.getRange("B1:B2").insert(Excel.InsertShiftDirection.right);
  sheet.getRange("B1:B2").values = [[8], [9]];
  await context.sync();
});
