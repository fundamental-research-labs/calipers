// officejs: freeze the first row.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1:C1").values = [["Name", "Qty", "Price"]];
  sheet.getRange("A2:C3").values = [["a", 1, 2], ["b", 3, 4]];
  sheet.freezePanes.freezeRows(1);
  await context.sync();
});
