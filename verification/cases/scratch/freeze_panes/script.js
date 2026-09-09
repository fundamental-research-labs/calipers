// Scratch: freeze the first row (worksheet setting).
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1:C1").values = [["Name", "Qty", "Price"]];
  sheet.freezePanes.freezeRows(1);
  await context.sync();
});
