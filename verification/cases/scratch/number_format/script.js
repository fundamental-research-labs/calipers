// Scratch: apply a number format to a cell.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1").values = [[0.25]];
  sheet.getRange("A1").numberFormat = [["0.00%"]];
  await context.sync();
});
