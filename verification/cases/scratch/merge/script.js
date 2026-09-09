// Scratch: merge a cell range.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1").values = [["merged"]];
  sheet.getRange("A1:B2").merge();
  await context.sync();
});
