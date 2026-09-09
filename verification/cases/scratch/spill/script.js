// Scratch: write a spilling dynamic-array formula.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1").formulas = [["=SEQUENCE(4,1,1,1)"]];
  await context.sync();
});
