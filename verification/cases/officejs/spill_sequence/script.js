// officejs: spilling SEQUENCE formula.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1").formulas = [["=SEQUENCE(5,1,1,1)"]];
  await context.sync();
});
