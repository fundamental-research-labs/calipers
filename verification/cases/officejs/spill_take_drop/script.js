// officejs: TAKE and DROP of a SEQUENCE spill.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1").formulas = [["=SEQUENCE(8,1,1,1)"]];
  sheet.getRange("C1").formulas = [["=TAKE(A1#,3)"]];
  sheet.getRange("E1").formulas = [["=DROP(A1#,3)"]];
  await context.sync();
});
