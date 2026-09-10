// officejs: CHOOSEROWS of a SEQUENCE spill.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1").formulas = [["=CHOOSEROWS(SEQUENCE(5,1,1,1),1,3,5)"]];
  await context.sync();
});
