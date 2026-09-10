// officejs: hide worksheet gridlines.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1").values = [["nogrid"]];
  sheet.showGridlines = false;
  await context.sync();
});
