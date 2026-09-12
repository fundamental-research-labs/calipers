// officejs: add a worksheet then delete it.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1").values = [["kept"]];
  const doomed = context.workbook.worksheets.add("Doomed");
  doomed.getRange("A1").values = [["gone"]];
  doomed.delete();
  await context.sync();
});
