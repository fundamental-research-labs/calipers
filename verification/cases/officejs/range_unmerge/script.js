// officejs: merge a block then unmerge it.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1").values = [["merged"]];
  sheet.getRange("A1:B2").merge();
  sheet.getRange("A1:B2").unmerge();
  sheet.getRange("B2").values = [["split"]];
  await context.sync();
});
