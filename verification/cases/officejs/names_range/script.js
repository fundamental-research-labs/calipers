// officejs: workbook named range.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1").values = [[42]];
  context.workbook.names.add("Answer", sheet.getRange("A1"));
  await context.sync();
});
