// officejs: underline font.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1").values = [["under"]];
  sheet.getRange("A1").format.font.underline = Excel.RangeUnderlineStyle.single;
  await context.sync();
});
