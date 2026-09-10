// officejs: apply the built-in Input named style.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1").values = [[42]];
  sheet.getRange("A1").style = "Input";
  await context.sync();
});
