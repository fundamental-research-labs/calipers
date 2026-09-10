// officejs: apply the built-in Good named style.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1").values = [["ok"]];
  sheet.getRange("A1").style = "Good";
  await context.sync();
});
