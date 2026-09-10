// officejs: set worksheet tab color.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1").values = [["tab"]];
  sheet.tabColor = "#FF0000";
  await context.sync();
});
