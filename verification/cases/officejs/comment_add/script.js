// officejs: add a threaded comment on A1.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1").values = [["noted"]];
  context.workbook.comments.add("Sheet1!A1", "hello from officejs");
  await context.sync();
});
