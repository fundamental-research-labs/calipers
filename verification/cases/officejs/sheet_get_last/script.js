// officejs: write via worksheets.getLast.
await Excel.run(async (context) => {
  context.workbook.worksheets.add("Second");
  context.workbook.worksheets.getLast().getRange("A1").values = [["last"]];
  await context.sync();
});
