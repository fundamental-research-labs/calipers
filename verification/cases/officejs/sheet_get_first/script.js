// officejs: write via worksheets.getFirst.
await Excel.run(async (context) => {
  context.workbook.worksheets.add("Second");
  context.workbook.worksheets.getFirst().getRange("A1").values = [["first"]];
  await context.sync();
});
