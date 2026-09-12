// officejs: write via worksheets.getItem.
await Excel.run(async (context) => {
  context.workbook.worksheets.add("Second");
  context.workbook.worksheets.getItem("Second").getRange("A1").values = [["via getItem"]];
  await context.sync();
});
