// officejs: write via Worksheet.getNext.
await Excel.run(async (context) => {
  const first = context.workbook.worksheets.getActiveWorksheet();
  first.name = "First";
  first.getRange("A1").values = [["first"]];
  context.workbook.worksheets.add("Second");
  first.getNext().getRange("B1").values = [["via next"]];
  await context.sync();
});
