// officejs: write via Worksheet.getPrevious.
await Excel.run(async (context) => {
  const first = context.workbook.worksheets.getActiveWorksheet();
  first.name = "First";
  first.getRange("A1").values = [["first"]];
  const second = context.workbook.worksheets.add("Second");
  second.getPrevious().getRange("B1").values = [["via prev"]];
  await context.sync();
});
