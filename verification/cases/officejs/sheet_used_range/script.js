// officejs: write via Worksheet.getUsedRange.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("B2").values = [["x"]];
  sheet.getRange("D4").values = [["y"]];
  sheet.getUsedRange().values = [
    ["a", "b", "c"],
    ["d", "e", "f"],
    ["g", "h", "i"],
  ];
  await context.sync();
});
