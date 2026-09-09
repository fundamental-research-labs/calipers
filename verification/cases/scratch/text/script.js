// Scratch: write text values into an empty workbook.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1:B2").values = [
    ["alpha", "bravo"],
    ["charlie", "delta"],
  ];
  await context.sync();
});
