// officejs: mixed text, number, and boolean values.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1:C2").values = [
    ["n", 1, true],
    ["m", 2, false],
  ];
  await context.sync();
});
