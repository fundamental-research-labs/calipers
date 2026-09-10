// officejs: write 500 rows of values in one assignment.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const rows = [];
  for (let i = 1; i <= 500; i++) {
    rows.push([i, i * 2, "r" + i]);
  }
  sheet.getRange("A1:C500").values = rows;
  await context.sync();
});
