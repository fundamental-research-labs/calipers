// Bulk-write a mixed values/formulas matrix across two rows.
// Catches mog #328: formula strings on later rows must survive export.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1:C2").formulas = [
    [100, "=A1", "=B1+B2"],
    [200, "=A2", null],
  ];
  await context.sync();
});
