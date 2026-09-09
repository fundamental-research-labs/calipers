// First Office.js corpus case (issue #2): set A1 to a known value.
// Init is a copy of roundtrip/simple (that case stays load+save only).
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1").values = [["calipers"]];
  await context.sync();
});
