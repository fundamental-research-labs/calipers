// First Office.js corpus case (issue #2): set A1 to a known value.
// Paired with init.xlsx in this directory. Missing/empty script.js means load+save only.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1").values = [["calipers"]];
  await context.sync();
});
