// Load an Excel-shaped workbook (calcPr + extLst) and add a defined name.
// Catches mog #332: <definedNames> must appear before <calcPr>/<extLst>.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  context.workbook.names.add("Repro_Name", sheet.getRange("A1:A10"));
  await context.sync();
});
