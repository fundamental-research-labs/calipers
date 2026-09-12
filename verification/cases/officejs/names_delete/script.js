// officejs: add a named range then NamedItem.delete.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1").values = [[42]];
  const named = context.workbook.names.add("TempName", sheet.getRange("A1"));
  named.delete();
  sheet.getRange("B1").values = [["gone"]];
  await context.sync();
});
