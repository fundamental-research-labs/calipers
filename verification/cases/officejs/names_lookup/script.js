// officejs: named-item getCount / getItem / getItemOrNullObject and getRange / getRangeOrNullObject.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1").values = [[42]];
  context.workbook.names.add("Answer", sheet.getRange("A1"));
  const named = context.workbook.names.getItem("Answer");
  const resolved = named.getRange();
  resolved.values = [["via getItem"]];
  const maybeRange = named.getRangeOrNullObject();
  maybeRange.getCell(0, 0).values = [["via getRangeOrNull"]];
  const missing = context.workbook.names.getItemOrNullObject("Nope");
  const count = context.workbook.names.getCount();
  await context.sync();
  sheet.getRange("C1").values = [[count.value]];
  sheet.getRange("D1").values = [[missing.isNullObject ? "gone" : "hit"]];
  await context.sync();
});
