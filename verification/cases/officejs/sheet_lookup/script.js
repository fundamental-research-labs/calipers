// officejs: worksheet collection getCount / getItemOrNullObject and getNextOrNullObject / getPreviousOrNullObject.
await Excel.run(async (context) => {
  const sheets = context.workbook.worksheets;
  const first = sheets.getActiveWorksheet();
  first.name = "First";
  const second = sheets.add("Second");
  const next = first.getNextOrNullObject();
  next.getRange("A1").values = [["via nextOrNull"]];
  const beyond = second.getNextOrNullObject();
  const before = first.getPreviousOrNullObject();
  const missing = context.workbook.worksheets.getItemOrNullObject("Missing");
  const count = context.workbook.worksheets.getCount();
  await context.sync();
  first.getRange("B1").values = [[count.value]];
  first.getRange("C1").values = [[beyond.isNullObject ? "no-next" : "next"]];
  first.getRange("D1").values = [[before.isNullObject ? "no-prev" : "prev"]];
  first.getRange("E1").values = [[missing.isNullObject ? "missing" : "hit"]];
  await context.sync();
});
