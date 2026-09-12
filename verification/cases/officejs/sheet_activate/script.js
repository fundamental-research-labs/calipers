// officejs: activate a newly added worksheet and write through it.
await Excel.run(async (context) => {
  const extra = context.workbook.worksheets.add("Second");
  extra.getRange("A1").values = [["second"]];
  extra.activate();
  const active = context.workbook.worksheets.getActiveWorksheet();
  active.getRange("B1").values = [["active"]];
  await context.sync();
});
