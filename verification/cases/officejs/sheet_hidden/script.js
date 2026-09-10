// officejs: add a sheet and hide it.
await Excel.run(async (context) => {
  const hidden = context.workbook.worksheets.add("Hidden");
  hidden.getRange("A1").values = [["secret"]];
  hidden.visibility = Excel.SheetVisibility.hidden;
  await context.sync();
});
