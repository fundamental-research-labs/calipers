// officejs: add a sheet and move it to position 0.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const extra = context.workbook.worksheets.add("Front");
  extra.position = 0;
  await context.sync();
});
