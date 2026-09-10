// officejs: add a worksheet.
await Excel.run(async (context) => {
  context.workbook.worksheets.add("Second");
  await context.sync();
});
