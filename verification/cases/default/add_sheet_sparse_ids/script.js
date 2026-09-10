// Import a workbook whose only sheetId is 2, then add a worksheet.
// Catches mog #334: generated sheetId values must stay unique and positive.
await Excel.run(async (context) => {
  context.workbook.worksheets.add("Added");
  await context.sync();
});
