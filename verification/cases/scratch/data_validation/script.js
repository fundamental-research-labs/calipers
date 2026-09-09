// Scratch: apply list data validation to a cell.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const range = sheet.getRange("A1");
  range.dataValidation.rule = {
    list: { inCellDropDown: true, source: "Yes,No" },
  };
  await context.sync();
});
