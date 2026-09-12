// officejs: set list validation then dataValidation.clear.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const range = sheet.getRange("A1");
  range.dataValidation.rule = {
    list: { inCellDropDown: true, source: "Yes,No,Maybe" },
  };
  range.dataValidation.clear();
  await context.sync();
});
