// officejs: list data validation.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const range = sheet.getRange("A1");
  range.dataValidation.rule = {
    list: { inCellDropDown: true, source: "Yes,No,Maybe" },
  };
  await context.sync();
});
