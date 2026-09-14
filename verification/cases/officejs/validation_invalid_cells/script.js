// officejs: dataValidation.getInvalidCells and getInvalidCellsOrNullObject.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const range = sheet.getRange("A1");
  range.dataValidation.rule = {
    list: { inCellDropDown: true, source: "Yes,No" },
  };
  range.values = [["Yes"]];
  const none = range.dataValidation.getInvalidCellsOrNullObject();
  await context.sync();
  sheet.getRange("B1").values = [[none.isNullObject ? "none" : "some"]];
  range.values = [["Nope"]];
  const invalid = range.dataValidation.getInvalidCells();
  invalid.values = [["Yes"]];
  await context.sync();
  await context.sync();
});
