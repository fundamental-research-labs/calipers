// officejs: whole-number data validation.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const range = sheet.getRange("A1:A5");
  range.dataValidation.rule = {
    wholeNumber: {
      formula1: "1",
      formula2: "10",
      operator: Excel.DataValidationOperator.between,
    },
  };
  await context.sync();
});
