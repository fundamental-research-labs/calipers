// officejs: web hyperlink on a cell.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1").values = [["example"]];
  sheet.getRange("A1").hyperlink = {
    address: "https://example.com/",
    textToDisplay: "example",
    screenTip: "example.com",
  };
  await context.sync();
});
