// officejs: write a border via RangeBorderCollection.getItemAt.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1").values = [["x"]];
  const top = sheet.getRange("A1").format.borders.getItemAt(0);
  top.style = Excel.BorderLineStyle.continuous;
  top.color = "#FF0000";
  await context.sync();
});
