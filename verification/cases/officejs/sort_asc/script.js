// officejs: sort a range ascending by the first column.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1:B4").values = [
    [3, "c"],
    [1, "a"],
    [4, "d"],
    [2, "b"],
  ];
  sheet.getRange("A1:B4").sort.apply([{ key: 0, ascending: true }]);
  await context.sync();
});
