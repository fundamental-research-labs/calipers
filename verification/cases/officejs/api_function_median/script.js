// Excel.Functions.median: values and boundary behavior.
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();

  sheet.name = "Sheet1";
  sheet.getRange("A1:B3").values = [[1, "text"], [2, true], [3, null]];
  const f = context.workbook.functions;
  const first = f.median(sheet.getRange("A1:A3"));
  const second = f.median(1, 2, 8, 9);
  first.load("value,error");
  second.load("value,error");
  await context.sync();
  // Preserve text beginning with '=' as text when writing through Range.values.
  const cellValue = result => result.error ? "" :
    (typeof result.value === "string" && /^[=+\-']/.test(result.value) ? "'" + result.value : result.value);
  sheet.getRange("D1:F3").values = [
    ["Scenario", "Value", "Error"],
    ["primary", cellValue(first), first.error || ""],
    ["boundary", cellValue(second), second.error || ""]
  ];
  await context.sync();
});
