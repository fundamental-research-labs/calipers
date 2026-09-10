package main

func rangeSpecs() []spec {
	return []spec{
		sheetJS("range_set_text", "officejs: set text values on a small range.", `  sheet.getRange("A1:B2").values = [
    ["alpha", "bravo"],
    ["charlie", "delta"],
  ];`),
		sheetJS("range_set_numbers", "officejs: set numeric values on a small range.", `  sheet.getRange("A1:C2").values = [
    [1, 2.5, -3],
    [0, 100, 0.125],
  ];`),
		sheetJS("range_set_booleans", "officejs: set boolean cell values.", `  sheet.getRange("A1:A4").values = [[true], [false], [true], [false]];`),
		sheetJS("range_set_dates", "officejs: write Excel serial dates and a date number format.", `  sheet.getRange("A1:A3").values = [[44927], [44928], [44929]];
  sheet.getRange("A1:A3").numberFormat = [["yyyy-mm-dd"], ["yyyy-mm-dd"], ["yyyy-mm-dd"]];`),
		sheetJS("range_set_mixed", "officejs: mixed text, number, and boolean values.", `  sheet.getRange("A1:C2").values = [
    ["n", 1, true],
    ["m", 2, false],
  ];`),
		sheetJS("range_formulas_sum", "officejs: SUM / AVERAGE formulas over a small range.", `  sheet.getRange("A1:A4").values = [[1], [2], [3], [4]];
  sheet.getRange("B1").formulas = [["=SUM(A1:A4)"]];
  sheet.getRange("B2").formulas = [["=AVERAGE(A1:A4)"]];`),
		sheetJS("range_formulas_if", "officejs: IF formulas over a small range.", `  sheet.getRange("A1:A3").values = [[5], [15], [25]];
  sheet.getRange("B1:B3").formulas = [["=IF(A1>=10,\"hi\",\"lo\")"], ["=IF(A2>=10,\"hi\",\"lo\")"], ["=IF(A3>=10,\"hi\",\"lo\")"]];`),
		sheetJS("range_formulas_nested", "officejs: nested IF / AND formulas.", `  sheet.getRange("A1:B3").values = [
    [1, 10],
    [2, 20],
    [3, 30],
  ];
  sheet.getRange("C1").formulas = [["=IF(AND(A1>0,B1>5),B1-A1,0)"]];
  sheet.getRange("C2").formulas = [["=IF(OR(A2=2,B2>100),\"ok\",\"no\")"]];`),
		sheetJS("range_formulas_lookup", "officejs: INDEX/MATCH lookup formulas.", `  sheet.getRange("A1:B4").values = [
    ["Key", "Val"],
    ["a", 10],
    ["b", 20],
    ["c", 30],
  ];
  sheet.getRange("D1").values = [["b"]];
  sheet.getRange("E1").formulas = [["=INDEX(B2:B4,MATCH(D1,A2:A4,0))"]];`),
		sheetJS("range_clear", "officejs: write values then clear the range.", `  sheet.getRange("A1:B2").values = [[1, 2], [3, 4]];
  sheet.getRange("A1:B2").clear();
  sheet.getRange("C1").values = [["cleared"]];`),
		sheetJS("range_copy_from", "officejs: copyFrom a source range to a destination.", `  sheet.getRange("A1:A3").values = [[1], [2], [3]];
  sheet.getRange("C1").copyFrom(sheet.getRange("A1:A3"));`),
		sheetJS("range_by_indexes", "officejs: write a 2x2 block via getRangeByIndexes.", `  sheet.getRangeByIndexes(0, 0, 2, 2).values = [
    [1, 2],
    [3, 4],
  ];`),
		sheetJS("range_offset", "officejs: write via getOffsetRange from A1.", `  sheet.getRange("A1").values = [[1]];
  sheet.getRange("A1").getOffsetRange(1, 1).values = [[2]];
  sheet.getRange("A1").getOffsetRange(2, 2).values = [[3]];`),
		sheetJS("small_three_cells", "officejs: tiny workbook — two inputs and a sum.", `  sheet.getRange("A1").values = [[10]];
  sheet.getRange("B1").values = [[20]];
  sheet.getRange("C1").formulas = [["=A1+B1"]];`),
		sheetJS("big_values_500", "officejs: write 500 rows of values in one assignment.", `  const rows = [];
  for (let i = 1; i <= 500; i++) {
    rows.push([i, i * 2, "r" + i]);
  }
  sheet.getRange("A1:C500").values = rows;`),
		sheetJS("big_formulas_400", "officejs: 400 rows of values plus dependent formulas.", `  const vals = [];
  const fms = [];
  for (let i = 1; i <= 400; i++) {
    vals.push([i]);
    fms.push(["=A" + i + "*2"]);
  }
  sheet.getRange("A1:A400").values = vals;
  sheet.getRange("B1:B400").formulas = fms;`),
		sheetJS("big_grid_200x12", "officejs: 200x12 numeric grid write.", `  const rows = [];
  for (let r = 1; r <= 200; r++) {
    const row = [];
    for (let c = 1; c <= 12; c++) {
      row.push(r * 100 + c);
    }
    rows.push(row);
  }
  sheet.getRange("A1:L200").values = rows;`),
		rawJS("big_multi_sheet", `// officejs: three sheets, 200 rows each.
await Excel.run(async (context) => {
  const names = ["One", "Two", "Three"];
  for (let n = 0; n < names.length; n++) {
    const s = n === 0
      ? context.workbook.worksheets.getActiveWorksheet()
      : context.workbook.worksheets.add(names[n]);
    if (n === 0) {
      s.name = names[0];
    }
    await context.sync();
    const rows = [];
    for (let i = 1; i <= 200; i++) {
      rows.push([names[n], i, i * 2]);
    }
    s.getRange("A1:C200").values = rows;
    await context.sync();
  }
});
`),
	}
}
