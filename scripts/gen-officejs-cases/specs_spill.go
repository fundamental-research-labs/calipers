package main

func spillSpecs() []spec {
	return []spec{
		sheetJS("spill_sequence", "officejs: spilling SEQUENCE formula.", `  sheet.getRange("A1").formulas = [["=SEQUENCE(5,1,1,1)"]];`),
		sheetJS("spill_filter", "officejs: spilling FILTER formula.", `  sheet.getRange("A1:B6").values = [
    ["Region", "Sales"],
    ["East", 10],
    ["West", 20],
    ["East", 30],
    ["West", 40],
    ["East", 15],
  ];
  sheet.getRange("D1").formulas = [["=FILTER(A2:B6,A2:A6=\"East\")"]];`),
		sheetJS("spill_unique", "officejs: spilling UNIQUE formula.", `  sheet.getRange("A1:A8").values = [["a"], ["b"], ["a"], ["c"], ["b"], ["d"], ["a"], ["e"]];
  sheet.getRange("C1").formulas = [["=UNIQUE(A1:A8)"]];`),
		sheetJS("spill_sort", "officejs: spilling SORT formula.", `  sheet.getRange("A1:A6").values = [[5], [1], [4], [2], [3], [0]];
  sheet.getRange("C1").formulas = [["=SORT(A1:A6,1,1)"]];`),
		sheetJS("spill_sortby", "officejs: spilling SORTBY formula.", `  sheet.getRange("A1:B5").values = [
    ["n", 3],
    ["a", 1],
    ["c", 4],
    ["b", 2],
    ["d", 0],
  ];
  sheet.getRange("D1").formulas = [["=SORTBY(A1:A5,B1:B5,1)"]];`),
		sheetJS("spill_xlookup", "officejs: XLOOKUP formula.", `  sheet.getRange("A1:B4").values = [
    ["a", 10],
    ["b", 20],
    ["c", 30],
    ["d", 40],
  ];
  sheet.getRange("D1").values = [["c"]];
  sheet.getRange("E1").formulas = [["=XLOOKUP(D1,A1:A4,B1:B4)"]];`),
		sheetJS("spill_xmatch", "officejs: XMATCH formula.", `  sheet.getRange("A1:A4").values = [["a"], ["b"], ["c"], ["d"]];
  sheet.getRange("C1").values = [["c"]];
  sheet.getRange("D1").formulas = [["=XMATCH(C1,A1:A4)"]];`),
		sheetJS("spill_transpose", "officejs: spilling TRANSPOSE formula.", `  sheet.getRange("A1:C1").values = [[1, 2, 3]];
  sheet.getRange("A3").formulas = [["=TRANSPOSE(A1:C1)"]];`),
		sheetJS("spill_vstack", "officejs: spilling VSTACK formula.", `  sheet.getRange("A1:A2").values = [[1], [2]];
  sheet.getRange("C1:C2").values = [[3], [4]];
  sheet.getRange("E1").formulas = [["=VSTACK(A1:A2,C1:C2)"]];`),
		sheetJS("spill_take_drop", "officejs: TAKE and DROP of a SEQUENCE spill.", `  sheet.getRange("A1").formulas = [["=SEQUENCE(8,1,1,1)"]];
  sheet.getRange("C1").formulas = [["=TAKE(A1#,3)"]];
  sheet.getRange("E1").formulas = [["=DROP(A1#,3)"]];`),
		sheetJS("spill_chooserows", "officejs: CHOOSEROWS of a SEQUENCE spill.", `  sheet.getRange("A1").formulas = [["=CHOOSEROWS(SEQUENCE(5,1,1,1),1,3,5)"]];`),
	}
}
