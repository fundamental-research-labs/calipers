package main

func styleSpecs() []spec {
	return []spec{
		sheetJS("font_name_size", "officejs: set font name and size.", `  sheet.getRange("A1").values = [["Hello"]];
  sheet.getRange("A1").format.font.name = "Calibri";
  sheet.getRange("A1").format.font.size = 16;`),
		sheetJS("font_bold_italic", "officejs: bold and italic font.", `  sheet.getRange("A1").values = [["Bold"]];
  sheet.getRange("A1").format.font.bold = true;
  sheet.getRange("B1").values = [["Italic"]];
  sheet.getRange("B1").format.font.italic = true;`),
		sheetJS("font_color", "officejs: font color.", `  sheet.getRange("A1").values = [["red"]];
  sheet.getRange("A1").format.font.color = "#FF0000";
  sheet.getRange("B1").values = [["blue"]];
  sheet.getRange("B1").format.font.color = "#0000FF";`),
		sheetJS("font_underline", "officejs: underline font.", `  sheet.getRange("A1").values = [["under"]];
  sheet.getRange("A1").format.font.underline = Excel.RangeUnderlineStyle.single;`),
		sheetJS("fill_solid", "officejs: solid fill colors.", `  sheet.getRange("A1").values = [["y"]];
  sheet.getRange("A1").format.fill.color = "#FFFF00";
  sheet.getRange("B1").values = [["o"]];
  sheet.getRange("B1").format.fill.color = "#FF9900";`),
		sheetJS("borders_outline", "officejs: outline borders on a block.", `  sheet.getRange("A1:C3").values = [
    [1, 2, 3],
    [4, 5, 6],
    [7, 8, 9],
  ];
  const b = sheet.getRange("A1:C3").format.borders;
  b.getItem("EdgeTop").style = Excel.BorderLineStyle.continuous;
  b.getItem("EdgeBottom").style = Excel.BorderLineStyle.continuous;
  b.getItem("EdgeLeft").style = Excel.BorderLineStyle.continuous;
  b.getItem("EdgeRight").style = Excel.BorderLineStyle.continuous;`),
		sheetJS("borders_inside", "officejs: inside horizontal/vertical borders.", `  sheet.getRange("A1:B2").values = [[1, 2], [3, 4]];
  const b = sheet.getRange("A1:B2").format.borders;
  b.getItem("InsideHorizontal").style = Excel.BorderLineStyle.continuous;
  b.getItem("InsideVertical").style = Excel.BorderLineStyle.continuous;
  b.getItem("InsideHorizontal").color = "#666666";
  b.getItem("InsideVertical").color = "#666666";`),
		sheetJS("align_center", "officejs: horizontal and vertical alignment.", `  sheet.getRange("A1").values = [["ctr"]];
  sheet.getRange("A1").format.horizontalAlignment = Excel.HorizontalAlignment.center;
  sheet.getRange("A1").format.verticalAlignment = Excel.VerticalAlignment.center;
  sheet.getRange("A1").format.rowHeight = 24;
  sheet.getRange("A1").format.columnWidth = 16;`),
		sheetJS("align_wrap", "officejs: wrap a long text cell.", `  sheet.getRange("A1").values = [["a long string that should wrap onto more than one line"]];
  sheet.getRange("A1").format.wrapText = true;
  sheet.getRange("A1").format.columnWidth = 18;
  sheet.getRange("A1").format.rowHeight = 36;`),
		sheetJS("style_good", "officejs: apply the built-in Good named style.", `  sheet.getRange("A1").values = [["ok"]];
  sheet.getRange("A1").style = "Good";`),
		sheetJS("style_input", "officejs: apply the built-in Input named style.", `  sheet.getRange("A1").values = [[42]];
  sheet.getRange("A1").style = "Input";`),
		sheetJS("numfmt_currency", "officejs: currency number format.", `  sheet.getRange("A1:A3").values = [[1.5], [20], [-3.25]];
  sheet.getRange("A1:A3").numberFormat = [["$#,##0.00"], ["$#,##0.00"], ["$#,##0.00"]];`),
		sheetJS("numfmt_percent", "officejs: percent number format.", `  sheet.getRange("A1:A3").values = [[0.25], [0.5], [1]];
  sheet.getRange("A1:A3").numberFormat = [["0.00%"], ["0.00%"], ["0.00%"]];`),
		sheetJS("numfmt_date", "officejs: date number format on serials.", `  sheet.getRange("A1:A2").values = [[44927], [45300]];
  sheet.getRange("A1:A2").numberFormat = [["m/d/yyyy"], ["yyyy-mm-dd"]];`),
		sheetJS("numfmt_accounting", "officejs: accounting number format.", `  sheet.getRange("A1:A2").values = [[1234.5], [-50]];
  sheet.getRange("A1:A2").numberFormat = [["_($* #,##0.00_)"], ["_($* #,##0.00_)"]];`),
		sheetJS("col_row_size", "officejs: column width and row height.", `  sheet.getRange("A1").values = [["sized"]];
  sheet.getRange("A1").format.columnWidth = 28;
  sheet.getRange("A1").format.rowHeight = 22;
  sheet.getRange("B1").format.columnWidth = 12;`),
		sheetJS("fill_clear", "officejs: set a fill then format.fill.clear.", `  sheet.getRange("A1").values = [["x"]];
  sheet.getRange("A1").format.fill.color = "#FFFF00";
  sheet.getRange("A1").format.fill.clear();`),
	}
}
