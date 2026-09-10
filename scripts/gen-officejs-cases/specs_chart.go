package main

func chartSpecs() []spec {
	const cats = `  sheet.getRange("A1:B5").values = [
    ["Name", "Value"],
    ["a", 10],
    ["b", 20],
    ["c", 15],
    ["d", 25],
  ];`

	return []spec{
		sheetJS("chart_column", "officejs: clustered column chart.", cats+`
  sheet.charts.add(Excel.ChartType.columnClustered, sheet.getRange("A1:B5"));`),
		sheetJS("chart_bar", "officejs: clustered bar chart.", cats+`
  sheet.charts.add(Excel.ChartType.barClustered, sheet.getRange("A1:B5"));`),
		sheetJS("chart_line", "officejs: line chart.", cats+`
  sheet.charts.add(Excel.ChartType.line, sheet.getRange("A1:B5"));`),
		sheetJS("chart_pie", "officejs: pie chart.", cats+`
  sheet.charts.add(Excel.ChartType.pie, sheet.getRange("A1:B5"));`),
		sheetJS("chart_area", "officejs: area chart.", cats+`
  sheet.charts.add(Excel.ChartType.area, sheet.getRange("A1:B5"));`),
		sheetJS("chart_scatter", "officejs: XY scatter chart from numeric pairs.", `  sheet.getRange("A1:B5").values = [
    ["X", "Y"],
    [1, 2],
    [2, 4],
    [3, 5],
    [4, 4],
  ];
  sheet.charts.add(Excel.ChartType.xyScatter, sheet.getRange("A1:B5"));`),
		sheetJS("chart_doughnut", "officejs: doughnut chart.", cats+`
  sheet.charts.add(Excel.ChartType.doughnut, sheet.getRange("A1:B5"));`),
		sheetJS("chart_title", "officejs: column chart with a title.", cats+`
  const chart = sheet.charts.add(Excel.ChartType.columnClustered, sheet.getRange("A1:B5"));
  chart.title.text = "Values";
  chart.title.visible = true;`),
		sheetJS("chart_legend", "officejs: chart legend position.", cats+`
  const chart = sheet.charts.add(Excel.ChartType.columnClustered, sheet.getRange("A1:B5"));
  chart.legend.visible = true;
  chart.legend.position = Excel.ChartLegendPosition.bottom;`),
		sheetJS("chart_axes", "officejs: chart with category and value axis titles.", cats+`
  const chart = sheet.charts.add(Excel.ChartType.columnClustered, sheet.getRange("A1:B5"));
  chart.axes.categoryAxis.title.text = "Name";
  chart.axes.valueAxis.title.text = "Value";`),
	}
}
