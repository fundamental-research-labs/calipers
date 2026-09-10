package xlsxmodel

import (
	"archive/zip"
	"bytes"
	"os"
	"path/filepath"
	"runtime"
	"sort"
	"strconv"
	"strings"
	"testing"
	"time"
)

func TestCompareEqualOnBinary64FloatText(t *testing.T) {
	a := mustXLSX(t, map[string]string{
		"xl/worksheets/sheet1.xml": `<?xml version="1.0"?><worksheet><sheetData><row r="4"><c r="C4"><v>92.3</v></c></row></sheetData></worksheet>`,
	})
	b := mustXLSX(t, map[string]string{
		"xl/worksheets/sheet1.xml": `<?xml version="1.0"?><worksheet><sheetData><row r="4"><c r="C4"><v>92.299999999999997</v></c></row></sheetData></worksheet>`,
	})
	got, err := Compare(a, b)
	if err != nil {
		t.Fatal(err)
	}
	if !got.Equal {
		t.Fatalf("92.3 vs 92.299999999999997 must be equal (binary64), got %v", got.Diffs)
	}
	t.Log("equal: sheet cell 92.3 vs 92.299999999999997")
}

func TestCompareUnequalOnCellString(t *testing.T) {
	a := mustXLSX(t, map[string]string{
		"xl/workbook.xml":          workbookXML("0"),
		"xl/worksheets/sheet1.xml": sheetInline("hello"),
	})
	b := mustXLSX(t, map[string]string{
		"xl/workbook.xml":          workbookXML("0"),
		"xl/worksheets/sheet1.xml": sheetInline("world"),
	})
	got, err := Compare(a, b)
	if err != nil {
		t.Fatal(err)
	}
	if got.Equal {
		t.Fatal("hello vs world must be unequal")
	}
	found := false
	for _, d := range got.Diffs {
		if d.Axis == "values" && d.Location == "Sheet1!A1" {
			found = true
		}
	}
	if !found {
		t.Fatalf("want values Sheet1!A1, got %v", got.Diffs)
	}
	t.Logf("unequal: hello vs world %v", got.Diffs)
}

func TestCompareUnequalOnDate1904(t *testing.T) {
	a := mustXLSX(t, map[string]string{
		"xl/workbook.xml":          workbookXML("0"),
		"xl/worksheets/sheet1.xml": sheetInline("hello"),
	})
	b := mustXLSX(t, map[string]string{
		"xl/workbook.xml":          workbookXML("1"),
		"xl/worksheets/sheet1.xml": sheetInline("hello"),
	})
	got, err := Compare(a, b)
	if err != nil {
		t.Fatal(err)
	}
	if got.Equal {
		t.Fatal("date1904 0 vs 1 must be unequal")
	}
	found := false
	for _, d := range got.Diffs {
		if d.Axis == "date1904" {
			found = true
		}
	}
	if !found {
		t.Fatalf("want date1904 diff, got %v", got.Diffs)
	}
	t.Logf("unequal: date1904 0 vs 1 %v", got.Diffs)
}

func TestCompareUnequalOnDate1904NotClobberedByX15WorkbookPr(t *testing.T) {
	// Corpus Excel files emit a later x15:workbookPr with no date1904.
	// Matching Local=="workbookPr" would overwrite date1904=1 with false.
	with1904 := mustXLSX(t, map[string]string{
		"xl/workbook.xml":          workbookXML1904WithX15("1"),
		"xl/worksheets/sheet1.xml": sheetInline("hello"),
	})
	without := mustXLSX(t, map[string]string{
		"xl/workbook.xml":          workbookXML("0"),
		"xl/worksheets/sheet1.xml": sheetInline("hello"),
	})
	got, err := Compare(with1904, without)
	if err != nil {
		t.Fatal(err)
	}
	if got.Equal {
		t.Fatal("date1904=1 plus x15:workbookPr vs date1904=0 must be unequal")
	}
	found := false
	for _, d := range got.Diffs {
		if d.Axis == "date1904" {
			found = true
		}
	}
	if !found {
		t.Fatalf("want date1904 diff, got %v", got.Diffs)
	}
	t.Logf("unequal: date1904=1 with trailing x15:workbookPr vs 0 %v", got.Diffs)
}

func TestCompareEqualWhenDate1904AbsentMeans1900(t *testing.T) {
	absent := mustXLSX(t, map[string]string{
		"xl/workbook.xml":          `<workbook><workbookPr/><sheets><sheet name="Sheet1" r:id="rId1"/></sheets></workbook>`,
		"xl/worksheets/sheet1.xml": sheetInline("hello"),
	})
	zero := mustXLSX(t, map[string]string{
		"xl/workbook.xml":          workbookXML("0"),
		"xl/worksheets/sheet1.xml": sheetInline("hello"),
	})
	got, err := Compare(absent, zero)
	if err != nil {
		t.Fatal(err)
	}
	if !got.Equal {
		t.Fatalf("absent date1904 must match 0, got %v", got.Diffs)
	}
}

func TestCompareEqualOnEmptyVsSelfClosingCompany(t *testing.T) {
	a := mustXLSX(t, map[string]string{
		"docProps/app.xml":         `<?xml version="1.0"?><Properties><Application>Mog</Application><Company/><AppVersion>0.1</AppVersion></Properties>`,
		"xl/worksheets/sheet1.xml": sheetInline("x"),
	})
	b := mustXLSX(t, map[string]string{
		"docProps/app.xml":         `<?xml version="1.0"?><Properties><Application>Excel</Application><Company></Company><AppVersion>16.0300</AppVersion></Properties>`,
		"xl/worksheets/sheet1.xml": sheetInline("x"),
	})
	got, err := Compare(a, b)
	if err != nil {
		t.Fatal(err)
	}
	if !got.Equal {
		t.Fatalf("Company serialization must not fail value compare, got %v", got.Diffs)
	}
}

func TestCompareUnequalOnBooleanVsNumber(t *testing.T) {
	a := mustXLSX(t, map[string]string{
		"xl/worksheets/sheet1.xml": `<?xml version="1.0"?><worksheet><sheetData><row r="3"><c r="B3" t="b"><v>1</v></c></row></sheetData></worksheet>`,
	})
	b := mustXLSX(t, map[string]string{
		"xl/worksheets/sheet1.xml": `<?xml version="1.0"?><worksheet><sheetData><row r="3"><c r="B3"><v>1</v></c></row></sheetData></worksheet>`,
	})
	got, err := Compare(a, b)
	if err != nil {
		t.Fatal(err)
	}
	if got.Equal {
		t.Fatal("boolean t=b vs number 1 must be unequal")
	}
	found := false
	for _, d := range got.Diffs {
		if d.Axis == "types" && d.Location == "Sheet1!B3" {
			found = true
		}
	}
	if !found {
		t.Fatalf("want types Sheet1!B3, got %v", got.Diffs)
	}
	t.Logf("unequal: boolean vs number %v", got.Diffs)
}

func TestCompareUnequalOnFormulaText(t *testing.T) {
	a := mustXLSX(t, map[string]string{
		"xl/worksheets/sheet1.xml": `<?xml version="1.0"?><worksheet><sheetData><row r="1"><c r="C1"><f>A1+B1</f><v>15</v></c></row></sheetData></worksheet>`,
	})
	b := mustXLSX(t, map[string]string{
		"xl/worksheets/sheet1.xml": `<?xml version="1.0"?><worksheet><sheetData><row r="1"><c r="C1"><f>A1-B1</f><v>15</v></c></row></sheetData></worksheet>`,
	})
	got, err := Compare(a, b)
	if err != nil {
		t.Fatal(err)
	}
	if got.Equal {
		t.Fatal("A1+B1 vs A1-B1 must be unequal even with the same cached v")
	}
	found := false
	for _, d := range got.Diffs {
		if d.Axis == "formulas" && d.Location == "Sheet1!C1" {
			found = true
		}
	}
	if !found {
		t.Fatalf("want formulas Sheet1!C1, got %v", got.Diffs)
	}
	for _, d := range got.Diffs {
		if d.Axis == "values" {
			t.Fatalf("cached value matches; extra values diff: %v", got.Diffs)
		}
	}
	t.Logf("unequal: formula text %v", got.Diffs)
}

func TestCompareUnequalOnCachedFormulaValue(t *testing.T) {
	a := mustXLSX(t, map[string]string{
		"xl/worksheets/sheet1.xml": `<?xml version="1.0"?><worksheet><sheetData><row r="1"><c r="C1"><f>A1+B1</f><v>0</v></c></row></sheetData></worksheet>`,
	})
	b := mustXLSX(t, map[string]string{
		"xl/worksheets/sheet1.xml": `<?xml version="1.0"?><worksheet><sheetData><row r="1"><c r="C1"><f>A1+B1</f><v>15</v></c></row></sheetData></worksheet>`,
	})
	got, err := Compare(a, b)
	if err != nil {
		t.Fatal(err)
	}
	if got.Equal {
		t.Fatal("cached v=0 vs v=15 must be unequal")
	}
	found := false
	for _, d := range got.Diffs {
		if d.Axis == "values" && d.Location == "Sheet1!C1" {
			found = true
		}
	}
	if !found {
		t.Fatalf("want values Sheet1!C1, got %v", got.Diffs)
	}
	for _, d := range got.Diffs {
		if d.Axis == "formulas" {
			t.Fatalf("formula text matches; extra formulas diff: %v", got.Diffs)
		}
	}
	t.Logf("unequal: cached formula value %v", got.Diffs)
}

func TestCompareEqualOnSharedFormulaExpansion(t *testing.T) {
	shared := mustXLSX(t, map[string]string{
		"xl/worksheets/sheet1.xml": `<?xml version="1.0"?><worksheet><sheetData>` +
			`<row r="2"><c r="A2"><f t="shared" ref="A2:A3" si="0">B2</f><v>1</v></c></row>` +
			`<row r="3"><c r="A3"><f t="shared" si="0"/><v>2</v></c></row>` +
			`</sheetData></worksheet>`,
	})
	expanded := mustXLSX(t, map[string]string{
		"xl/worksheets/sheet1.xml": `<?xml version="1.0"?><worksheet><sheetData>` +
			`<row r="2"><c r="A2"><f>B2</f><v>1</v></c></row>` +
			`<row r="3"><c r="A3"><f>B3</f><v>2</v></c></row>` +
			`</sheetData></worksheet>`,
	})
	got, err := Compare(shared, expanded)
	if err != nil {
		t.Fatal(err)
	}
	if !got.Equal {
		t.Fatalf("shared B2 at A2/A3 must match expanded B2/B3, got %v", got.Diffs)
	}
	t.Log("equal: shared formula expansion B2 -> B3")
}

func TestCompareEqualOnSSTVsInlineStr(t *testing.T) {
	sst := mustXLSX(t, map[string]string{
		"xl/sharedStrings.xml":     `<sst><si><t>hello</t></si></sst>`,
		"xl/worksheets/sheet1.xml": `<?xml version="1.0"?><worksheet><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData></worksheet>`,
	})
	inline := mustXLSX(t, map[string]string{
		"xl/worksheets/sheet1.xml": sheetInline("hello"),
	})
	got, err := Compare(sst, inline)
	if err != nil {
		t.Fatal(err)
	}
	if !got.Equal {
		t.Fatalf("SST and inlineStr are the same string type, got %v", got.Diffs)
	}
}

func TestCompareFilesRoundtripFormulasInitVsGolden(t *testing.T) {
	_, file, _, ok := runtime.Caller(0)
	if !ok {
		t.Fatal("runtime.Caller failed")
	}
	dir := filepath.Join(filepath.Dir(file), "..", "..", "verification", "cases", "roundtrip", "formulas")
	got, err := CompareFiles(filepath.Join(dir, "init.xlsx"), filepath.Join(dir, "golden.xlsx"))
	if err != nil {
		t.Fatal(err)
	}
	if got.Equal {
		t.Fatal("formulas init v=0 vs golden cached values must be unequal")
	}
	found := false
	for _, d := range got.Diffs {
		if d.Axis == "values" && d.Location == "Sheet1!C1" {
			found = true
		}
		if d.Axis == "formulas" {
			t.Fatalf("formula text should match; got %v", got.Diffs)
		}
	}
	if !found {
		t.Fatalf("want values Sheet1!C1 (0 vs 15), got %v", got.Diffs)
	}
	t.Logf("unequal: roundtrip/formulas cached values %v", got.Diffs)
}

func TestCompareUnequalOnNumberFormat(t *testing.T) {
	percent := mustXLSX(t, map[string]string{
		"xl/styles.xml":            stylesXML(10),
		"xl/worksheets/sheet1.xml": `<?xml version="1.0"?><worksheet><sheetData><row r="1"><c r="A1" s="1"><v>0.25</v></c></row></sheetData></worksheet>`,
	})
	general := mustXLSX(t, map[string]string{
		"xl/styles.xml":            stylesXML(0),
		"xl/worksheets/sheet1.xml": `<?xml version="1.0"?><worksheet><sheetData><row r="1"><c r="A1" s="1"><v>0.25</v></c></row></sheetData></worksheet>`,
	})
	got, err := Compare(percent, general)
	if err != nil {
		t.Fatal(err)
	}
	if got.Equal {
		t.Fatal("numFmt 0.00% vs General must be unequal")
	}
	found := false
	for _, d := range got.Diffs {
		if d.Axis == "styles" && d.Location == "Sheet1!A1" {
			found = true
		}
	}
	if !found {
		t.Fatalf("want styles Sheet1!A1, got %v", got.Diffs)
	}
	t.Logf("unequal: number format %v", got.Diffs)
}

func TestCompareEqualOnSameResolvedStyleDifferentIndex(t *testing.T) {
	a := mustXLSX(t, map[string]string{
		"xl/styles.xml":            stylesXMLTwoPercentXFs(),
		"xl/worksheets/sheet1.xml": `<?xml version="1.0"?><worksheet><sheetData><row r="1"><c r="A1" s="1"><v>0.25</v></c></row></sheetData></worksheet>`,
	})
	b := mustXLSX(t, map[string]string{
		"xl/styles.xml":            stylesXMLTwoPercentXFs(),
		"xl/worksheets/sheet1.xml": `<?xml version="1.0"?><worksheet><sheetData><row r="1"><c r="A1" s="2"><v>0.25</v></c></row></sheetData></worksheet>`,
	})
	got, err := Compare(a, b)
	if err != nil {
		t.Fatal(err)
	}
	if !got.Equal {
		t.Fatalf("same resolved 0.00%% at different xf indexes must be equal, got %v", got.Diffs)
	}
}

func TestCompareEqualOnThemeSchemeFontTypeface(t *testing.T) {
	cell := `<?xml version="1.0"?><worksheet><sheetData><row r="1"><c r="A1" s="0"><v>1</v></c></row></sheetData></worksheet>`
	excelFont := `<fonts count="1"><font><sz val="11"/><color theme="1"/><name val="Aptos Narrow"/><scheme val="minor"/></font></fonts>`
	mogFont := `<fonts count="1"><font><sz val="11"/><color theme="1"/><name val="Calibri"/><scheme val="minor"/></font></fonts>`
	a := mustXLSX(t, map[string]string{
		"xl/styles.xml":            stylesXMLWith(excelFont, defaultFills, emptySidesBorder, defaultXFs),
		"xl/theme/theme1.xml":      themeXML("Office Theme", "Office", "000000", "FFFFFF"),
		"xl/worksheets/sheet1.xml": cell,
	})
	b := mustXLSX(t, map[string]string{
		"xl/styles.xml":            stylesXMLWith(mogFont, defaultFills, omittedSidesBorder, defaultXFs),
		"xl/theme/theme1.xml":      themeXML("Office Theme", "Office", "000000", "FFFFFF"),
		"xl/worksheets/sheet1.xml": cell,
	})
	got, err := Compare(a, b)
	if err != nil {
		t.Fatal(err)
	}
	if !got.Equal {
		t.Fatalf("scheme=minor Aptos Narrow vs Calibri must be equal, got %v", got.Diffs)
	}
}

func TestCompareUnequalOnAuthoredFontWithoutScheme(t *testing.T) {
	cell := `<?xml version="1.0"?><worksheet><sheetData><row r="1"><c r="A1" s="0"><v>1</v></c></row></sheetData></worksheet>`
	arial := `<fonts count="1"><font><sz val="11"/><name val="Arial"/></font></fonts>`
	calibri := `<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>`
	a := mustXLSX(t, map[string]string{
		"xl/styles.xml":            stylesXMLWith(arial, defaultFills, omittedSidesBorder, defaultXFs),
		"xl/worksheets/sheet1.xml": cell,
	})
	b := mustXLSX(t, map[string]string{
		"xl/styles.xml":            stylesXMLWith(calibri, defaultFills, omittedSidesBorder, defaultXFs),
		"xl/worksheets/sheet1.xml": cell,
	})
	got, err := Compare(a, b)
	if err != nil {
		t.Fatal(err)
	}
	if got.Equal {
		t.Fatal("Arial vs Calibri without scheme must be unequal")
	}
	found := false
	for _, d := range got.Diffs {
		if d.Axis == "styles" && d.Location == "Sheet1!A1" && strings.Contains(d.Detail, "font:") {
			found = true
		}
	}
	if !found {
		t.Fatalf("want styles font diff, got %v", got.Diffs)
	}
}

func TestCompareEqualOnExcelTintEncoding(t *testing.T) {
	cell := `<?xml version="1.0"?><worksheet><sheetData><row r="1"><c r="A1" s="1"><v>1</v></c></row></sheetData></worksheet>`
	excelTint := `<fonts count="2"><font><sz val="11"/><name val="Calibri"/><scheme val="minor"/></font>` +
		`<font><sz val="11"/><color theme="4" tint="0.19998779259620961"/><name val="Calibri"/><scheme val="minor"/></font></fonts>`
	mogTint := `<fonts count="2"><font><sz val="11"/><name val="Calibri"/><scheme val="minor"/></font>` +
		`<font><sz val="11"/><color theme="4" tint="0.2"/><name val="Calibri"/><scheme val="minor"/></font></fonts>`
	xfs := `<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>` +
		`<xf numFmtId="0" fontId="1" fillId="0" borderId="0" applyFont="1"/></cellXfs>`
	theme := themeXML("Office Theme", "Office", "000000", "FFFFFF")
	a := mustXLSX(t, map[string]string{
		"xl/styles.xml":            stylesXMLWith(excelTint, defaultFills, emptySidesBorder, xfs),
		"xl/theme/theme1.xml":      theme,
		"xl/worksheets/sheet1.xml": cell,
	})
	b := mustXLSX(t, map[string]string{
		"xl/styles.xml":            stylesXMLWith(mogTint, defaultFills, emptySidesBorder, xfs),
		"xl/theme/theme1.xml":      theme,
		"xl/worksheets/sheet1.xml": cell,
	})
	got, err := Compare(a, b)
	if err != nil {
		t.Fatal(err)
	}
	if !got.Equal {
		t.Fatalf("Excel tint 0.19998779259620961 vs 0.2 must be equal, got %v", got.Diffs)
	}
}

func TestCompareUnequalOnDifferentTint(t *testing.T) {
	cell := `<?xml version="1.0"?><worksheet><sheetData><row r="1"><c r="A1" s="1"><v>1</v></c></row></sheetData></worksheet>`
	tint02 := `<fonts count="2"><font><sz val="11"/><name val="Calibri"/><scheme val="minor"/></font>` +
		`<font><sz val="11"/><color theme="4" tint="0.2"/><name val="Calibri"/><scheme val="minor"/></font></fonts>`
	tint04 := `<fonts count="2"><font><sz val="11"/><name val="Calibri"/><scheme val="minor"/></font>` +
		`<font><sz val="11"/><color theme="4" tint="0.4"/><name val="Calibri"/><scheme val="minor"/></font></fonts>`
	xfs := `<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>` +
		`<xf numFmtId="0" fontId="1" fillId="0" borderId="0" applyFont="1"/></cellXfs>`
	theme := themeXML("Office Theme", "Office", "000000", "FFFFFF")
	a := mustXLSX(t, map[string]string{
		"xl/styles.xml":            stylesXMLWith(tint02, defaultFills, emptySidesBorder, xfs),
		"xl/theme/theme1.xml":      theme,
		"xl/worksheets/sheet1.xml": cell,
	})
	b := mustXLSX(t, map[string]string{
		"xl/styles.xml":            stylesXMLWith(tint04, defaultFills, emptySidesBorder, xfs),
		"xl/theme/theme1.xml":      theme,
		"xl/worksheets/sheet1.xml": cell,
	})
	got, err := Compare(a, b)
	if err != nil {
		t.Fatal(err)
	}
	if got.Equal {
		t.Fatal("tint 0.2 vs 0.4 must be unequal")
	}
	found := false
	for _, d := range got.Diffs {
		if d.Axis == "styles" && strings.Contains(d.Detail, "font:") {
			found = true
		}
	}
	if !found {
		t.Fatalf("want styles font/tint diff, got %v", got.Diffs)
	}
}

func TestCompareEqualOnEmptyBorderSides(t *testing.T) {
	cell := `<?xml version="1.0"?><worksheet><sheetData><row r="1"><c r="A1" s="0"><v>1</v></c></row></sheetData></worksheet>`
	font := `<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>`
	a := mustXLSX(t, map[string]string{
		"xl/styles.xml":            stylesXMLWith(font, defaultFills, emptySidesBorder, defaultXFs),
		"xl/worksheets/sheet1.xml": cell,
	})
	b := mustXLSX(t, map[string]string{
		"xl/styles.xml":            stylesXMLWith(font, defaultFills, omittedSidesBorder, defaultXFs),
		"xl/worksheets/sheet1.xml": cell,
	})
	got, err := Compare(a, b)
	if err != nil {
		t.Fatal(err)
	}
	if !got.Equal {
		t.Fatalf("empty <left/> sides vs omitted sides must be equal, got %v", got.Diffs)
	}
}

func TestCompareUnequalOnBorderStyle(t *testing.T) {
	cell := `<?xml version="1.0"?><worksheet><sheetData><row r="1"><c r="A1" s="0"><v>1</v></c></row></sheetData></worksheet>`
	font := `<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>`
	thin := `<borders count="1"><border><left style="thin"/><right/><top/><bottom/><diagonal/></border></borders>`
	a := mustXLSX(t, map[string]string{
		"xl/styles.xml":            stylesXMLWith(font, defaultFills, thin, defaultXFs),
		"xl/worksheets/sheet1.xml": cell,
	})
	b := mustXLSX(t, map[string]string{
		"xl/styles.xml":            stylesXMLWith(font, defaultFills, omittedSidesBorder, defaultXFs),
		"xl/worksheets/sheet1.xml": cell,
	})
	got, err := Compare(a, b)
	if err != nil {
		t.Fatal(err)
	}
	if got.Equal {
		t.Fatal("thin left border vs none must be unequal")
	}
	found := false
	for _, d := range got.Diffs {
		if d.Axis == "styles" && strings.Contains(d.Detail, "border:") {
			found = true
		}
	}
	if !found {
		t.Fatalf("want styles border diff, got %v", got.Diffs)
	}
}

func TestCompareEqualOnThemeDisplayNames(t *testing.T) {
	cell := `<?xml version="1.0"?><worksheet><sheetData><row r="1"><c r="A1" s="0"><v>1</v></c></row></sheetData></worksheet>`
	st := stylesXML(0)
	a := mustXLSX(t, map[string]string{
		"xl/styles.xml":            st,
		"xl/theme/theme1.xml":      themeXML("Office Theme", "Office", "000000", "FFFFFF"),
		"xl/worksheets/sheet1.xml": cell,
	})
	b := mustXLSX(t, map[string]string{
		"xl/styles.xml":            st,
		"xl/theme/theme1.xml":      themeXML("Office 2013 - 2022 Theme", "Office 2013 - 2022", "000000", "FFFFFF"),
		"xl/worksheets/sheet1.xml": cell,
	})
	got, err := Compare(a, b)
	if err != nil {
		t.Fatal(err)
	}
	if !got.Equal {
		t.Fatalf("theme display names must not fail, got %v", got.Diffs)
	}
	t.Log("equal: theme name= labels with matching color slots")
}

func TestCompareUnequalOnThemeSlotColor(t *testing.T) {
	cell := `<?xml version="1.0"?><worksheet><sheetData><row r="1"><c r="A1" s="0"><v>1</v></c></row></sheetData></worksheet>`
	st := stylesXML(0)
	a := mustXLSX(t, map[string]string{
		"xl/styles.xml":            st,
		"xl/theme/theme1.xml":      themeXML("Office Theme", "Office", "000000", "FFFFFF"),
		"xl/worksheets/sheet1.xml": cell,
	})
	b := mustXLSX(t, map[string]string{
		"xl/styles.xml":            st,
		"xl/theme/theme1.xml":      themeXML("Office Theme", "Office", "FF0000", "FFFFFF"),
		"xl/worksheets/sheet1.xml": cell,
	})
	got, err := Compare(a, b)
	if err != nil {
		t.Fatal(err)
	}
	if got.Equal {
		t.Fatal("theme slot color change must fail")
	}
	found := false
	for _, d := range got.Diffs {
		if d.Axis == "styles" {
			found = true
		}
	}
	if !found {
		t.Fatalf("want styles diff, got %v", got.Diffs)
	}
}

func TestCompareEqualOnDefaultRowHeightNoise(t *testing.T) {
	a := mustXLSX(t, map[string]string{
		"xl/worksheets/sheet1.xml": `<?xml version="1.0"?><worksheet><sheetFormatPr defaultRowHeight="16"/><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>`,
	})
	b := mustXLSX(t, map[string]string{
		"xl/worksheets/sheet1.xml": `<?xml version="1.0"?><worksheet><sheetFormatPr defaultRowHeight="15.5" defaultColWidth="10.6640625"/><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>`,
	})
	got, err := Compare(a, b)
	if err != nil {
		t.Fatal(err)
	}
	if !got.Equal {
		t.Fatalf("default row/col size rewrite must not fail, got %v", got.Diffs)
	}
}

func TestCompareUnequalOnSheetOrder(t *testing.T) {
	a := mustXLSX(t, map[string]string{
		"xl/workbook.xml": workbookSheets(`<sheet name="Sales" sheetId="1" r:id="rId1"/><sheet name="Expenses" sheetId="2" r:id="rId2"/>`),
	})
	b := mustXLSX(t, map[string]string{
		"xl/workbook.xml": workbookSheets(`<sheet name="Expenses" sheetId="1" r:id="rId1"/><sheet name="Sales" sheetId="2" r:id="rId2"/>`),
	})
	got, err := Compare(a, b)
	if err != nil {
		t.Fatal(err)
	}
	if got.Equal {
		t.Fatal("Sales/Expenses vs reverse order must be unequal")
	}
	found := false
	for _, d := range got.Diffs {
		if d.Axis == "sheets" {
			found = true
		}
	}
	if !found {
		t.Fatalf("want sheets diff, got %v", got.Diffs)
	}
	t.Logf("unequal: sheet order %v", got.Diffs)
}

func TestCompareEqualOnSheetIdAndRid(t *testing.T) {
	a := mustXLSX(t, map[string]string{
		"xl/workbook.xml": workbookSheets(`<sheet name="Sales" sheetId="1" r:id="rId1"/><sheet name="Expenses" sheetId="2" r:id="rId2"/>`),
	})
	b := mustXLSX(t, map[string]string{
		"xl/workbook.xml": workbookSheets(`<sheet name="Sales" sheetId="9" r:id="rId9"/><sheet name="Expenses" sheetId="3" r:id="rId3"/>`),
	})
	got, err := Compare(a, b)
	if err != nil {
		t.Fatal(err)
	}
	if !got.Equal {
		t.Fatalf("sheetId/rId must not fail, got %v", got.Diffs)
	}
}

func TestCompareUnequalOnHiddenSheet(t *testing.T) {
	a := mustXLSX(t, map[string]string{
		"xl/workbook.xml": workbookSheets(`<sheet name="Sheet1" sheetId="1" state="hidden" r:id="rId1"/><sheet name="Sheet2" sheetId="2" r:id="rId2"/>`),
	})
	b := mustXLSX(t, map[string]string{
		"xl/workbook.xml": workbookSheets(`<sheet name="Sheet1" sheetId="1" r:id="rId1"/><sheet name="Sheet2" sheetId="2" r:id="rId2"/>`),
	})
	got, err := Compare(a, b)
	if err != nil {
		t.Fatal(err)
	}
	if got.Equal {
		t.Fatal("hidden vs visible Sheet1 must be unequal")
	}
	found := false
	for _, d := range got.Diffs {
		if d.Axis == "sheets" {
			found = true
		}
	}
	if !found {
		t.Fatalf("want sheets diff, got %v", got.Diffs)
	}
}

func TestCompareUnequalOnDefinedName(t *testing.T) {
	a := mustXLSX(t, map[string]string{
		"xl/workbook.xml": workbookNames(`<definedName name="Answer">Sheet1!$A$1</definedName>`),
	})
	b := mustXLSX(t, map[string]string{
		"xl/workbook.xml": workbookNames(`<definedName name="Other">Sheet1!$A$1</definedName>`),
	})
	got, err := Compare(a, b)
	if err != nil {
		t.Fatal(err)
	}
	if got.Equal {
		t.Fatal("defined name Answer vs Other must be unequal")
	}
	found := false
	for _, d := range got.Diffs {
		if d.Axis == "names" {
			found = true
		}
	}
	if !found {
		t.Fatalf("want names diff, got %v", got.Diffs)
	}
	t.Logf("unequal: defined names %v", got.Diffs)
}

func TestCompareEqualOnDefinedNameOrder(t *testing.T) {
	a := mustXLSX(t, map[string]string{
		"xl/workbook.xml": workbookNames(`<definedName name="Answer">Sheet1!$A$1</definedName><definedName name="col1_">'Named Ranges'!$A$2:$A$6</definedName>`),
	})
	b := mustXLSX(t, map[string]string{
		"xl/workbook.xml": workbookNames(`<definedName name="col1_">'Named Ranges'!$A$2:$A$6</definedName><definedName name="Answer">Sheet1!$A$1</definedName>`),
	})
	got, err := Compare(a, b)
	if err != nil {
		t.Fatal(err)
	}
	if !got.Equal {
		t.Fatalf("definedName element order must not fail, got %v", got.Diffs)
	}
}

func TestCompareUnequalOnMerge(t *testing.T) {
	a := mustXLSX(t, map[string]string{
		"xl/worksheets/sheet1.xml": `<?xml version="1.0"?><worksheet><sheetData/><mergeCells count="1"><mergeCell ref="A1:B2"/></mergeCells></worksheet>`,
	})
	b := mustXLSX(t, map[string]string{
		"xl/worksheets/sheet1.xml": `<?xml version="1.0"?><worksheet><sheetData/></worksheet>`,
	})
	got, err := Compare(a, b)
	if err != nil {
		t.Fatal(err)
	}
	if got.Equal {
		t.Fatal("merge A1:B2 vs none must be unequal")
	}
	found := false
	for _, d := range got.Diffs {
		if d.Axis == "merges" && strings.Contains(d.Location, "A1:B2") {
			found = true
		}
	}
	if !found {
		t.Fatalf("want merges A1:B2, got %v", got.Diffs)
	}
	t.Logf("unequal: merge %v", got.Diffs)
}

func TestCompareEqualOnMergeOrder(t *testing.T) {
	a := mustXLSX(t, map[string]string{
		"xl/worksheets/sheet1.xml": `<?xml version="1.0"?><worksheet><sheetData/><mergeCells count="2"><mergeCell ref="A1:B2"/><mergeCell ref="C1:C2"/></mergeCells></worksheet>`,
	})
	b := mustXLSX(t, map[string]string{
		"xl/worksheets/sheet1.xml": `<?xml version="1.0"?><worksheet><sheetData/><mergeCells count="2"><mergeCell ref="C1:C2"/><mergeCell ref="A1:B2"/></mergeCells></worksheet>`,
	})
	got, err := Compare(a, b)
	if err != nil {
		t.Fatal(err)
	}
	if !got.Equal {
		t.Fatalf("mergeCell order must not fail, got %v", got.Diffs)
	}
}

func TestCompareUnequalOnFreeze(t *testing.T) {
	a := mustXLSX(t, map[string]string{
		"xl/worksheets/sheet1.xml": `<?xml version="1.0"?><worksheet><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><sheetData/></worksheet>`,
	})
	b := mustXLSX(t, map[string]string{
		"xl/worksheets/sheet1.xml": `<?xml version="1.0"?><worksheet><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetData/></worksheet>`,
	})
	got, err := Compare(a, b)
	if err != nil {
		t.Fatal(err)
	}
	if got.Equal {
		t.Fatal("frozen ySplit=1 vs none must be unequal")
	}
	found := false
	for _, d := range got.Diffs {
		if d.Axis == "freeze" {
			found = true
		}
	}
	if !found {
		t.Fatalf("want freeze diff, got %v", got.Diffs)
	}
	t.Logf("unequal: freeze %v", got.Diffs)
}

func TestCompareFilesFreezePanesInitVsGolden(t *testing.T) {
	_, file, _, ok := runtime.Caller(0)
	if !ok {
		t.Fatal("runtime.Caller failed")
	}
	dir := filepath.Join(filepath.Dir(file), "..", "..", "verification", "cases", "scratch", "freeze_panes")
	got, err := CompareFiles(filepath.Join(dir, "init.xlsx"), filepath.Join(dir, "golden.xlsx"))
	if err != nil {
		t.Fatal(err)
	}
	if got.Equal {
		t.Fatal("freeze_panes init vs golden must be unequal")
	}
	found := false
	for _, d := range got.Diffs {
		if d.Axis == "freeze" {
			found = true
		}
	}
	if !found {
		t.Fatalf("want freeze axis, got %v", got.Diffs)
	}
	t.Logf("unequal: scratch/freeze_panes %v", got.Diffs)
}

func TestCompareFilesRoundtripSimpleInitVsGolden(t *testing.T) {
	_, file, _, ok := runtime.Caller(0)
	if !ok {
		t.Fatal("runtime.Caller failed")
	}
	dir := filepath.Join(filepath.Dir(file), "..", "..", "verification", "cases", "roundtrip", "simple")
	got, err := CompareFiles(filepath.Join(dir, "init.xlsx"), filepath.Join(dir, "golden.xlsx"))
	if err != nil {
		t.Fatal(err)
	}
	if !got.Equal {
		t.Fatalf("roundtrip/simple init vs golden must match on values/date1904 (C4 92.3), got %v", got.Diffs)
	}
	t.Log("equal: roundtrip/simple init.xlsx vs golden.xlsx (resolved values including C4)")
}

func TestCompareFilesMissing(t *testing.T) {
	dir := t.TempDir()
	p := filepath.Join(dir, "a.xlsx")
	if err := os.WriteFile(p, mustXLSX(t, map[string]string{"xl/worksheets/sheet1.xml": sheetInline("x")}), 0o644); err != nil {
		t.Fatal(err)
	}
	if _, err := CompareFiles(p, filepath.Join(dir, "missing.xlsx")); err == nil {
		t.Fatal("missing golden should error")
	}
}

func workbookSheets(inner string) string {
	return `<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>` + inner + `</sheets>
</workbook>`
}

func workbookNames(inner string) string {
	return `<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
<definedNames>` + inner + `</definedNames>
</workbook>`
}

func workbookXML(date1904 string) string {
	return `<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<workbookPr calcId="1" date1904="` + date1904 + `"/>
<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>`
}

func workbookXML1904WithX15(date1904 string) string {
	return `<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">
<workbookPr calcId="1" date1904="` + date1904 + `"/>
<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
<extLst><ext uri="{140A7094-0E35-4892-8432-C4D2E57EDEB5}"><x15:workbookPr chartTrackingRefBase="1"/></ext></extLst>
</workbook>`
}

const (
	defaultFills       = `<fills count="1"><fill><patternFill patternType="none"/></fill></fills>`
	emptySidesBorder   = `<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>`
	omittedSidesBorder = `<borders count="1"><border></border></borders>`
	defaultXFs         = `<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs>`
)

func stylesXML(cellXfNumFmt int) string {
	return `<?xml version="1.0"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
		`<fonts count="1"><font><sz val="12"/><color theme="1"/><name val="Calibri"/></font></fonts>` +
		`<fills count="1"><fill><patternFill patternType="none"/></fill></fills>` +
		`<borders count="1"><border/></borders>` +
		`<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>` +
		`<xf numFmtId="` + itoa(cellXfNumFmt) + `" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/></cellXfs></styleSheet>`
}

func stylesXMLWith(fonts, fills, borders, xfs string) string {
	return `<?xml version="1.0"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
		fonts + fills + borders + xfs + `</styleSheet>`
}

func stylesXMLTwoPercentXFs() string {
	return `<?xml version="1.0"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
		`<fonts count="1"><font><sz val="12"/><name val="Calibri"/></font></fonts>` +
		`<fills count="1"><fill><patternFill patternType="none"/></fill></fills>` +
		`<borders count="1"><border/></borders>` +
		`<cellXfs count="3">` +
		`<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>` +
		`<xf numFmtId="10" fontId="0" fillId="0" borderId="0"/>` +
		`<xf numFmtId="10" fontId="0" fillId="0" borderId="0"/>` +
		`</cellXfs></styleSheet>`
}

func themeXML(themeName, schemeName, dk1, lt1 string) string {
	return `<?xml version="1.0"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="` + themeName + `">` +
		`<a:themeElements><a:clrScheme name="` + schemeName + `">` +
		`<a:dk1><a:srgbClr val="` + dk1 + `"/></a:dk1>` +
		`<a:lt1><a:srgbClr val="` + lt1 + `"/></a:lt1>` +
		`<a:dk2><a:srgbClr val="44546A"/></a:dk2>` +
		`<a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>` +
		`<a:accent1><a:srgbClr val="4472C4"/></a:accent1>` +
		`<a:accent2><a:srgbClr val="ED7D31"/></a:accent2>` +
		`<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>` +
		`<a:accent4><a:srgbClr val="FFC000"/></a:accent4>` +
		`<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>` +
		`<a:accent6><a:srgbClr val="70AD47"/></a:accent6>` +
		`<a:hlink><a:srgbClr val="0563C1"/></a:hlink>` +
		`<a:folHlink><a:srgbClr val="954F72"/></a:folHlink>` +
		`</a:clrScheme></a:themeElements></a:theme>`
}

func itoa(n int) string {
	if n == 0 {
		return "0"
	}
	return strconv.Itoa(n)
}

func sheetInline(value string) string {
	return `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>` + value + `</t></is></c></row></sheetData>
</worksheet>`
}

func mustXLSX(t *testing.T, parts map[string]string) []byte {
	t.Helper()
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	names := make([]string, 0, len(parts))
	for name := range parts {
		names = append(names, name)
	}
	sort.Strings(names)
	for _, name := range names {
		h := &zip.FileHeader{Name: name, Method: zip.Deflate, Modified: time.Time{}}
		w, err := zw.CreateHeader(h)
		if err != nil {
			t.Fatal(err)
		}
		if _, err := w.Write([]byte(parts[name])); err != nil {
			t.Fatal(err)
		}
	}
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	return buf.Bytes()
}
