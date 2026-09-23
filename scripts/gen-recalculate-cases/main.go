// Generate synthetic cache-free inputs for Mog issue #401. No Excel goldens
// are synthesized here; -check-goldens validates the Windows Excel results.
package main

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"flag"
	"fmt"
	"os"
	"path/filepath"
	"sort"
	"strconv"
	"strings"

	"github.com/fundamental-research-labs/calipers/internal/xlsxmodel"
)

type formula struct {
	cell, text string
	want       any
}
type spec struct {
	name     string
	formulas []formula
}

func specs() []spec {
	return []spec{
		{"offset_controls", []formula{
			{"H2", `COUNTA(OFFSET(A2:E2,0,0,1,3))`, 3},
			{"H3", `COUNTA(OFFSET(A2,,,,3))`, 3},
			{"H4", `MATCH("Q3",OFFSET(A10:E11,0,0,1,5),0)`, 4},
			{"H5", `VLOOKUP("k2",LookupBlock,4,FALSE)`, 30},
			{"H6", `INDEX(PeriodRow,,3)`, "FY"},
			{"H7", `COUNTIF(A2:C2,"FQ")`, 2},
			{"H8", `SUM(OFFSET(B11,0,0,2,2))`, 140},
		}},
		{"offset_named_bases", []formula{
			{"H2", `COUNTA(OFFSET(PeriodRange,0,0,1,3))`, 3},
			{"H3", `COUNTA(OFFSET(SingleCellBase,0,0,1,3))`, 3},
			{"H4", `COUNTA(OFFSET(PeriodRow,0,0,1,3))`, 3},
			{"H5", `SUM(OFFSET(DataStart,0,0,2,2))`, 140},
			{"H6", `SUM(OFFSET(DataStart,1,1,1,2))`, 130},
		}},
		{"offset_index_bases", []formula{
			{"H2", `COUNTA(OFFSET(INDEX(A2:E2,,1),,,,3))`, 3},
			{"H3", `COUNTA(OFFSET(INDEX(PeriodRange,,1),,,,3))`, 3},
			{"H4", `COUNTA(OFFSET(INDEX(PeriodRow,,1),,,,3))`, 3},
			{"H5", `SUM(OFFSET(INDEX(B11:E12,1,1),0,0,2,2))`, 140},
		}},
		{"offset_period_counts", []formula{
			{"A20", `COUNTIF(OFFSET(INDEX(PeriodRow,,1),,,,COLUMN()),INDEX(PeriodRow,,COLUMN()))`, 1},
			{"B20", `COUNTIF(OFFSET(INDEX(PeriodRow,,1),,,,COLUMN()),INDEX(PeriodRow,,COLUMN()))`, 2},
			{"C20", `COUNTIF(OFFSET(INDEX(PeriodRow,,1),,,,COLUMN()),INDEX(PeriodRow,,COLUMN()))`, 1},
			{"D20", `COUNTIF(OFFSET(INDEX(PeriodRow,,1),,,,COLUMN()),INDEX(PeriodRow,,COLUMN()))`, 3},
		}},
		{"offset_lookup", []formula{
			{"H2", `MATCH("Q3",OFFSET(LookupBlock,0,0,1,5),0)`, 4},
			{"H3", `MATCH("Q3",OFFSET(LookupBlock,0,0,1,COLUMNS(LookupBlock)),0)`, 4},
			{"H4", `IFERROR(VLOOKUP("k2",LookupBlock,MATCH("Q3",OFFSET(LookupBlock,0,0,1,COLUMNS(LookupBlock)),0),FALSE),"-")`, 30},
		}},
	}
}

const ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

func escape(s string) string {
	var b bytes.Buffer
	_ = xml.EscapeText(&b, []byte(s))
	return b.String()
}

func workbook(s spec) ([]byte, error) {
	rows := map[int][]string{}
	add := func(cell, body string) {
		i := strings.IndexFunc(cell, func(r rune) bool { return r >= '0' && r <= '9' })
		row, _ := strconv.Atoi(cell[i:])
		rows[row] = append(rows[row], `<c r="`+cell+`"`+body+`</c>`)
	}
	for row, vals := range map[int][]any{
		2:  {"FQ", "FQ", "FY", "FQ", "FY"},
		10: {"Key", "Q1", "Q2", "Q3", "Q4"},
		11: {"k2", 10, 20, 30, 40},
		12: {"k3", 50, 60, 70, 80},
	} {
		for col, v := range vals {
			cell := fmt.Sprintf("%c%d", 'A'+col, row)
			if text, ok := v.(string); ok {
				add(cell, ` t="inlineStr"><is><t>`+escape(text)+`</t></is>`)
			} else {
				add(cell, `><v>`+fmt.Sprint(v)+`</v>`)
			}
		}
	}
	for _, f := range s.formulas {
		add(f.cell, `><f>`+escape(f.text)+`</f>`)
	} // deliberately NO <v> or <is>
	var sheet strings.Builder
	sheet.WriteString(`<worksheet xmlns="` + ns + `"><sheetData>`)
	var keys []int
	for row := range rows {
		keys = append(keys, row)
	}
	sort.Ints(keys)
	for _, row := range keys {
		sort.Strings(rows[row])
		fmt.Fprintf(&sheet, `<row r="%d">%s</row>`, row, strings.Join(rows[row], ""))
	}
	sheet.WriteString(`</sheetData></worksheet>`)
	parts := map[string]string{
		"[Content_Types].xml":        `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`,
		"_rels/.rels":                `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
		"xl/_rels/workbook.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`,
		"xl/workbook.xml":            `<workbook xmlns="` + ns + `" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets><definedNames><definedName name="PeriodRange">S!$A$2:$E$2</definedName><definedName name="SingleCellBase">S!$A$2</definedName><definedName name="PeriodRow">S!$2:$2</definedName><definedName name="LookupBlock">S!$A$10:$E$12</definedName><definedName name="DataStart">S!$B$11</definedName></definedNames><calcPr calcMode="manual" fullCalcOnLoad="1" forceFullCalc="1"/></workbook>`,
		"xl/worksheets/sheet1.xml":   sheet.String(),
		"xl/styles.xml":              `<styleSheet xmlns="` + ns + `"><fonts count="1"><font><sz val="11"/><color rgb="FF000000"/><name val="Calibri"/><family val="2"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`,
	}
	var b bytes.Buffer
	zw := zip.NewWriter(&b)
	var names []string
	for n := range parts {
		names = append(names, n)
	}
	sort.Strings(names)
	for _, n := range names {
		// Store these tiny parts verbatim so fixtures are stable across Go
		// releases with different DEFLATE implementations.
		w, err := zw.CreateHeader(&zip.FileHeader{Name: n, Method: zip.Store})
		if err != nil {
			return nil, err
		}
		if _, err = w.Write([]byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + parts[n])); err != nil {
			return nil, err
		}
	}
	if err := zw.Close(); err != nil {
		return nil, err
	}
	return b.Bytes(), nil
}

func checkGolden(path string, s spec) error {
	data, err := os.ReadFile(path)
	if err != nil {
		return err
	}
	wb, err := xlsxmodel.Parse(data)
	if err != nil {
		return err
	}
	for _, f := range s.formulas {
		v := wb.Cells["S"][f.cell]
		good := v.Formula != ""
		switch want := f.want.(type) {
		case int:
			good = good && v.Type == xlsxmodel.TypeNumber && v.F == float64(want)
		case string:
			good = good && v.Type == xlsxmodel.TypeString && v.S == want
		}
		if !good {
			return fmt.Errorf("%s S!%s: want formula result %v, got %+v", s.name, f.cell, f.want, v)
		}
	}
	return nil
}

func run(root string, check bool) error {
	for _, s := range specs() {
		dir := filepath.Join(root, "recalculate", s.name)
		if check {
			if err := checkGolden(filepath.Join(dir, "golden.xlsx"), s); err != nil {
				return err
			}
			fmt.Println(s.name + ": expected results confirmed")
			continue
		}
		data, err := workbook(s)
		if err != nil {
			return err
		}
		if err := os.MkdirAll(dir, 0o755); err != nil {
			return err
		}
		if err := os.WriteFile(filepath.Join(dir, "init.xlsx"), data, 0o644); err != nil {
			return err
		}
	}
	return nil
}

func main() {
	root := flag.String("cases-dir", "verification/cases", "cases root")
	check := flag.Bool("check-goldens", false, "validate Windows-generated goldens against the synthetic fixture's expected results")
	flag.Parse()
	if err := run(*root, *check); err != nil {
		fmt.Fprintln(os.Stderr, err)
		os.Exit(1)
	}
}
