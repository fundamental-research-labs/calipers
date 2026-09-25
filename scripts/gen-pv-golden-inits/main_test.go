package main

import (
	"archive/zip"
	"bytes"
	"os"
	"path/filepath"
	"reflect"
	"strings"
	"testing"

	"github.com/fundamental-research-labs/calipers/internal/cases"
	"github.com/fundamental-research-labs/calipers/internal/xlsxmodel"
)

func TestStripWorksheet(t *testing.T) {
	// Values marked REMOVE are caches. All other bytes must survive unchanged.
	// An array anchor deliberately follows a cached follower in document order.
	worksheet := `<?xml version="1.0"?>
<x:worksheet xmlns:x="` + spreadsheetNS + `" xmlns:e="urn:extension"><x:sheetData>
<x:row r="1">
 <x:c r="A1"><x:v>7</x:v></x:c>
 <x:c r="B1" t="s"><x:v>0</x:v></x:c>
 <x:c r="C1" t="inlineStr"><x:is><x:t>literal</x:t></x:is></x:c>
 <x:c r="D1" s="4"><x:f>SUM(A1,2)&amp;"test"</x:f>REMOVE<x:v>9</x:v>END</x:c>
 <x:c r="E1" t="b"><x:f>TRUE()</x:f>REMOVE<x:v>1</x:v>END</x:c>
 <x:c r="F1" t="e"><x:f>1/0</x:f>REMOVE<x:v>#DIV/0!</x:v>END</x:c>
 <x:c r="G1" t="str"><x:f>"hi"</x:f>REMOVE<x:v>hi</x:v>END</x:c>
 <x:c r="H1" t="inlineStr"><x:f>"hi"</x:f>REMOVE<x:is><x:t>hi</x:t></x:is>END</x:c>
 <x:c r="I1"><x:f>1</x:f></x:c>
 <x:c r="J1"><x:f t="shared" si="0" ref="J1:J2">A1</x:f>REMOVE<x:v>7</x:v>END</x:c>
</x:row><x:row r="2">
 <x:c r="J2"><x:f t="shared" si="0"/>REMOVE<x:v/>END</x:c>
 <x:c r="A2">REMOVE<x:v>20</x:v>END</x:c>
 <x:c r="B2"><x:f t="array" ref="$A$2:$B$3">SEQUENCE(2,2)</x:f>REMOVE<x:v>1</x:v>END</x:c>
 <x:c r="C2"><x:v>123</x:v></x:c>
 <x:c r="D2"><x:f t="dataTable" ref="D2:E3" dt2D="1" r1="A1" r2="B1"/>REMOVE<x:v>4</x:v>END</x:c>
</x:row><x:row r="3">
 <x:c r="A3" t="str">REMOVE<x:v>text</x:v>END</x:c>
 <x:c r="B3" t="e">REMOVE<x:v>#N/A</x:v>END</x:c>
 <x:c r="C3"><x:v>456</x:v></x:c>
 <x:c r="D3">REMOVE<x:v>5</x:v>END</x:c>
 <x:c r="E3">REMOVE<x:v>6</x:v>END</x:c>
</x:row></x:sheetData><x:extLst><e:c><e:f>extension</e:f><e:v>keep</e:v></e:c></x:extLst></x:worksheet>`
	input := strings.NewReplacer("REMOVE", "", "END", "").Replace(worksheet)
	want := worksheet
	for strings.Contains(want, "REMOVE") {
		start := strings.Index(want, "REMOVE")
		end := strings.Index(want[start:], "END") + start + 3
		want = want[:start] + want[end:]
	}
	got, err := stripWorksheet([]byte(input))
	if err != nil {
		t.Fatal(err)
	}
	if string(got) != want {
		t.Fatalf("only formula-result caches should change:\n%s\nwant:\n%s", got, want)
	}
	again, err := stripWorksheet(got)
	if err != nil || !bytes.Equal(got, again) {
		t.Fatalf("not idempotent: %v", err)
	}
}

func TestRejectMalformedWorksheetAndRanges(t *testing.T) {
	for _, input := range []string{`<worksheet>`, `<worksheet><broken></worksheet>`} {
		if _, err := stripWorksheet([]byte(input)); err == nil {
			t.Fatalf("accepted malformed XML %q", input)
		}
	}
	for _, ref := range []string{"", "A0", "XFE1", "A1048577", "B2:A1", "A1:B2:C3", "Sheet!A1"} {
		if _, err := parseRange(ref); err == nil {
			t.Fatalf("accepted range %q", ref)
		}
	}
}

func TestZipPreservesUnrelatedParts(t *testing.T) {
	var buf bytes.Buffer
	w := zip.NewWriter(&buf)
	if err := w.SetComment("original comment"); err != nil {
		t.Fatal(err)
	}
	original := map[string]string{
		"xl/worksheets/sheet1.xml": `<worksheet xmlns="` + spreadsheetNS + `"><sheetData><row r="1"><c r="A1"><f>1+1</f><v>2</v></c></row></sheetData></worksheet>`,
		"xl/workbook.xml":          `<workbook><calcPr calcMode="manual"/></workbook>`,
		"xl/calcChain.xml":         `<calcChain/>`,
		"xl/charts/chart1.xml":     `<c:v>cached chart point</c:v>`,
		"xl/media/image1.png":      "\x00\x01\xff",
	}
	// Fixed order for the package invariant check.
	for _, name := range []string{"xl/worksheets/sheet1.xml", "xl/workbook.xml", "xl/calcChain.xml", "xl/charts/chart1.xml", "xl/media/image1.png"} {
		dest, err := w.Create(name)
		if err != nil {
			t.Fatal(err)
		}
		if _, err := dest.Write([]byte(original[name])); err != nil {
			t.Fatal(err)
		}
	}
	if err := w.Close(); err != nil {
		t.Fatal(err)
	}
	got, err := withoutCaches(buf.Bytes())
	if err != nil {
		t.Fatal(err)
	}
	r, err := zip.NewReader(bytes.NewReader(got), int64(len(got)))
	if err != nil {
		t.Fatal(err)
	}
	if len(r.File) != len(original) || r.Comment != "original comment" {
		t.Fatal("package members/comment changed")
	}
	for _, f := range r.File {
		data, err := readPart(f)
		if err != nil {
			t.Fatal(err)
		}
		want := original[f.Name]
		if f.Name == "xl/worksheets/sheet1.xml" {
			want = strings.ReplaceAll(want, "<v>2</v>", "")
		}
		if string(data) != want {
			t.Errorf("changed %s: %s", f.Name, data)
		}
	}
	if err := equalParts(got, buf.Bytes()); err == nil {
		t.Fatal("cache-bearing input passed the integrity check")
	}
}

func TestCommittedPVGoldens(t *testing.T) {
	root := filepath.Join("..", "..", suiteDir)
	m, err := readManifest(root)
	if err != nil {
		t.Fatal(err)
	}
	if m.Repository != "https://github.com/lyfegame/shortcut" || m.Revision != "e351394e607b7b8b9e8791893abc052a0ead89f3" || m.Root != "packages/spreadsheet-verification/tests/fixtures/pv_goldens" {
		t.Fatal("source provenance changed")
	}
	corpus, err := cases.LoadCorpus(filepath.Dir(root))
	if err != nil {
		t.Fatal(err)
	}
	suite, err := cases.FilterSuite(corpus.Cases, "pv_goldens", corpus.Suites)
	if err != nil {
		t.Fatal(err)
	}
	if len(suite) != 107 || len(m.Fixtures) != 107 || len(cases.GoldenComparePass(suite)) != 107 {
		t.Fatal("expected all 107 cases in the verification pass")
	}
	byID := map[string]cases.Case{}
	for _, c := range suite {
		byID[c.Name] = c
	}
	counts := map[string]int{}
	for _, f := range m.Fixtures {
		t.Run(f.ID, func(t *testing.T) {
			c, ok := byID[f.ID]
			if !ok {
				t.Fatal("missing or duplicate manifest case")
			}
			delete(byID, f.ID)
			release := strings.Split(f.Golden, "/")[0]
			counts[release]++
			if f.Golden != release+"/"+f.ID+"/golden.xlsx" {
				t.Fatal("invalid source path")
			}
			if !c.Recalculate || c.RunScript() || !reflect.DeepEqual(c.Compare, xlsxmodel.Options{}) {
				t.Fatal("expected unscripted full recalculation with no comparison exceptions")
			}
			if err := generate(root, f, true); err != nil {
				t.Fatal(err)
			}
			// Exercise the actual semantic reader on every reference and stripped init.
			// A host that merely copies init must not pass when cached results existed.
			comparison, err := xlsxmodel.CompareFiles(c.InitPath, c.GoldenPath, c.Compare)
			if err != nil {
				t.Fatal(err)
			}
			if comparison.Equal {
				t.Fatal("copying init without recalculating must fail semantic comparison")
			}
			for _, name := range []string{"input.xlsx", "script.js", "golden.xlsx.meta.json"} {
				if _, err := os.Stat(filepath.Join(c.Dir, name)); !os.IsNotExist(err) {
					t.Errorf("unexpected %s (do not import inputs or invent capture metadata)", name)
				}
			}
		})
	}
	if len(byID) != 0 || !reflect.DeepEqual(counts, map[string]int{"pv23": 66, "pv26": 20, "pv26_2": 21}) {
		t.Fatalf("coverage: remaining=%d releases=%v", len(byID), counts)
	}
}
