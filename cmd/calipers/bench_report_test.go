package main

import (
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"testing"
	"time"
)

func TestRunBenchReportHelp(t *testing.T) {
	out, err := captureStdout(t, func() error {
		return run([]string{"bench-report", "-h"})
	})
	if err != nil {
		t.Fatal(err)
	}
	low := strings.ToLower(out)
	for _, want := range []string{"us letter", "letter", "svg", "com", "office.js", "sequential", "save", "run", "export as pdf", "coverage"} {
		if !strings.Contains(low, want) {
			t.Fatalf("bench-report help missing %q:\n%s", want, out)
		}
	}
}

func TestRunBenchReportFixture(t *testing.T) {
	jsonPath := writeBenchFixture(t, t.TempDir(), dualEngineFixture(3))
	outDir := t.TempDir()
	if err := run([]string{"bench-report", "--json", jsonPath, "--out", outDir}); err != nil {
		t.Fatal(err)
	}
	html := readFile(t, filepath.Join(outDir, benchReportHTMLName))
	assertBenchHTMLCommon(t, html)
	assertChartSVG(t, readFile(t, filepath.Join(outDir, "speed.svg")))
	assertChartSVG(t, readFile(t, filepath.Join(outDir, "memory.svg")))
	if strings.Contains(html, "Excel COM series was not collected") {
		t.Fatal("dual-engine fixture should include Excel COM")
	}
}

func TestRunBenchReportThousandCases(t *testing.T) {
	jsonPath := writeBenchFixture(t, t.TempDir(), dualEngineFixture(1000))
	outDir := t.TempDir()
	if err := run([]string{"bench-report", "--json", jsonPath, "--out", outDir}); err != nil {
		t.Fatal(err)
	}
	html := readFile(t, filepath.Join(outDir, benchReportHTMLName))
	assertBenchHTMLCommon(t, html)
	nPages := strings.Count(html, `class="page"`)
	if nPages < 20 {
		t.Fatalf("1000-row table must paginate US Letter; page count = %d", nPages)
	}
	nSVG := strings.Count(html, "<svg")
	if nSVG != 2 {
		t.Fatalf("want exactly two per-task plots, svg count = %d", nSVG)
	}
	if strings.Count(html, "<tr") < 1000 {
		t.Fatalf("compact table must list the cases; tr count = %d", strings.Count(html, "<tr"))
	}
	speed := readFile(t, filepath.Join(outDir, "speed.svg"))
	mem := readFile(t, filepath.Join(outDir, "memory.svg"))
	assertChartSVG(t, speed)
	assertChartSVG(t, mem)
	if strings.Count(speed, "M") < 500 {
		t.Fatalf("speed plot must have one mark per task, path commands = %d", strings.Count(speed, "M"))
	}
}

func TestRunBenchReportCoverage(t *testing.T) {
	jsonPath := writeBenchFixture(t, t.TempDir(), mogOnlyFixture(4))
	covPath := filepath.Join(t.TempDir(), "coverage.json")
	cov := `{
  "version": 1,
  "generatedAt": "2026-09-11T00:00:00Z",
  "source": {"title": "Microsoft Excel JavaScript API", "url": "https://learn.microsoft.com/javascript/api/excel", "note": "test"},
  "summary": {
    "classes": 2,
    "objectModelClasses": 1,
    "methods": 3,
    "objectModelMethods": 2,
    "objectModelProperties": 1,
    "events": 0,
    "functionsClassMethods": 1,
    "implementedObjectModelMethods": 1,
    "verifiedObjectModelMethods": 1,
    "implementedObjectModelProperties": 1,
    "verifiedObjectModelProperties": 0,
    "formulasInScripts": 1
  },
  "runtime": [{"id": "Excel.run", "implemented": true, "cases": ["officejs/table_add"]}],
  "formulas": [{"name": "SUM", "cases": ["officejs/range_formulas_sum"]}],
  "classes": [
    {
      "id": "Excel.Range",
      "family": "object-model",
      "methodCount": 2,
      "propertyCount": 1,
      "eventCount": 0,
      "implementedMethods": 1,
      "verifiedMethods": 1,
      "implementedProperties": 1,
      "verifiedProperties": 0,
      "members": [
        {"id": "Excel.Range.clear", "name": "clear", "kind": "method", "signature": "clear(applyTo)", "apiSet": "ExcelApi 1.1", "preview": false, "implemented": true, "cases": ["officejs/range_clear"]},
        {"id": "Excel.Range.getCell", "name": "getCell", "kind": "method", "signature": "getCell(row, column)", "apiSet": "ExcelApi 1.1", "preview": false, "implemented": false, "cases": []},
        {"id": "Excel.Range.values", "name": "values", "kind": "property", "signature": "values", "apiSet": "ExcelApi 1.1", "preview": false, "implemented": true, "cases": []}
      ]
    },
    {
      "id": "Excel.Functions",
      "family": "functions",
      "methodCount": 1,
      "propertyCount": 0,
      "eventCount": 0,
      "implementedMethods": 0,
      "verifiedMethods": 0,
      "implementedProperties": 0,
      "verifiedProperties": 0,
      "members": [
        {"id": "Excel.Functions.sum", "name": "sum", "kind": "method", "signature": "sum(values)", "apiSet": "ExcelApi 1.2", "preview": false, "implemented": false, "cases": []}
      ]
    }
  ]
}`
	if err := os.WriteFile(covPath, []byte(cov), 0o644); err != nil {
		t.Fatal(err)
	}
	outDir := t.TempDir()
	if err := run([]string{"bench-report", "--json", jsonPath, "--out", outDir, "--coverage", covPath}); err != nil {
		t.Fatal(err)
	}
	html := readFile(t, filepath.Join(outDir, benchReportHTMLName))
	assertBenchHTMLCommon(t, html)
	for _, want := range []string{
		"Office.js Excel API coverage",
		"Excel.Range",
		"clear(applyTo)",
		"Excel.Functions",
		"object-model methods",
		"Defined in Mog",
		"Verified in this corpus",
		"SUM",
	} {
		if !strings.Contains(html, want) {
			t.Fatalf("coverage HTML missing %q", want)
		}
	}
	copied := readFile(t, filepath.Join(outDir, coverageJSONName))
	if !strings.Contains(copied, `"Excel.Range"`) {
		t.Fatal("report dir must keep a copy of the coverage JSON")
	}
}

func TestRunBenchReportMogOnlyFixture(t *testing.T) {
	jsonPath := writeBenchFixture(t, t.TempDir(), mogOnlyFixture(8))
	outDir := t.TempDir()
	if err := run([]string{"bench-report", "--json", jsonPath, "--out", outDir}); err != nil {
		t.Fatal(err)
	}
	html := readFile(t, filepath.Join(outDir, benchReportHTMLName))
	assertBenchHTMLCommon(t, html)
	if !strings.Contains(html, "Excel COM series was not collected") {
		t.Fatal("mog-only report must explain missing Excel")
	}
	speed := readFile(t, filepath.Join(outDir, "speed.svg"))
	assertChartSVG(t, speed)
	assertChartSVG(t, readFile(t, filepath.Join(outDir, "memory.svg")))
	if !strings.Contains(speed, ">Excel</text>") || !strings.Contains(speed, ">Mog</text>") {
		t.Fatal("legend must name Excel and Mog even when Excel has no points")
	}
	if !strings.Contains(speed, excelColor) || !strings.Contains(speed, mogColor) {
		t.Fatalf("legend must use saturated Excel/Mog colors %s %s", excelColor, mogColor)
	}
}

func TestRunBenchReportMemoryAxisRoundUnits(t *testing.T) {
	doc := mogOnlyFixture(6)
	for i := range doc.Cases {
		r := doc.Cases[i].Results["mog"]
		if i%2 == 0 {
			r.PeakBytes = 40 << 20
		} else {
			r.PeakBytes = 400 << 20
		}
		doc.Cases[i].Results["mog"] = r
	}
	jsonPath := writeBenchFixture(t, t.TempDir(), doc)
	outDir := t.TempDir()
	if err := run([]string{"bench-report", "--json", jsonPath, "--out", outDir}); err != nil {
		t.Fatal(err)
	}
	svg := readFile(t, filepath.Join(outDir, "memory.svg"))
	if strings.Contains(svg, "953") || strings.Contains(svg, ".7 MB") {
		t.Fatalf("memory axis used 10^n-byte ticks labeled as MB:\n%s", svg)
	}
	ok := false
	for _, want := range []string{"10 MB", "20 MB", "50 MB", "100 MB", "200 MB", "500 MB", "1 GB"} {
		if strings.Contains(svg, want) {
			ok = true
			break
		}
	}
	if !ok {
		t.Fatalf("memory axis want round 1/2/5 MB or GB ticks:\n%s", svg)
	}
}

func assertBenchHTMLCommon(t *testing.T, html string) {
	t.Helper()
	if strings.TrimSpace(html) == "" {
		t.Fatal("empty HTML")
	}
	low := strings.ToLower(html)
	for _, want := range []string{
		"@page",
		"8.5in",
		"11in",
		"page-break",
		"Export as PDF",
		"window.print",
		"Excel engine verifier",
		"Excel.Application",
		"COM",
		"sideloaded",
		"Office.js",
		"add-in",
		"AppSource",
		"Office Scripts",
		"sequential",
		"same Windows machine",
		"save",
		"run",
		"wall time",
		"peak working set",
		"<svg",
		"<table",
	} {
		hay, needle := html, want
		if !strings.HasPrefix(want, "<") && !strings.HasPrefix(want, "@") {
			hay, needle = low, strings.ToLower(want)
		}
		if !strings.Contains(hay, needle) {
			t.Fatalf("HTML missing %q", want)
		}
	}
	if strings.Contains(low, "calipers") {
		t.Fatal("HTML must not mention calipers")
	}
	if strings.Contains(low, "median") {
		t.Fatal("HTML must not summarize with a median across unlike tasks")
	}
	if strings.Contains(html, "no single average") || strings.Contains(html, "Each mark below is one task") {
		t.Fatal("HTML must not apologize for the plot shape")
	}
	if strings.Contains(html, `type="module"`) || strings.Contains(html, "import ") {
		t.Fatal("HTML must not use ES modules (file:// Print-to-PDF)")
	}
}

func assertChartSVG(t *testing.T, svg string) {
	t.Helper()
	if !strings.Contains(svg, "<svg") {
		t.Fatal("not an SVG")
	}
	if !strings.Contains(svg, "<path") && !strings.Contains(svg, "<circle") && !strings.Contains(svg, "<line") {
		t.Fatalf("SVG has no axes/series/points:\n%s", svg[:min(len(svg), 400)])
	}
	if !strings.Contains(svg, "<text") {
		t.Fatal("SVG chart must label axes or series")
	}
}

func dualEngineFixture(n int) BenchFile {
	now := time.Date(2026, 9, 11, 12, 0, 0, 0, time.UTC)
	doc := BenchFile{
		Version:     1,
		GeneratedAt: now,
		StartedAt:   now,
		EndedAt:     now.Add(time.Hour),
		Host:        BenchHost{OS: "windows", Arch: "amd64", GOOS: "windows"},
		Engines: []BenchEngine{
			{ID: "excel", Kind: "excel-com", Spec: "excel"},
			{ID: "mog", Kind: "binary", Spec: `C:\mog.exe`},
		},
	}
	suites := []string{"roundtrip", "default", "scratch", "officejs"}
	doc.Cases = make([]BenchCase, n)
	for i := 0; i < n; i++ {
		suite := suites[i%len(suites)]
		id := suite + "/case_" + strconv.Itoa(i)
		exMs := int64(200 + (i%17)*40 + (i%3)*250)
		mogMs := int64(80 + (i%13)*25)
		if i%41 == 0 {
			mogMs = exMs * 4
		}
		exMem := int64(40<<20) + int64(i%9)*1<<20
		mogMem := int64(18<<20) + int64(i%5)*1<<20
		start := now.Add(time.Duration(i) * 20 * time.Millisecond)
		err := ""
		if i == 7 {
			err = "engine: boom"
		}
		doc.Cases[i] = BenchCase{
			ID:     id,
			Suite:  suite,
			Script: i%4 == 0,
			Results: map[string]BenchResult{
				"excel": {DurationMs: exMs, PeakBytes: exMem, StartedAt: start, EndedAt: start.Add(time.Duration(exMs) * time.Millisecond)},
				"mog":   {DurationMs: mogMs, PeakBytes: mogMem, StartedAt: start.Add(time.Second), EndedAt: start.Add(time.Second + time.Duration(mogMs)*time.Millisecond), Error: err},
			},
		}
	}
	return doc
}

func mogOnlyFixture(n int) BenchFile {
	now := time.Date(2026, 9, 11, 12, 0, 0, 0, time.UTC)
	doc := BenchFile{
		Version:     1,
		GeneratedAt: now,
		StartedAt:   now,
		EndedAt:     now.Add(time.Minute),
		Host:        BenchHost{OS: "linux", Arch: "arm64", GOOS: "linux"},
		Engines:     []BenchEngine{{ID: "mog", Kind: "binary", Spec: "/path/to/mog"}},
	}
	doc.Cases = make([]BenchCase, n)
	for i := 0; i < n; i++ {
		start := now.Add(time.Duration(i) * time.Second)
		doc.Cases[i] = BenchCase{
			ID:    "roundtrip/c" + strconv.Itoa(i),
			Suite: "roundtrip",
			Results: map[string]BenchResult{
				"mog": {DurationMs: int64(100 + i*10), PeakBytes: int64(20<<20) + int64(i)<<18, StartedAt: start, EndedAt: start.Add(120 * time.Millisecond)},
			},
		}
	}
	return doc
}

func writeBenchFixture(t *testing.T, dir string, doc BenchFile) string {
	t.Helper()
	path := filepath.Join(dir, "bench.json")
	if err := writeBenchJSON(path, doc); err != nil {
		t.Fatal(err)
	}
	return path
}

func readFile(t *testing.T, path string) string {
	t.Helper()
	b, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	return string(b)
}
