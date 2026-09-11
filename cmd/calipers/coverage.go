package main

import (
	"encoding/json"
	"fmt"
	"html"
	"os"
	"path/filepath"
	"strings"
)

const (
	coverageJSONName          = "officejs-coverage.json"
	coverageClassRowsPerPage  = 28
	coverageMethodRowsPerPage = 32
	coverageFnPerPage         = 96
)

// CoverageFile is the JSON produced by scripts/officejs-coverage.
type CoverageFile struct {
	Version     int               `json:"version"`
	GeneratedAt string            `json:"generatedAt"`
	Source      CoverageSource    `json:"source"`
	Summary     CoverageSummary   `json:"summary"`
	Runtime     []CoverageRuntime `json:"runtime"`
	Formulas    []CoverageFormula `json:"formulas"`
	Classes     []CoverageClass   `json:"classes"`
}

type CoverageSource struct {
	Title string `json:"title"`
	URL   string `json:"url"`
	Note  string `json:"note"`
}

type CoverageSummary struct {
	Classes                          int `json:"classes"`
	ObjectModelClasses               int `json:"objectModelClasses"`
	Methods                          int `json:"methods"`
	ObjectModelMethods               int `json:"objectModelMethods"`
	ObjectModelProperties            int `json:"objectModelProperties"`
	Events                           int `json:"events"`
	FunctionsClassMethods            int `json:"functionsClassMethods"`
	ImplementedObjectModelMethods    int `json:"implementedObjectModelMethods"`
	VerifiedObjectModelMethods       int `json:"verifiedObjectModelMethods"`
	ImplementedObjectModelProperties int `json:"implementedObjectModelProperties"`
	VerifiedObjectModelProperties    int `json:"verifiedObjectModelProperties"`
	FormulasInScripts                int `json:"formulasInScripts"`
}

type CoverageRuntime struct {
	ID          string   `json:"id"`
	Implemented bool     `json:"implemented"`
	Cases       []string `json:"cases"`
}

type CoverageFormula struct {
	Name  string   `json:"name"`
	Cases []string `json:"cases"`
}

type CoverageClass struct {
	ID                    string           `json:"id"`
	Family                string           `json:"family"`
	MethodCount           int              `json:"methodCount"`
	PropertyCount         int              `json:"propertyCount"`
	EventCount            int              `json:"eventCount"`
	ImplementedMethods    int              `json:"implementedMethods"`
	VerifiedMethods       int              `json:"verifiedMethods"`
	ImplementedProperties int              `json:"implementedProperties"`
	VerifiedProperties    int              `json:"verifiedProperties"`
	Members               []CoverageMember `json:"members"`
}

type CoverageMember struct {
	ID          string   `json:"id"`
	Name        string   `json:"name"`
	Kind        string   `json:"kind"`
	Signature   string   `json:"signature"`
	APISet      string   `json:"apiSet"`
	Preview     bool     `json:"preview"`
	Implemented bool     `json:"implemented"`
	Cases       []string `json:"cases"`
}

func loadCoverage(path string) (*CoverageFile, error) {
	if strings.TrimSpace(path) == "" {
		return nil, nil
	}
	data, err := os.ReadFile(path)
	if err != nil {
		return nil, err
	}
	var doc CoverageFile
	if err := json.Unmarshal(data, &doc); err != nil {
		return nil, fmt.Errorf("coverage json: %w", err)
	}
	if len(doc.Classes) == 0 {
		return nil, fmt.Errorf("coverage json has no classes")
	}
	return &doc, nil
}

func coverageCSS() string {
	return `
.cov-hero { display: grid; grid-template-columns: 1fr 1fr; gap: 0.7em; margin: 0 0 0.7em; }
.cov-card {
  border: 1px solid #d8dee4; border-radius: 4px; padding: 0.45em 0.6em;
  background: #f7f9fb;
}
.cov-card h3 { margin: 0 0 0.2em; font-size: 9.5pt; font-weight: 600; }
.cov-card .num { font-size: 16pt; font-weight: 700; letter-spacing: -0.02em; }
.cov-card .den { font-size: 9.5pt; color: #5a6570; font-weight: 500; }
.meter { height: 8px; background: #e4eaef; border-radius: 4px; overflow: hidden; margin: 0.35em 0 0.15em; }
.meter > span { display: block; height: 100%; }
.meter-verified > span { background: #00B050; }
.meter-implemented > span { background: #FF3B00; }
.cov-legend { font-size: 8.5pt; color: #5a6570; margin: 0 0 0.55em; }
.pill { display: inline-block; font-size: 7.5pt; font-weight: 600; padding: 0.05em 0.38em; border-radius: 3px; margin-right: 0.2em; }
.pill-yes { background: #e4f7ea; color: #0b6b2c; }
.pill-no { background: #f3f5f7; color: #6a7380; }
.pill-impl { background: #ffe8dc; color: #9a3400; }
.pill-preview { background: #fff3cd; color: #6b5300; }
.member-grid { column-count: 3; column-gap: 0.7em; font-size: 7.5pt; line-height: 1.35; }
.member-grid .fn { break-inside: avoid; }
.filterbar { display: flex; gap: 0.4em; flex-wrap: wrap; margin: 0 0 0.45em; }
.filterbar button {
  font: 600 8.5pt/1.2 "Segoe UI", Helvetica, Arial, sans-serif;
  background: #eef2f5; color: #1f2a33; border: 0; border-radius: 3px;
  padding: 0.25em 0.55em; cursor: pointer;
}
.filterbar button.on { background: #1f2a33; color: #fff; }
tr.cov-hidden { display: none; }
@media print { tr.cov-hidden { display: table-row; } .filterbar { display: none !important; } }
`
}

func pct(num, den int) float64 {
	if den <= 0 {
		return 0
	}
	return 100 * float64(num) / float64(den)
}

func meter(kind string, num, den int) string {
	width := pct(num, den)
	if width < 0 {
		width = 0
	}
	if width > 100 {
		width = 100
	}
	return fmt.Sprintf(
		`<div class="meter meter-%s"><span style="width:%.1f%%"></span></div>`,
		html.EscapeString(kind), width,
	)
}

func coverageFirstPageNote(cov *CoverageFile) string {
	if cov == nil {
		return ""
	}
	s := cov.Summary
	return fmt.Sprintf(
		`<p class="meta">Office.js Excel API · <strong>%d</strong> object-model methods · <strong>%d</strong> verified in this corpus · <strong>%d</strong> defined in Mog · plus <strong>%d</strong> <code>Excel.Functions</code> worksheet wrappers</p>`,
		s.ObjectModelMethods, s.VerifiedObjectModelMethods, s.ImplementedObjectModelMethods, s.FunctionsClassMethods,
	)
}

func renderCoveragePages(cov *CoverageFile) string {
	if cov == nil {
		return ""
	}
	var b strings.Builder
	s := cov.Summary
	b.WriteString(`<section class="page">
<h2>Office.js Excel API coverage</h2>
`)
	fmt.Fprintf(&b, `<p class="caption">Microsoft’s Excel JavaScript API is the contract. Object-model methods are Range, Worksheet, Table, Chart, and the rest of the workbook types. <code>Excel.Functions</code> is a separate family of worksheet-function wrappers (<code>workbook.functions.sum</code>); Mog evaluates those as cell formulas through <code>Range.formulas</code>.</p>`)
	b.WriteString(`<div class="cov-hero">`)
	fmt.Fprintf(&b, `<div class="cov-card"><h3>Verified in this corpus</h3><div><span class="num">%d</span> <span class="den">/ %d methods</span></div>%s<p class="cov-legend">A verification script calls the method. %.0f%% of the object model.</p></div>`,
		s.VerifiedObjectModelMethods, s.ObjectModelMethods,
		meter("verified", s.VerifiedObjectModelMethods, s.ObjectModelMethods),
		pct(s.VerifiedObjectModelMethods, s.ObjectModelMethods))
	fmt.Fprintf(&b, `<div class="cov-card"><h3>Defined in Mog</h3><div><span class="num">%d</span> <span class="den">/ %d methods</span></div>%s<p class="cov-legend">The headless Office.js host exposes the method. %.0f%% of the object model.</p></div>`,
		s.ImplementedObjectModelMethods, s.ObjectModelMethods,
		meter("implemented", s.ImplementedObjectModelMethods, s.ObjectModelMethods),
		pct(s.ImplementedObjectModelMethods, s.ObjectModelMethods))
	b.WriteString(`</div>`)
	fmt.Fprintf(&b, `<p class="meta">%s · %d classes · %d object-model properties · %d events · %d <code>Excel.Functions</code> methods · %d worksheet formulas written via <code>Range.formulas</code> in corpus scripts</p>`,
		html.EscapeString(cov.Source.Title),
		s.ObjectModelClasses, s.ObjectModelProperties, s.Events, s.FunctionsClassMethods, s.FormulasInScripts)
	if len(cov.Formulas) > 0 {
		names := make([]string, 0, len(cov.Formulas))
		for _, f := range cov.Formulas {
			names = append(names, f.Name)
		}
		fmt.Fprintf(&b, `<p class="caption">Formulas in Office.js scripts: %s</p>`, html.EscapeString(strings.Join(names, ", ")))
	}
	b.WriteString(`<h2>Classes with coverage</h2>
<p class="caption">Mog = method is defined on the host. Corpus = a verification <code>script.js</code> calls it.</p>
<table><thead><tr><th>Class</th><th>Methods</th><th>Mog</th><th>Corpus</th><th>Props Mog / corpus</th></tr></thead><tbody>
`)
	covered := make([]CoverageClass, 0)
	rest := make([]CoverageClass, 0)
	var functions *CoverageClass
	for i := range cov.Classes {
		c := cov.Classes[i]
		if c.Family == "functions" {
			cp := c
			functions = &cp
			continue
		}
		if c.VerifiedMethods > 0 || c.ImplementedMethods > 0 || c.VerifiedProperties > 0 || c.ImplementedProperties > 0 {
			covered = append(covered, c)
		} else {
			rest = append(rest, c)
		}
	}
	rows := coverageClassRows(covered)
	// First page holds the hero + as many class rows as fit (~12).
	firstRows := 12
	if firstRows > len(rows) {
		firstRows = len(rows)
	}
	for _, row := range rows[:firstRows] {
		b.WriteString(row)
	}
	b.WriteString("</tbody></table></section>\n")

	remain := rows[firstRows:]
	pages := chunkStrings(remain, coverageClassRowsPerPage)
	for i, page := range pages {
		b.WriteString(`<section class="page">`)
		fmt.Fprintf(&b, `<h2>Classes with coverage (%d/%d)</h2>`, i+2, len(pages)+1)
		b.WriteString(`<table><thead><tr><th>Class</th><th>Methods</th><th>Mog</th><th>Corpus</th><th>Props Mog / corpus</th></tr></thead><tbody>`)
		for _, row := range page {
			b.WriteString(row)
		}
		b.WriteString("</tbody></table></section>\n")
	}

	b.WriteString(coverageMethodPages(covered))
	if functions != nil {
		b.WriteString(coverageFunctionsPages(*functions))
	}
	b.WriteString(coverageRestPages(rest))
	return b.String()
}

func coverageClassRows(classes []CoverageClass) []string {
	rows := make([]string, 0, len(classes))
	for _, c := range classes {
		rows = append(rows, fmt.Sprintf(
			`<tr><td>%s</td><td>%d</td><td>%d</td><td>%d</td><td>%d / %d</td></tr>`+"\n",
			html.EscapeString(c.ID), c.MethodCount, c.ImplementedMethods, c.VerifiedMethods,
			c.ImplementedProperties, c.VerifiedProperties,
		))
	}
	return rows
}

func coverageMethodPages(classes []CoverageClass) string {
	type row struct {
		cls CoverageClass
		m   CoverageMember
	}
	var items []row
	for _, c := range classes {
		for _, m := range c.Members {
			if m.Kind != "method" {
				continue
			}
			if !m.Implemented && len(m.Cases) == 0 {
				continue
			}
			items = append(items, row{c, m})
		}
	}
	if len(items) == 0 {
		return ""
	}
	pages := (len(items) + coverageMethodRowsPerPage - 1) / coverageMethodRowsPerPage
	var b strings.Builder
	for p := 0; p < pages; p++ {
		start := p * coverageMethodRowsPerPage
		end := start + coverageMethodRowsPerPage
		if end > len(items) {
			end = len(items)
		}
		b.WriteString(`<section class="page">`)
		if p == 0 {
			b.WriteString(`<h2>Object-model methods in Mog or the corpus</h2>`)
			b.WriteString(`<p class="caption">Each row is one Microsoft Excel JavaScript API method. Corpus lists verification case ids.</p>`)
			b.WriteString(`<div class="filterbar" id="cov-filter">
<button type="button" data-filter="all" class="on">All</button>
<button type="button" data-filter="verified">Corpus</button>
<button type="button" data-filter="implemented">Mog</button>
<button type="button" data-filter="both">Mog + corpus</button>
</div>`)
		} else {
			fmt.Fprintf(&b, `<h2>Object-model methods in Mog or the corpus (%d/%d)</h2>`, p+1, pages)
		}
		b.WriteString(`<table class="cov-methods"><thead><tr><th>Method</th><th>API set</th><th>Mog</th><th>Corpus</th></tr></thead><tbody>`)
		for _, it := range items[start:end] {
			cls := ""
			if len(it.m.Cases) > 0 {
				cls += " verified"
			}
			if it.m.Implemented {
				cls += " implemented"
			}
			mog := `<span class="pill pill-no">no</span>`
			if it.m.Implemented {
				mog = `<span class="pill pill-impl">yes</span>`
			}
			corpus := `<span class="pill pill-no">—</span>`
			if len(it.m.Cases) > 0 {
				corpus = `<span class="pill pill-yes">` + html.EscapeString(strconvMin(len(it.m.Cases), 99)) + `</span> ` + html.EscapeString(joinCases(it.m.Cases, 3))
			}
			preview := ""
			if it.m.Preview {
				preview = ` <span class="pill pill-preview">preview</span>`
			}
			fmt.Fprintf(&b, `<tr class="cov-row%s"><td>%s.%s%s</td><td>%s</td><td>%s</td><td>%s</td></tr>`+"\n",
				cls,
				html.EscapeString(strings.TrimPrefix(it.cls.ID, "Excel.")),
				html.EscapeString(it.m.Signature),
				preview,
				html.EscapeString(shortAPISet(it.m.APISet)),
				mog, corpus)
		}
		b.WriteString(`</tbody></table></section>` + "\n")
	}
	if pages > 0 {
		b.WriteString(`<script>
(function () {
  var bar = document.getElementById("cov-filter");
  if (!bar) return;
  bar.addEventListener("click", function (ev) {
    var btn = ev.target.closest("button[data-filter]");
    if (!btn) return;
    bar.querySelectorAll("button").forEach(function (b) { b.classList.toggle("on", b === btn); });
    var f = btn.getAttribute("data-filter");
    document.querySelectorAll("tr.cov-row").forEach(function (tr) {
      var v = tr.classList.contains("verified");
      var i = tr.classList.contains("implemented");
      var show = f === "all" || (f === "verified" && v) || (f === "implemented" && i) || (f === "both" && v && i);
      tr.classList.toggle("cov-hidden", !show);
    });
  });
})();
</script>
`)
	}
	return b.String()
}

func coverageFunctionsPages(fn CoverageClass) string {
	names := make([]string, 0, len(fn.Members))
	for _, m := range fn.Members {
		if m.Kind == "method" {
			names = append(names, m.Name)
		}
	}
	if len(names) == 0 {
		return ""
	}
	pages := (len(names) + coverageFnPerPage - 1) / coverageFnPerPage
	var b strings.Builder
	for p := 0; p < pages; p++ {
		start := p * coverageFnPerPage
		end := start + coverageFnPerPage
		if end > len(names) {
			end = len(names)
		}
		b.WriteString(`<section class="page">`)
		if p == 0 {
			b.WriteString(`<h2>Excel.Functions</h2>`)
			fmt.Fprintf(&b, `<p class="caption">%d worksheet-function wrappers on <code>workbook.functions</code> (ExcelApi 1.2). This corpus does not call them; scripts write formulas onto ranges instead. Listed so coverage of the Office.js surface stays honest.</p>`, len(names))
		} else {
			fmt.Fprintf(&b, `<h2>Excel.Functions (%d/%d)</h2>`, p+1, pages)
		}
		b.WriteString(`<div class="member-grid">`)
		for _, name := range names[start:end] {
			fmt.Fprintf(&b, `<div class="fn"><code>%s</code></div>`, html.EscapeString(name))
		}
		b.WriteString(`</div></section>` + "\n")
	}
	return b.String()
}

func coverageRestPages(classes []CoverageClass) string {
	if len(classes) == 0 {
		return ""
	}
	var b strings.Builder
	type entry struct {
		id      string
		methods int
	}
	items := make([]entry, 0, len(classes))
	for _, c := range classes {
		if c.MethodCount == 0 && c.PropertyCount == 0 {
			continue
		}
		items = append(items, entry{c.ID, c.MethodCount})
	}
	const perPage = 120
	pages := (len(items) + perPage - 1) / perPage
	for p := 0; p < pages; p++ {
		start := p * perPage
		end := start + perPage
		if end > len(items) {
			end = len(items)
		}
		b.WriteString(`<section class="page">`)
		if p == 0 {
			b.WriteString(`<h2>Not yet in Mog or the corpus</h2>`)
			b.WriteString(`<p class="caption">Object-model classes with no implemented host methods and no verification script calls yet.</p>`)
		} else {
			fmt.Fprintf(&b, `<h2>Not yet in Mog or the corpus (%d/%d)</h2>`, p+1, pages)
		}
		b.WriteString(`<div class="member-grid">`)
		for _, it := range items[start:end] {
			fmt.Fprintf(&b, `<div class="fn"><code>%s</code> · %d</div>`, html.EscapeString(it.id), it.methods)
		}
		b.WriteString(`</div></section>` + "\n")
	}
	return b.String()
}

func chunkStrings(rows []string, n int) [][]string {
	if n <= 0 || len(rows) == 0 {
		return nil
	}
	var out [][]string
	for i := 0; i < len(rows); i += n {
		end := i + n
		if end > len(rows) {
			end = len(rows)
		}
		out = append(out, rows[i:end])
	}
	return out
}

func shortAPISet(s string) string {
	s = strings.TrimSpace(s)
	s = strings.TrimPrefix(s, "ExcelApi ")
	s = strings.TrimPrefix(s, "ExcelApi")
	if s == "" {
		return "—"
	}
	return s
}

func joinCases(cases []string, max int) string {
	if len(cases) == 0 {
		return ""
	}
	if len(cases) <= max {
		return strings.Join(cases, ", ")
	}
	return strings.Join(cases[:max], ", ") + "…"
}

func strconvMin(n, capN int) string {
	if n > capN {
		return fmt.Sprintf("%d", capN) + "+"
	}
	return fmt.Sprintf("%d", n)
}

func writeCoverageCopy(src, outDir string) error {
	if strings.TrimSpace(src) == "" {
		return nil
	}
	data, err := os.ReadFile(src)
	if err != nil {
		return err
	}
	return os.WriteFile(filepath.Join(outDir, coverageJSONName), data, 0o644)
}
