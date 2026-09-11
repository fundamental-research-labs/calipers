package main

import (
	"flag"
	"fmt"
	"html"
	"math"
	"os"
	"path/filepath"
	"sort"
	"strings"
	"time"
)

const (
	benchReportHTMLName = "report.html"
	rowsPerLetterPage   = 36
	benchReportUsage    = `calipers bench-report --json IN.json --out DIR

  Read a JSON file from calipers bench and write report.html plus SVG
  charts into DIR. The HTML is self-contained (inline SVG, no ES modules)
  so a browser can open it via file:// and Print to PDF.

  Pagination is US Letter: @page { size: letter } and explicit page
  breaks so a ~1000-row per-case table does not become one sheet.

  Layout: setup (Excel COM, sideloaded Office.js add-in, sequential
  same-machine series, Mog save/run, wall time vs peak working set),
  then summary and suite aggregates, then compact ratios/outliers,
  then a paginated per-case table.
`
)

func benchReportCmd(args []string) error {
	fs := flag.NewFlagSet("bench-report", flag.ContinueOnError)
	fs.SetOutput(os.Stderr)
	jsonPath := fs.String("json", "", "input JSON from calipers bench (required)")
	outDir := fs.String("out", "", "directory for report.html and SVG charts (required)")
	fs.Usage = func() {
		fmt.Fprint(os.Stdout, benchReportUsage)
	}
	if err := fs.Parse(args); err != nil {
		if err == flag.ErrHelp {
			return nil
		}
		return err
	}
	if fs.NArg() > 0 || strings.TrimSpace(*jsonPath) == "" || strings.TrimSpace(*outDir) == "" {
		return fmt.Errorf("usage: calipers bench-report --json IN.json --out DIR")
	}
	return renderBenchReport(*jsonPath, *outDir)
}

func renderBenchReport(jsonPath, outDir string) error {
	doc, err := loadBenchFile(jsonPath)
	if err != nil {
		return err
	}
	if err := os.MkdirAll(outDir, 0o755); err != nil {
		return err
	}
	svgs := buildBenchCharts(doc)
	for name, markup := range svgs {
		if err := os.WriteFile(filepath.Join(outDir, name), []byte(markup), 0o644); err != nil {
			return err
		}
	}
	page := renderBenchHTML(doc, svgs)
	return os.WriteFile(filepath.Join(outDir, benchReportHTMLName), []byte(page), 0o644)
}

func renderBenchHTML(doc BenchFile, svgs map[string]string) string {
	var b strings.Builder
	b.WriteString(`<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Excel vs Mog verifier-corpus benchmark</title>
<style>
@page { size: letter; margin: 0.6in; }
html { font-family: "Segoe UI", "Helvetica Neue", Helvetica, Arial, sans-serif; color: #1a1a1a; }
body { margin: 0 auto; max-width: 7.3in; padding: 0.4in 0.2in 0.8in; font-size: 10.5pt; line-height: 1.45; }
h1 { font-size: 18pt; margin: 0 0 0.3em; line-height: 1.2; }
h2 { font-size: 13pt; margin: 1.2em 0 0.4em; page-break-after: avoid; }
h3 { font-size: 11pt; margin: 0.9em 0 0.3em; page-break-after: avoid; }
p, li { orphans: 3; widows: 3; }
.lede { font-size: 11pt; color: #333; }
.meta { color: #555; font-size: 9.5pt; }
.setup, .note { background: #f4f6f8; border: 1px solid #d5dbe0; padding: 0.55em 0.7em; margin: 0.6em 0; }
.setup h3 { margin-top: 0.4em; }
.setup h3:first-child { margin-top: 0; }
code, .mono { font-family: "Cascadia Mono", "SFMono-Regular", Consolas, Menlo, monospace; font-size: 9.5pt; }
.cards { display: flex; flex-wrap: wrap; gap: 0.45em; margin: 0.6em 0 0.8em; }
.card { border: 1px solid #cfd6dc; padding: 0.45em 0.6em; min-width: 1.5in; flex: 1 1 1.6in; page-break-inside: avoid; }
.card .k { display: block; color: #555; font-size: 8.5pt; text-transform: uppercase; letter-spacing: 0.04em; }
.card .v { display: block; font-size: 14pt; font-weight: 600; }
.chart { margin: 0.5em 0 0.8em; page-break-inside: avoid; }
.chart svg { width: 100%; height: auto; display: block; }
table { border-collapse: collapse; width: 100%; font-size: 8.5pt; table-layout: fixed; }
th, td { border-bottom: 1px solid #d8dee4; padding: 0.22em 0.28em; text-align: right; vertical-align: top; word-wrap: break-word; }
th:first-child, td:first-child, th:nth-child(2), td:nth-child(2) { text-align: left; }
th { background: #eef2f5; font-weight: 600; }
tr.error td { background: #fdecea; }
.page-break { break-after: page; page-break-after: always; }
.sheet { page-break-after: always; break-after: page; }
.sheet:last-child { page-break-after: auto; break-after: auto; }
thead { display: table-header-group; }
tr { break-inside: avoid; page-break-inside: avoid; }
.caption { color: #555; font-size: 9pt; margin: 0.2em 0 0.5em; }
footer { margin-top: 1.2em; color: #666; font-size: 8.5pt; }
@media print {
  body { max-width: none; padding: 0; }
  a { color: inherit; text-decoration: none; }
}
</style>
</head>
<body>
`)
	b.WriteString(`<article class="sheet">
<h1>Excel vs Mog verifier-corpus benchmark</h1>
<p class="lede">Wall time and peak working set for every committed-golden calipers verify case, measured once on desktop Excel and once on Mog, sequentially on the same machine.</p>
`)
	fmt.Fprintf(&b, `<p class="meta">Collected %s · host %s/%s · %d cases · %d engine(s) · JSON schema v%d</p>`,
		html.EscapeString(doc.GeneratedAt.UTC().Format(time.RFC3339)),
		html.EscapeString(doc.Host.GOOS), html.EscapeString(doc.Host.Arch),
		len(doc.Cases), len(doc.Engines), doc.Version)

	b.WriteString(`<h2>Setup</h2>
<div class="setup">
<h3>Same machine, one task at a time</h3>
<p>Both series run on the <strong>same Windows machine</strong>. Excel can only be driven through COM on Windows; Mog is cross-platform, so the fair comparison is both series on that Windows box. Cases and engines never overlap: the next process starts only after the previous one has exited. Sequential execution keeps wall-clock time and peak-memory sampling from being spoiled by concurrency.</p>
<h3>Excel via COM (not AppSource, not Office Scripts)</h3>
<p>Desktop Microsoft Excel is the <code>excel</code> engine, the same host as <code>calipers excel-save</code>, <code>excel-run</code>, and <code>verify --engine excel</code>. Calipers creates an <code>Excel.Application</code> COM object on an STA thread, with <code>DisplayAlerts</code> and <code>Visible</code> off, then <code>Workbooks.Open</code> and <code>SaveAs</code> xlsx (format 51).</p>
<p>Office.js corpus scripts are <strong>not</strong> eval’d through COM and are <strong>not</strong> Office Scripts / Automate. They run inside Excel through a <strong>sideloaded Office.js add-in</strong> (Web Extension Framework, Developer sideload + TrustedCatalogs — not an AppSource listing). A local HTTP server serves the task pane; the add-in <code>Excel.run</code>s the case script; the workbook is saved as xlsx.</p>
<h3>Mog as a <code>save</code> / <code>run</code> binary</h3>
<p>Mog is a caller-supplied engine binary using the same argv contract as <code>calipers verify --engine PATH</code>:</p>
<pre class="mono">mog save &lt;in.xlsx&gt; &lt;out.xlsx&gt;
mog run  &lt;in.xlsx&gt; &lt;script.js&gt; &lt;out.xlsx&gt;</pre>
<p>Load+save cases call <code>save</code>. Cases with a non-empty <code>script.js</code> call <code>run</code>. This command does not change verify PASS/FAIL semantics and does not write <code>config.json</code> budgets (<code>measure-budgets</code> remains separate).</p>
<h3>What the numbers are</h3>
<p><strong>Wall time</strong> is elapsed time of one case on one engine (open <code>init.xlsx</code>, run Office.js when present, export xlsx). <strong>Peak working set</strong> is sampled every 25&nbsp;ms while that process runs: Windows <code>PeakWorkingSetSize</code> of <code>EXCEL.EXE</code> or the child, Unix <code>VmHWM</code>. A case error is recorded in the JSON and the series continues.</p>
</div>
`)
	b.WriteString("<h3>This run</h3>")
	if !benchHasExcelCOM(doc) {
		b.WriteString(`<p class="note"><strong>Excel COM series was not collected.</strong> This JSON is a Mog-only (or other binary) preview — useful off Windows to inspect the HTML before a colleague runs <code>--engine excel --engine &lt;mog&gt;</code> on a Windows box with desktop Excel. Dual-engine ratios and the Excel column appear once that series is included.</p>`)
	}
	b.WriteString("<ul>")
	for _, eng := range doc.Engines {
		fmt.Fprintf(&b, `<li><code>%s</code> · %s · <code>%s</code></li>`,
			html.EscapeString(eng.ID), html.EscapeString(eng.Kind), html.EscapeString(eng.Spec))
	}
	b.WriteString("</ul>")
	if !doc.StartedAt.IsZero() && !doc.EndedAt.IsZero() {
		fmt.Fprintf(&b, `<p class="meta">Series wall clock (first start to last end): %s.</p>`,
			html.EscapeString(doc.EndedAt.Sub(doc.StartedAt).Truncate(time.Millisecond).String()))
	}
	b.WriteString("</article>\n")

	sum := summarizeBench(doc)
	b.WriteString(`<article class="sheet">
<h2>Summary</h2>
<div class="cards">
`)
	fmt.Fprintf(&b, `<div class="card"><span class="k">Cases</span><span class="v">%d</span></div>`, len(doc.Cases))
	fmt.Fprintf(&b, `<div class="card"><span class="k">Engines</span><span class="v">%d</span></div>`, len(doc.Engines))
	fmt.Fprintf(&b, `<div class="card"><span class="k">Errors</span><span class="v">%d</span></div>`, sum.errors)
	for _, eng := range doc.Engines {
		st := sum.byEngine[eng.ID]
		fmt.Fprintf(&b, `<div class="card"><span class="k">%s median time</span><span class="v">%s</span></div>`,
			html.EscapeString(eng.ID), html.EscapeString(formatMs(st.medianMs)))
		fmt.Fprintf(&b, `<div class="card"><span class="k">%s p95 time</span><span class="v">%s</span></div>`,
			html.EscapeString(eng.ID), html.EscapeString(formatMs(st.p95Ms)))
		fmt.Fprintf(&b, `<div class="card"><span class="k">%s median peak</span><span class="v">%s</span></div>`,
			html.EscapeString(eng.ID), html.EscapeString(formatBytes(st.medianBytes)))
	}
	b.WriteString("</div>\n")
	if len(doc.Engines) >= 2 {
		a, bID := doc.Engines[0].ID, doc.Engines[1].ID
		fmt.Fprintf(&b, `<p class="caption">Ratios are %s / %s. Values below 1 mean %s used less time or memory.</p>`,
			html.EscapeString(bID), html.EscapeString(a), html.EscapeString(bID))
	}
	b.WriteString("<h3>Suite aggregates</h3>\n")
	b.WriteString(sum.suiteTable)
	writeInlineChart(&b, svgs, "suite-duration.svg")
	writeInlineChart(&b, svgs, "suite-memory.svg")
	b.WriteString("</article>\n")

	b.WriteString(`<article class="sheet">
<h2>Comparison</h2>
<p class="caption">At ~1000 cases, per-case bar charts are unreadable. The figures below are suite bars, ratio histograms, and a duration scatter — not one bar per test. Outliers are listed; every case is in the paginated table that follows.</p>
`)
	writeInlineChart(&b, svgs, "duration-hist.svg")
	writeInlineChart(&b, svgs, "memory-hist.svg")
	writeInlineChart(&b, svgs, "duration-ratio.svg")
	writeInlineChart(&b, svgs, "memory-ratio.svg")
	writeInlineChart(&b, svgs, "duration-scatter.svg")
	b.WriteString("<h3>Outliers</h3>\n")
	b.WriteString(sum.outlierTable)
	b.WriteString("</article>\n")

	b.WriteString(sum.casePages)
	b.WriteString(`<footer>Print this page to PDF from the browser (Letter, backgrounds on). file:// is supported: no ES modules, charts are inline SVG.</footer>
</body>
</html>
`)
	return b.String()
}

func writeInlineChart(b *strings.Builder, svgs map[string]string, name string) {
	markup, ok := svgs[name]
	if !ok {
		return
	}
	fmt.Fprintf(b, `<div class="chart">%s</div>`, markup)
}

type engineStat struct {
	medianMs    int64
	p95Ms       int64
	medianBytes int64
}

type benchSummary struct {
	errors       int
	byEngine     map[string]engineStat
	suiteTable   string
	outlierTable string
	casePages    string
}

func summarizeBench(doc BenchFile) benchSummary {
	sum := benchSummary{byEngine: map[string]engineStat{}}
	ms := map[string][]float64{}
	mem := map[string][]float64{}
	for _, c := range doc.Cases {
		for id, r := range c.Results {
			if r.Error != "" {
				sum.errors++
				continue
			}
			ms[id] = append(ms[id], float64(r.DurationMs))
			mem[id] = append(mem[id], float64(r.PeakBytes))
		}
	}
	for _, eng := range doc.Engines {
		sum.byEngine[eng.ID] = engineStat{
			medianMs:    int64(percentile(ms[eng.ID], 0.5)),
			p95Ms:       int64(percentile(ms[eng.ID], 0.95)),
			medianBytes: int64(percentile(mem[eng.ID], 0.5)),
		}
	}
	sum.suiteTable = suiteAggregateTable(doc)
	sum.outlierTable = outlierTable(doc)
	sum.casePages = caseTablePages(doc)
	return sum
}

func suiteAggregateTable(doc BenchFile) string {
	type row struct {
		name string
		n    int
		err  int
		ms   map[string][]float64
		mem  map[string][]float64
	}
	order := []string{}
	by := map[string]*row{}
	for _, c := range doc.Cases {
		name := c.Suite
		if name == "" {
			name = "(flat)"
		}
		r, ok := by[name]
		if !ok {
			r = &row{name: name, ms: map[string][]float64{}, mem: map[string][]float64{}}
			by[name] = r
			order = append(order, name)
		}
		r.n++
		for id, res := range c.Results {
			if res.Error != "" {
				r.err++
				continue
			}
			r.ms[id] = append(r.ms[id], float64(res.DurationMs))
			r.mem[id] = append(r.mem[id], float64(res.PeakBytes))
		}
	}
	sort.Strings(order)
	var b strings.Builder
	b.WriteString("<table><thead><tr><th>Suite</th><th>N</th>")
	for _, eng := range doc.Engines {
		fmt.Fprintf(&b, "<th>%s med</th><th>%s p95</th><th>%s mem</th>",
			html.EscapeString(eng.ID), html.EscapeString(eng.ID), html.EscapeString(eng.ID))
	}
	if len(doc.Engines) >= 2 {
		b.WriteString("<th>Time ×</th><th>Mem ×</th>")
	}
	b.WriteString("<th>Err</th></tr></thead><tbody>\n")
	for _, name := range order {
		r := by[name]
		fmt.Fprintf(&b, "<tr><td>%s</td><td>%d</td>", html.EscapeString(r.name), r.n)
		medMs := map[string]float64{}
		medMem := map[string]float64{}
		for _, eng := range doc.Engines {
			medMs[eng.ID] = percentile(r.ms[eng.ID], 0.5)
			p95 := percentile(r.ms[eng.ID], 0.95)
			medMem[eng.ID] = percentile(r.mem[eng.ID], 0.5)
			fmt.Fprintf(&b, "<td>%s</td><td>%s</td><td>%s</td>",
				html.EscapeString(formatMs(int64(medMs[eng.ID]))),
				html.EscapeString(formatMs(int64(p95))),
				html.EscapeString(formatBytes(int64(medMem[eng.ID]))))
		}
		if len(doc.Engines) >= 2 {
			a, bID := doc.Engines[0].ID, doc.Engines[1].ID
			fmt.Fprintf(&b, "<td>%s</td><td>%s</td>",
				html.EscapeString(formatRatio(medMs[bID], medMs[a])),
				html.EscapeString(formatRatio(medMem[bID], medMem[a])))
		}
		fmt.Fprintf(&b, "<td>%d</td></tr>\n", r.err)
	}
	b.WriteString("</tbody></table>\n")
	return b.String()
}

type outlier struct {
	id    string
	ratio float64
	kind  string
	a, b  string
}

func outlierTable(doc BenchFile) string {
	if len(doc.Engines) < 2 {
		return `<p class="note">Outliers need two engines.</p>`
	}
	a, bID := doc.Engines[0].ID, doc.Engines[1].ID
	var timeOut, memOut []outlier
	for _, c := range doc.Cases {
		ra, okA := c.Results[a]
		rb, okB := c.Results[bID]
		if !okA || !okB || ra.Error != "" || rb.Error != "" {
			continue
		}
		if ra.DurationMs > 0 {
			timeOut = append(timeOut, outlier{id: c.ID, ratio: float64(rb.DurationMs) / float64(ra.DurationMs), kind: "time", a: formatMs(ra.DurationMs), b: formatMs(rb.DurationMs)})
		}
		if ra.PeakBytes > 0 {
			memOut = append(memOut, outlier{id: c.ID, ratio: float64(rb.PeakBytes) / float64(ra.PeakBytes), kind: "mem", a: formatBytes(ra.PeakBytes), b: formatBytes(rb.PeakBytes)})
		}
	}
	rank := func(list []outlier) []outlier {
		sort.Slice(list, func(i, j int) bool {
			di := math.Abs(math.Log(math.Max(list[i].ratio, 1e-9)))
			dj := math.Abs(math.Log(math.Max(list[j].ratio, 1e-9)))
			return di > dj
		})
		if len(list) > 12 {
			list = list[:12]
		}
		return list
	}
	timeOut, memOut = rank(timeOut), rank(memOut)
	var b strings.Builder
	fmt.Fprintf(&b, `<p class="caption">Largest |log ratio| versus %s (up to 12 per metric). Time × and Mem × are %s / %s.</p>`,
		html.EscapeString(a), html.EscapeString(bID), html.EscapeString(a))
	b.WriteString("<table><thead><tr><th>Case</th><th>Metric</th>")
	fmt.Fprintf(&b, "<th>%s</th><th>%s</th><th>×</th></tr></thead><tbody>\n",
		html.EscapeString(a), html.EscapeString(bID))
	for _, o := range append(timeOut, memOut...) {
		fmt.Fprintf(&b, "<tr><td>%s</td><td>%s</td><td>%s</td><td>%s</td><td>%s</td></tr>\n",
			html.EscapeString(o.id), html.EscapeString(o.kind), html.EscapeString(o.a), html.EscapeString(o.b),
			html.EscapeString(formatRatio(o.ratio, 1)))
	}
	if len(timeOut)+len(memOut) == 0 {
		b.WriteString(`<tr><td colspan="5">No paired successful samples.</td></tr>`)
	}
	b.WriteString("</tbody></table>\n")
	return b.String()
}

func caseTablePages(doc BenchFile) string {
	var b strings.Builder
	headers := func() string {
		var h strings.Builder
		h.WriteString("<thead><tr><th>Case</th><th>Suite</th>")
		for _, eng := range doc.Engines {
			fmt.Fprintf(&h, "<th>%s ms</th><th>%s peak</th>", html.EscapeString(eng.ID), html.EscapeString(eng.ID))
		}
		if len(doc.Engines) >= 2 {
			h.WriteString("<th>Time ×</th><th>Mem ×</th>")
		}
		h.WriteString("<th>Notes</th></tr></thead>")
		return h.String()
	}
	writeRow := func(c BenchCase) string {
		var r strings.Builder
		cls := ""
		notes := []string{}
		if c.Script {
			notes = append(notes, "script")
		}
		for _, eng := range doc.Engines {
			if res, ok := c.Results[eng.ID]; ok && res.Error != "" {
				cls = ` class="error"`
				notes = append(notes, eng.ID+": "+res.Error)
			}
		}
		suite := c.Suite
		if suite == "" {
			suite = "—"
		}
		fmt.Fprintf(&r, "<tr%s><td>%s</td><td>%s</td>", cls, html.EscapeString(c.ID), html.EscapeString(suite))
		ms := map[string]int64{}
		mem := map[string]int64{}
		for _, eng := range doc.Engines {
			res := c.Results[eng.ID]
			if res.Error != "" {
				r.WriteString("<td>—</td><td>—</td>")
				continue
			}
			ms[eng.ID] = res.DurationMs
			mem[eng.ID] = res.PeakBytes
			fmt.Fprintf(&r, "<td>%d</td><td>%s</td>", res.DurationMs, html.EscapeString(formatBytes(res.PeakBytes)))
		}
		if len(doc.Engines) >= 2 {
			a, bID := doc.Engines[0].ID, doc.Engines[1].ID
			fmt.Fprintf(&r, "<td>%s</td><td>%s</td>",
				html.EscapeString(formatRatio(float64(ms[bID]), float64(ms[a]))),
				html.EscapeString(formatRatio(float64(mem[bID]), float64(mem[a]))))
		}
		fmt.Fprintf(&r, "<td>%s</td></tr>\n", html.EscapeString(strings.Join(notes, "; ")))
		return r.String()
	}

	n := len(doc.Cases)
	if n == 0 {
		return `<article class="sheet"><h2>Per-case results</h2><p>No cases.</p></article>`
	}
	pages := (n + rowsPerLetterPage - 1) / rowsPerLetterPage
	for p := 0; p < pages; p++ {
		start := p * rowsPerLetterPage
		end := start + rowsPerLetterPage
		if end > n {
			end = n
		}
		b.WriteString(`<article class="sheet">`)
		if p == 0 {
			b.WriteString("<h2>Per-case results</h2>")
			b.WriteString(`<p class="caption">Compact table, US Letter pagination (page-break after each block). Times are milliseconds; peaks are working set.</p>`)
		} else {
			fmt.Fprintf(&b, `<h3>Per-case results (continued %d/%d)</h3>`, p+1, pages)
		}
		b.WriteString("<table>")
		b.WriteString(headers())
		b.WriteString("<tbody>\n")
		for _, c := range doc.Cases[start:end] {
			b.WriteString(writeRow(c))
		}
		b.WriteString("</tbody></table></article>\n")
	}
	return b.String()
}

func buildBenchCharts(doc BenchFile) map[string]string {
	out := map[string]string{}
	suites, dur, mem := suiteSeries(doc)
	engNames := make([]string, len(doc.Engines))
	for i, e := range doc.Engines {
		engNames[i] = e.ID
	}
	out["suite-duration.svg"] = svgGroupedBars("Median wall time by suite", "milliseconds", suites, engNames, dur)
	out["suite-memory.svg"] = svgGroupedBars("Median peak working set by suite", "bytes", suites, engNames, mem)
	if len(doc.Engines) < 2 {
		id := doc.Engines[0].ID
		var durs, mems []float64
		for _, c := range doc.Cases {
			r := c.Results[id]
			if r.Error != "" {
				continue
			}
			durs = append(durs, float64(r.DurationMs))
			mems = append(mems, float64(r.PeakBytes))
		}
		out["duration-hist.svg"] = svgHistogram("Wall time distribution — "+id, "milliseconds", durs)
		out["memory-hist.svg"] = svgHistogram("Peak working set distribution — "+id, "bytes", mems)
		return out
	}
	a, b := doc.Engines[0].ID, doc.Engines[1].ID
	var tRatio, mRatio, xs, ys []float64
	for _, c := range doc.Cases {
		ra, okA := c.Results[a]
		rb, okB := c.Results[b]
		if !okA || !okB || ra.Error != "" || rb.Error != "" {
			continue
		}
		if ra.DurationMs > 0 {
			tRatio = append(tRatio, float64(rb.DurationMs)/float64(ra.DurationMs))
			xs = append(xs, float64(ra.DurationMs))
			ys = append(ys, float64(rb.DurationMs))
		}
		if ra.PeakBytes > 0 {
			mRatio = append(mRatio, float64(rb.PeakBytes)/float64(ra.PeakBytes))
		}
	}
	out["duration-ratio.svg"] = svgHistogram("Duration ratio "+b+" / "+a, "ratio (1 = same wall time)", tRatio)
	out["memory-ratio.svg"] = svgHistogram("Memory ratio "+b+" / "+a, "ratio (1 = same peak working set)", mRatio)
	out["duration-scatter.svg"] = svgScatter("Duration scatter", a+" ms", b+" ms", xs, ys)
	return out
}

func suiteSeries(doc BenchFile) (cats []string, dur, mem [][]float64) {
	type acc struct {
		ms, bytes map[string][]float64
	}
	by := map[string]*acc{}
	for _, c := range doc.Cases {
		name := c.Suite
		if name == "" {
			name = "(flat)"
		}
		a, ok := by[name]
		if !ok {
			a = &acc{ms: map[string][]float64{}, bytes: map[string][]float64{}}
			by[name] = a
			cats = append(cats, name)
		}
		for id, r := range c.Results {
			if r.Error != "" {
				continue
			}
			a.ms[id] = append(a.ms[id], float64(r.DurationMs))
			a.bytes[id] = append(a.bytes[id], float64(r.PeakBytes))
		}
	}
	sort.Strings(cats)
	dur = make([][]float64, len(doc.Engines))
	mem = make([][]float64, len(doc.Engines))
	for i, eng := range doc.Engines {
		dur[i] = make([]float64, len(cats))
		mem[i] = make([]float64, len(cats))
		for j, name := range cats {
			dur[i][j] = percentile(by[name].ms[eng.ID], 0.5)
			mem[i][j] = percentile(by[name].bytes[eng.ID], 0.5)
		}
	}
	return cats, dur, mem
}

var chartColors = []string{"#1f4e79", "#c45911", "#548235", "#7030a0"}

func svgGroupedBars(title, ylab string, cats, series []string, values [][]float64) string {
	const W, H, lpad, rpad, tpad, bpad = 720.0, 280.0, 64.0, 16.0, 36.0, 64.0
	innerW := W - lpad - rpad
	innerH := H - tpad - bpad
	maxV := 1.0
	for _, row := range values {
		for _, v := range row {
			if v > maxV {
				maxV = v
			}
		}
	}
	nCat := len(cats)
	if nCat == 0 {
		nCat = 1
	}
	nSer := len(series)
	if nSer == 0 {
		nSer = 1
	}
	slot := innerW / float64(nCat)
	barW := slot * 0.7 / float64(nSer)
	var b strings.Builder
	fmt.Fprintf(&b, `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 %.0f %.0f" role="img" aria-label="%s">`, W, H, html.EscapeString(title))
	b.WriteString(`<rect width="100%" height="100%" fill="#fff"/>`)
	fmt.Fprintf(&b, `<text x="%.1f" y="18" font-size="13" font-family="Segoe UI, Helvetica, Arial, sans-serif" fill="#1a1a1a">%s</text>`, lpad, html.EscapeString(title))
	fmt.Fprintf(&b, `<line x1="%.1f" y1="%.1f" x2="%.1f" y2="%.1f" stroke="#333" stroke-width="1"/>`, lpad, tpad, lpad, tpad+innerH)
	fmt.Fprintf(&b, `<line x1="%.1f" y1="%.1f" x2="%.1f" y2="%.1f" stroke="#333" stroke-width="1"/>`, lpad, tpad+innerH, lpad+innerW, tpad+innerH)
	for i := 0; i <= 4; i++ {
		y := tpad + innerH*float64(4-i)/4
		v := maxV * float64(i) / 4
		fmt.Fprintf(&b, `<line x1="%.1f" y1="%.1f" x2="%.1f" y2="%.1f" stroke="#e2e6ea" stroke-width="1"/>`, lpad, y, lpad+innerW, y)
		fmt.Fprintf(&b, `<text x="%.1f" y="%.1f" font-size="9" text-anchor="end" fill="#555">%s</text>`, lpad-6, y+3, html.EscapeString(compactNum(v)))
	}
	for i, cat := range cats {
		cx := lpad + slot*(float64(i)+0.5)
		fmt.Fprintf(&b, `<text x="%.1f" y="%.1f" font-size="9" text-anchor="middle" fill="#333">%s</text>`, cx, tpad+innerH+14, html.EscapeString(cat))
		for s := range series {
			if s >= len(values) || i >= len(values[s]) {
				continue
			}
			v := values[s][i]
			h := innerH * (v / maxV)
			x := lpad + slot*float64(i) + slot*0.15 + barW*float64(s)
			y := tpad + innerH - h
			color := chartColors[s%len(chartColors)]
			fmt.Fprintf(&b, `<rect x="%.1f" y="%.1f" width="%.1f" height="%.1f" fill="%s"/>`, x, y, barW-1, h, color)
		}
	}
	for s, name := range series {
		x := lpad + float64(s)*90
		color := chartColors[s%len(chartColors)]
		fmt.Fprintf(&b, `<rect x="%.1f" y="%.1f" width="10" height="10" fill="%s"/>`, x, H-18, color)
		fmt.Fprintf(&b, `<text x="%.1f" y="%.1f" font-size="10" fill="#333">%s</text>`, x+14, H-9, html.EscapeString(name))
	}
	fmt.Fprintf(&b, `<text x="12" y="%.1f" font-size="9" fill="#555" transform="rotate(-90 12 %.1f)">%s</text>`, tpad+innerH/2, tpad+innerH/2, html.EscapeString(ylab))
	b.WriteString("</svg>")
	return b.String()
}

func svgHistogram(title, xlab string, samples []float64) string {
	const W, H, lpad, rpad, tpad, bpad = 720.0, 260.0, 48.0, 16.0, 36.0, 48.0
	innerW := W - lpad - rpad
	innerH := H - tpad - bpad
	nBucket := 16
	hi := 2.0
	if p := percentile(samples, 0.95); p > hi {
		hi = p * 1.1
	}
	if hi < 1 {
		hi = 1
	}
	counts := make([]int, nBucket)
	for _, v := range samples {
		if math.IsNaN(v) || math.IsInf(v, 0) || v < 0 {
			continue
		}
		idx := int(v / hi * float64(nBucket))
		if idx >= nBucket {
			idx = nBucket - 1
		}
		counts[idx]++
	}
	maxC := 1
	for _, c := range counts {
		if c > maxC {
			maxC = c
		}
	}
	var b strings.Builder
	fmt.Fprintf(&b, `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 %.0f %.0f" role="img" aria-label="%s">`, W, H, html.EscapeString(title))
	b.WriteString(`<rect width="100%" height="100%" fill="#fff"/>`)
	fmt.Fprintf(&b, `<text x="%.1f" y="18" font-size="13" font-family="Segoe UI, Helvetica, Arial, sans-serif" fill="#1a1a1a">%s</text>`, lpad, html.EscapeString(title))
	fmt.Fprintf(&b, `<line x1="%.1f" y1="%.1f" x2="%.1f" y2="%.1f" stroke="#333" stroke-width="1"/>`, lpad, tpad, lpad, tpad+innerH)
	fmt.Fprintf(&b, `<line x1="%.1f" y1="%.1f" x2="%.1f" y2="%.1f" stroke="#333" stroke-width="1"/>`, lpad, tpad+innerH, lpad+innerW, tpad+innerH)
	bw := innerW / float64(nBucket)
	for i, c := range counts {
		h := innerH * float64(c) / float64(maxC)
		x := lpad + bw*float64(i)
		y := tpad + innerH - h
		fmt.Fprintf(&b, `<rect x="%.1f" y="%.1f" width="%.1f" height="%.1f" fill="#1f4e79"/>`, x+1, y, bw-2, h)
	}
	// unity line when 1 is in range
	if hi > 0 {
		x1 := lpad + innerW*(1/hi)
		if x1 >= lpad && x1 <= lpad+innerW {
			fmt.Fprintf(&b, `<line x1="%.1f" y1="%.1f" x2="%.1f" y2="%.1f" stroke="#c45911" stroke-dasharray="4 3" stroke-width="1"/>`, x1, tpad, x1, tpad+innerH)
		}
	}
	fmt.Fprintf(&b, `<text x="%.1f" y="%.1f" font-size="9" fill="#555">0</text>`, lpad, tpad+innerH+14)
	fmt.Fprintf(&b, `<text x="%.1f" y="%.1f" font-size="9" text-anchor="end" fill="#555">%s</text>`, lpad+innerW, tpad+innerH+14, html.EscapeString(compactNum(hi)))
	fmt.Fprintf(&b, `<text x="%.1f" y="%.1f" font-size="10" text-anchor="middle" fill="#333">%s</text>`, lpad+innerW/2, H-10, html.EscapeString(xlab))
	fmt.Fprintf(&b, `<text x="%.1f" y="%.1f" font-size="9" fill="#555">n=%d</text>`, lpad, tpad+innerH+28, len(samples))
	b.WriteString("</svg>")
	return b.String()
}

func svgScatter(title, xlab, ylab string, xs, ys []float64) string {
	const W, H, lpad, rpad, tpad, bpad = 720.0, 280.0, 56.0, 16.0, 36.0, 48.0
	innerW := W - lpad - rpad
	innerH := H - tpad - bpad
	maxX, maxY := 1.0, 1.0
	for i := range xs {
		if xs[i] > maxX {
			maxX = xs[i]
		}
		if i < len(ys) && ys[i] > maxY {
			maxY = ys[i]
		}
	}
	// cap axes at p99 so a few monsters do not squash the cloud
	if p := percentile(xs, 0.99); p > 0 {
		maxX = math.Max(p*1.15, 1)
	}
	if p := percentile(ys, 0.99); p > 0 {
		maxY = math.Max(p*1.15, 1)
	}
	var b strings.Builder
	fmt.Fprintf(&b, `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 %.0f %.0f" role="img" aria-label="%s">`, W, H, html.EscapeString(title))
	b.WriteString(`<rect width="100%" height="100%" fill="#fff"/>`)
	fmt.Fprintf(&b, `<text x="%.1f" y="18" font-size="13" font-family="Segoe UI, Helvetica, Arial, sans-serif" fill="#1a1a1a">%s</text>`, lpad, html.EscapeString(title))
	fmt.Fprintf(&b, `<line x1="%.1f" y1="%.1f" x2="%.1f" y2="%.1f" stroke="#333" stroke-width="1"/>`, lpad, tpad, lpad, tpad+innerH)
	fmt.Fprintf(&b, `<line x1="%.1f" y1="%.1f" x2="%.1f" y2="%.1f" stroke="#333" stroke-width="1"/>`, lpad, tpad+innerH, lpad+innerW, tpad+innerH)
	diag := math.Min(maxX, maxY)
	x2 := lpad + innerW*(diag/maxX)
	y2 := tpad + innerH - innerH*(diag/maxY)
	fmt.Fprintf(&b, `<line x1="%.1f" y1="%.1f" x2="%.1f" y2="%.1f" stroke="#c45911" stroke-dasharray="4 3" stroke-width="1"/>`, lpad, tpad+innerH, x2, y2)
	for i := range xs {
		if i >= len(ys) {
			break
		}
		px := lpad + innerW*(xs[i]/maxX)
		py := tpad + innerH - innerH*(ys[i]/maxY)
		if px < lpad || px > lpad+innerW || py < tpad || py > tpad+innerH {
			continue
		}
		fmt.Fprintf(&b, `<circle cx="%.1f" cy="%.1f" r="2.2" fill="#1f4e79" fill-opacity="0.4"/>`, px, py)
	}
	fmt.Fprintf(&b, `<text x="%.1f" y="%.1f" font-size="10" text-anchor="middle" fill="#333">%s</text>`, lpad+innerW/2, H-10, html.EscapeString(xlab))
	fmt.Fprintf(&b, `<text x="12" y="%.1f" font-size="9" fill="#555" transform="rotate(-90 12 %.1f)">%s</text>`, tpad+innerH/2, tpad+innerH/2, html.EscapeString(ylab))
	fmt.Fprintf(&b, `<text x="%.1f" y="%.1f" font-size="9" fill="#555">0</text>`, lpad, tpad+innerH+14)
	fmt.Fprintf(&b, `<text x="%.1f" y="%.1f" font-size="9" text-anchor="end" fill="#555">%s</text>`, lpad+innerW, tpad+innerH+14, html.EscapeString(compactNum(maxX)))
	b.WriteString("</svg>")
	return b.String()
}

func benchHasExcelCOM(doc BenchFile) bool {
	for _, e := range doc.Engines {
		if e.Kind == "excel-com" || e.ID == "excel" {
			return true
		}
	}
	return false
}

func percentile(values []float64, p float64) float64 {
	if len(values) == 0 {
		return 0
	}
	cp := append([]float64(nil), values...)
	sort.Float64s(cp)
	if p <= 0 {
		return cp[0]
	}
	if p >= 1 {
		return cp[len(cp)-1]
	}
	idx := p * float64(len(cp)-1)
	lo := int(math.Floor(idx))
	hi := int(math.Ceil(idx))
	if lo == hi {
		return cp[lo]
	}
	w := idx - float64(lo)
	return cp[lo]*(1-w) + cp[hi]*w
}

func formatMs(ms int64) string {
	if ms < 1 {
		return "0 ms"
	}
	if ms < 10000 {
		return fmt.Sprintf("%d ms", ms)
	}
	return fmt.Sprintf("%.2f s", float64(ms)/1000)
}

func formatBytes(n int64) string {
	if n <= 0 {
		return "0 B"
	}
	mb := float64(n) / (1024 * 1024)
	if mb < 10 {
		return fmt.Sprintf("%.2f MB", mb)
	}
	return fmt.Sprintf("%.1f MB", mb)
}

func formatRatio(num, den float64) string {
	if den == 0 || num == 0 {
		return "—"
	}
	r := num / den
	if r >= 10 {
		return fmt.Sprintf("%.1f×", r)
	}
	return fmt.Sprintf("%.2f×", r)
}

func compactNum(v float64) string {
	if v >= 1e9 {
		return fmt.Sprintf("%.1fG", v/1e9)
	}
	if v >= 1e6 {
		return fmt.Sprintf("%.1fM", v/1e6)
	}
	if v >= 1e3 {
		return fmt.Sprintf("%.1fk", v/1e3)
	}
	if v >= 10 {
		return fmt.Sprintf("%.0f", v)
	}
	return fmt.Sprintf("%.1f", v)
}
