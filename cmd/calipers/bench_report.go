package main

import (
	"flag"
	"fmt"
	"html"
	"math"
	"os"
	"path/filepath"
	"strings"
	"time"
)

const (
	benchReportHTMLName = "report.html"
	rowsPerLetterPage   = 28
	benchReportUsage    = `calipers bench-report --json IN.json --out DIR

  Read a JSON file from calipers bench and write report.html plus SVG
  charts into DIR. The HTML is self-contained (inline SVG, no ES modules)
  so a browser can open it via file:// .

  The page is laid out as stacked US Letter sheets (8.5×11in). Use the
  Export as PDF button (window.print); @page size is letter with one
  HTML sheet per PDF page.

  Layout: short setup (Excel COM, sideloaded Office.js add-in, sequential
  same-machine series, Mog save/run), two per-task plots (speed, memory),
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
<title>Excel vs Mog — verifier corpus</title>
<style>
@page { size: letter; margin: 0; }
html { font-family: "Segoe UI", "Helvetica Neue", Helvetica, Arial, sans-serif; color: #1a1a1a; }
body { margin: 0; background: #c5ccd4; font-size: 10pt; line-height: 1.4; }
.toolbar {
  position: sticky; top: 0; z-index: 2;
  display: flex; align-items: center; gap: 0.7em; flex-wrap: wrap;
  background: #1f2a33; color: #f3f5f7; padding: 0.55em 1em;
}
.toolbar p { margin: 0; font-size: 9.5pt; color: #c5ced6; }
.toolbar button {
  font: 600 10.5pt/1.2 "Segoe UI", Helvetica, Arial, sans-serif;
  background: #f4f7fa; color: #1f2a33; border: 0; border-radius: 4px;
  padding: 0.45em 0.9em; cursor: pointer;
}
.toolbar button:hover { background: #fff; }
.page {
  box-sizing: border-box;
  width: 8.5in; height: 11in;
  margin: 0.45in auto;
  padding: 0.55in 0.6in 0.5in;
  background: #fff;
  box-shadow: 0 1px 10px rgba(20, 28, 36, 0.28);
  overflow: hidden;
  page-break-after: always;
  break-after: page;
}
.page:last-of-type { page-break-after: auto; break-after: auto; }
h1 { font-size: 16pt; margin: 0 0 0.15em; letter-spacing: -0.01em; }
h2 { font-size: 11.5pt; margin: 0 0 0.35em; }
.meta { color: #5a6570; font-size: 8.5pt; margin: 0 0 0.55em; }
.setup { margin: 0 0 0.45em; }
.setup p { margin: 0 0 0.35em; }
.setup p:last-child { margin-bottom: 0; }
.note {
  background: #fff6e8; border: 1px solid #ead7b4; border-radius: 4px;
  padding: 0.35em 0.55em; margin: 0 0 0.45em; font-size: 9pt;
}
code {
  font-family: ui-monospace, "Cascadia Mono", "SFMono-Regular", Consolas, Menlo, monospace;
  font-size: 0.86em;
  background: #eef2f6;
  color: #1a3550;
  border-radius: 3px;
  padding: 0.07em 0.32em;
  white-space: nowrap;
}
.engines { margin: 0.15em 0 0.4em; padding: 0; list-style: none; font-size: 9pt; color: #33404a; }
.engines li { margin: 0.12em 0; }
.chart { margin: 0.15em 0 0.25em; }
.chart svg { width: 100%; height: auto; display: block; }
.caption { color: #5a6570; font-size: 8.5pt; margin: 0 0 0.35em; }
table { border-collapse: collapse; width: 100%; font-size: 8pt; table-layout: fixed; }
th, td { border-bottom: 1px solid #d8dee4; padding: 0.16em 0.22em; text-align: right; vertical-align: top; word-wrap: break-word; }
th:first-child, td:first-child, th:nth-child(2), td:nth-child(2) { text-align: left; }
th { background: #eef2f5; font-weight: 600; }
tr.error td { background: #fdecea; }
thead { display: table-header-group; }
tr { break-inside: avoid; page-break-inside: avoid; }
@media print {
  body { background: #fff; }
  .toolbar { display: none !important; }
  .page {
    margin: 0; box-shadow: none;
    width: 8.5in; height: 11in;
    page-break-after: always; break-after: page;
    break-inside: avoid; page-break-inside: avoid;
  }
  .page:last-of-type { page-break-after: auto; break-after: auto; }
}
</style>
</head>
<body>
<div class="toolbar">
<button type="button" id="export-pdf" onclick="window.print()">Export as PDF</button>
<p>US Letter pages below — in the dialog choose Save as PDF, paper US Letter, margins none or default.</p>
</div>
`)

	b.WriteString(`<section class="page">
<h1>Excel vs Mog</h1>
<p class="meta">Our Excel engine verifier · wall time and peak working set · one task at a time</p>
`)
	fmt.Fprintf(&b, `<p class="meta">%s · %s/%s · %d tasks · %d engine(s)</p>`,
		html.EscapeString(doc.GeneratedAt.UTC().Format(time.RFC3339)),
		html.EscapeString(doc.Host.GOOS), html.EscapeString(doc.Host.Arch),
		len(doc.Cases), len(doc.Engines))

	if !benchHasExcelCOM(doc) {
		b.WriteString(`<p class="note"><strong>Excel COM series was not collected.</strong> This is a Mog-only preview. Re-run on a Windows machine with desktop Excel to add the Excel series on the same plots.</p>`)
	}

	b.WriteString(`<div class="setup">
<p>Both series run on the <strong>same Windows machine</strong>, sequentially: the next process starts only after the previous one has exited, so concurrency cannot spoil monitoring.</p>
<p><strong>Excel</strong> — desktop Excel through COM (<code>Excel.Application</code> on an STA thread, alerts and window hidden, then open / save xlsx). Office.js tasks use a <strong>sideloaded Office.js add-in</strong>, not AppSource and not Office Scripts: the add-in <code>Excel.run</code>s the script, then the workbook is saved.</p>
<p><strong>Mog</strong> — <code>save in.xlsx out.xlsx</code> or <code>run in.xlsx script.js out.xlsx</code>.</p>
<p><strong>Wall time</strong> is that task’s elapsed time. <strong>Peak working set</strong> is sampled every 25&nbsp;ms (Windows <code>PeakWorkingSetSize</code> of <code>EXCEL.EXE</code> or the child; Unix <code>VmHWM</code>). Tasks differ a lot — there is no single average. Each mark below is one task.</p>
</div>
<ul class="engines">
`)
	for _, eng := range doc.Engines {
		fmt.Fprintf(&b, `<li><code>%s</code> · %s · <code>%s</code></li>`,
			html.EscapeString(eng.ID), html.EscapeString(eng.Kind), html.EscapeString(eng.Spec))
	}
	b.WriteString("</ul>\n")
	writeInlineChart(&b, svgs, "speed.svg")
	writeInlineChart(&b, svgs, "memory.svg")
	b.WriteString("</section>\n")
	b.WriteString(caseTablePages(doc))
	b.WriteString("</body>\n</html>\n")
	return b.String()
}

func writeInlineChart(b *strings.Builder, svgs map[string]string, name string) {
	markup, ok := svgs[name]
	if !ok {
		return
	}
	fmt.Fprintf(b, `<div class="chart">%s</div>`, markup)
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
		return `<section class="page"><h2>Per-task table</h2><p>No cases.</p></section>`
	}
	pages := (n + rowsPerLetterPage - 1) / rowsPerLetterPage
	for p := 0; p < pages; p++ {
		start := p * rowsPerLetterPage
		end := start + rowsPerLetterPage
		if end > n {
			end = n
		}
		b.WriteString(`<section class="page">`)
		if p == 0 {
			b.WriteString("<h2>Per-task table</h2>")
			b.WriteString(`<p class="caption">Each row is one task. Times are milliseconds; peaks are working set. The plots on page 1 are the overview — tasks are not averaged.</p>`)
		} else {
			fmt.Fprintf(&b, `<h2>Per-task table (%d/%d)</h2>`, p+1, pages)
		}
		b.WriteString("<table>")
		b.WriteString(headers())
		b.WriteString("<tbody>\n")
		for _, c := range doc.Cases[start:end] {
			b.WriteString(writeRow(c))
		}
		b.WriteString("</tbody></table></section>\n")
	}
	return b.String()
}

func buildBenchCharts(doc BenchFile) map[string]string {
	names := make([]string, len(doc.Engines))
	xs := make([][]float64, len(doc.Engines))
	speed := make([][]float64, len(doc.Engines))
	mem := make([][]float64, len(doc.Engines))
	for i, eng := range doc.Engines {
		names[i] = eng.ID
		for j, c := range doc.Cases {
			r := c.Results[eng.ID]
			if r.Error != "" || r.DurationMs <= 0 {
				continue
			}
			xs[i] = append(xs[i], float64(j+1))
			speed[i] = append(speed[i], float64(r.DurationMs))
			if r.PeakBytes > 0 {
				mem[i] = append(mem[i], float64(r.PeakBytes))
			} else {
				mem[i] = append(mem[i], math.NaN())
			}
		}
	}
	return map[string]string{
		"speed.svg": svgTaskMarks("Speed — one mark per task", "task index", "wall time", names, xs, speed, func(v float64) string {
			return formatMs(int64(math.Round(v)))
		}),
		"memory.svg": svgTaskMarks("Memory — one mark per task", "task index", "peak working set", names, xs, mem, func(v float64) string {
			return formatBytes(int64(math.Round(v)))
		}),
	}
}

var chartColors = []string{"#1f4e79", "#c45911", "#548235", "#7030a0"}

func svgTaskMarks(title, xlab, ylab string, series []string, xs, ys [][]float64, yfmt func(float64) string) string {
	const W, H, lpad, rpad, tpad, bpad = 720.0, 250.0, 72.0, 14.0, 28.0, 40.0
	innerW := W - lpad - rpad
	innerH := H - tpad - bpad
	maxX := 1.0
	minY, maxY := math.Inf(1), 0.0
	nPts := 0
	for s := range ys {
		for i, v := range ys[s] {
			if math.IsNaN(v) || v <= 0 {
				continue
			}
			nPts++
			if v < minY {
				minY = v
			}
			if v > maxY {
				maxY = v
			}
			if i < len(xs[s]) && xs[s][i] > maxX {
				maxX = xs[s][i]
			}
		}
	}
	if minY > maxY || maxY <= 0 {
		minY, maxY = 1, 10
	}
	if minY == maxY {
		minY *= 0.5
		maxY *= 1.5
		if minY <= 0 {
			minY = maxY / 10
		}
	}
	// log-y: tasks differ by orders of magnitude; a mean would be meaningless
	logY := maxY/minY >= 8
	if logY {
		minY = math.Pow(10, math.Floor(math.Log10(minY)))
		maxY = math.Pow(10, math.Ceil(math.Log10(maxY)))
		if maxY <= minY {
			maxY = minY * 10
		}
	}
	yAt := func(v float64) float64 {
		if logY {
			return tpad + innerH - innerH*(math.Log10(v)-math.Log10(minY))/(math.Log10(maxY)-math.Log10(minY))
		}
		return tpad + innerH - innerH*(v-minY)/(maxY-minY)
	}
	xAt := func(v float64) float64 {
		return lpad + innerW*(v-1)/math.Max(maxX-1, 1)
	}

	var b strings.Builder
	fmt.Fprintf(&b, `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 %.0f %.0f" role="img" aria-label="%s">`, W, H, html.EscapeString(title))
	b.WriteString(`<rect width="100%" height="100%" fill="#fff"/>`)
	fmt.Fprintf(&b, `<text x="%.1f" y="16" font-size="12" font-family="Segoe UI, Helvetica, Arial, sans-serif" fill="#1a1a1a">%s</text>`, lpad, html.EscapeString(title))
	fmt.Fprintf(&b, `<line x1="%.1f" y1="%.1f" x2="%.1f" y2="%.1f" stroke="#333" stroke-width="1"/>`, lpad, tpad, lpad, tpad+innerH)
	fmt.Fprintf(&b, `<line x1="%.1f" y1="%.1f" x2="%.1f" y2="%.1f" stroke="#333" stroke-width="1"/>`, lpad, tpad+innerH, lpad+innerW, tpad+innerH)

	ticks := yTicks(minY, maxY, logY)
	for _, v := range ticks {
		y := yAt(v)
		fmt.Fprintf(&b, `<line x1="%.1f" y1="%.1f" x2="%.1f" y2="%.1f" stroke="#e6eaee" stroke-width="1"/>`, lpad, y, lpad+innerW, y)
		label := compactNum(v)
		if yfmt != nil {
			label = yfmt(v)
		}
		fmt.Fprintf(&b, `<text x="%.1f" y="%.1f" font-size="8.5" text-anchor="end" fill="#555">%s</text>`, lpad-5, y+3, html.EscapeString(label))
	}

	arm := 2.4
	if nPts > 400 {
		arm = 1.8
	}
	for s := range series {
		if s >= len(ys) {
			continue
		}
		color := chartColors[s%len(chartColors)]
		b.WriteString(`<path fill="none" stroke="` + color + `" stroke-width="1.15" d="`)
		for i, v := range ys[s] {
			if math.IsNaN(v) || v <= 0 || i >= len(xs[s]) {
				continue
			}
			x, y := xAt(xs[s][i]), yAt(v)
			fmt.Fprintf(&b, "M%.1f %.1f h%.1f M%.1f %.1f v%.1f ", x-arm, y, 2*arm, x, y-arm, 2*arm)
		}
		b.WriteString(`"/>`)
	}

	fmt.Fprintf(&b, `<text x="%.1f" y="%.1f" font-size="9" fill="#555">1</text>`, lpad, tpad+innerH+12)
	fmt.Fprintf(&b, `<text x="%.1f" y="%.1f" font-size="9" text-anchor="end" fill="#555">%s</text>`, lpad+innerW, tpad+innerH+12, html.EscapeString(compactNum(maxX)))
	fmt.Fprintf(&b, `<text x="%.1f" y="%.1f" font-size="9.5" text-anchor="middle" fill="#333">%s</text>`, lpad+innerW/2, H-8, html.EscapeString(xlab))
	scale := "linear"
	if logY {
		scale = "log"
	}
	fmt.Fprintf(&b, `<text x="12" y="%.1f" font-size="8.5" fill="#555" transform="rotate(-90 12 %.1f)">%s (%s)</text>`, tpad+innerH/2, tpad+innerH/2, html.EscapeString(ylab), scale)
	for s, name := range series {
		x := lpad + float64(s)*88
		color := chartColors[s%len(chartColors)]
		fmt.Fprintf(&b, `<path d="M%.1f %.1f h8 M%.1f %.1f v8" fill="none" stroke="%s" stroke-width="1.4"/>`, x, H-14, x+4, H-18, color)
		fmt.Fprintf(&b, `<text x="%.1f" y="%.1f" font-size="9.5" fill="#333">%s</text>`, x+12, H-9, html.EscapeString(name))
	}
	b.WriteString("</svg>")
	return b.String()
}

func yTicks(minY, maxY float64, logY bool) []float64 {
	if logY {
		var out []float64
		for e := math.Floor(math.Log10(minY)); e <= math.Ceil(math.Log10(maxY))+1e-9; e++ {
			out = append(out, math.Pow(10, e))
		}
		if len(out) == 0 {
			return []float64{minY, maxY}
		}
		return out
	}
	return []float64{minY, minY + (maxY-minY)/4, minY + (maxY-minY)/2, minY + 3*(maxY-minY)/4, maxY}
}

func benchHasExcelCOM(doc BenchFile) bool {
	for _, e := range doc.Engines {
		if e.Kind == "excel-com" || e.ID == "excel" {
			return true
		}
	}
	return false
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
