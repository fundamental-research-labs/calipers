package main

import (
	"encoding/json"
	"errors"
	"os"
	"os/exec"
	"path/filepath"
	"runtime"
	"sort"
	"strings"
	"testing"
	"time"

	"github.com/fundamental-research-labs/calipers/internal/excel"
)

const engineStubSrc = `package main

import (
	"io"
	"os"
	"path/filepath"
	"strings"
	"time"
)

func main() {
	buf := make([]byte, 2<<20)
	for i := range buf {
		buf[i] = byte(i)
	}
	time.Sleep(50 * time.Millisecond)
	args := os.Args[1:]
	if len(args) < 3 {
		os.Exit(2)
	}
	var in, out string
	switch args[0] {
	case "save":
		in, out = args[1], args[2]
	case "run":
		if len(args) < 4 {
			os.Exit(2)
		}
		in, out = args[1], args[3]
	default:
		os.Exit(2)
	}
	slash := filepath.ToSlash(in)
	if strings.Contains(slash, "/boom/") {
		os.Exit(1)
	}
	f, err := os.Open(in)
	if err != nil {
		os.Exit(1)
	}
	defer f.Close()
	if err := os.MkdirAll(filepath.Dir(out), 0o755); err != nil {
		os.Exit(1)
	}
	g, err := os.Create(out)
	if err != nil {
		os.Exit(1)
	}
	defer g.Close()
	_, _ = io.Copy(g, f)
	_ = buf[0]
}
`

func TestRunBenchHelp(t *testing.T) {
	out, err := captureStdout(t, func() error {
		return run([]string{"bench", "-h"})
	})
	if err != nil {
		t.Fatal(err)
	}
	low := strings.ToLower(out)
	for _, want := range []string{"--engine", "excel", "json", "sequential", "com", "mog-only", "--report", "save", "run"} {
		if !strings.Contains(low, want) {
			t.Fatalf("bench help missing %q:\n%s", want, out)
		}
	}
}

func TestRunBenchRequiresEngineAndJSON(t *testing.T) {
	err := run([]string{"bench"})
	if err == nil || !strings.Contains(err.Error(), "--engine") {
		t.Fatalf("error = %v, want --engine", err)
	}
	err = run([]string{"bench", "--engine", "dummy"})
	if err == nil || !strings.Contains(err.Error(), "--json") {
		t.Fatalf("error = %v, want --json", err)
	}
}

func TestRunBenchExcelOffWindows(t *testing.T) {
	if runtime.GOOS == "windows" || os.Getenv("GOOS_FORCE") == "windows" {
		t.Skip("windows host implements Excel COM")
	}
	root := t.TempDir()
	writeNestedVerifyCase(t, root, "roundtrip", "simple", "hello", nil)
	err := run([]string{"bench", "--engine", "excel", "--cases-dir", root, "--json", filepath.Join(t.TempDir(), "x.json")})
	if err == nil {
		t.Fatal("expected error off Windows")
	}
	if !errors.Is(err, excel.ErrNotWindows) && !strings.Contains(err.Error(), "Windows") {
		t.Fatalf("bench excel error = %v", err)
	}
}

func TestRunBenchExcelEmptyDirOffWindows(t *testing.T) {
	if runtime.GOOS == "windows" || os.Getenv("GOOS_FORCE") == "windows" {
		t.Skip("windows host implements Excel COM")
	}
	jsonPath := filepath.Join(t.TempDir(), "empty.json")
	if err := run([]string{"bench", "--engine", "excel", "--cases-dir", t.TempDir(), "--json", jsonPath}); err != nil {
		t.Fatalf("empty corpus should not require Excel: %v", err)
	}
}

func TestRunBenchStubEnginesSequentialAndContinueAfterError(t *testing.T) {
	root := t.TempDir()
	js := "await Excel.run(async (context) => { await context.sync(); });\n"
	writeNestedVerifyCase(t, root, "roundtrip", "alpha", "hello", nil)
	writeNestedVerifyCase(t, root, "roundtrip", "boom", "hello", nil)
	writeNestedVerifyCase(t, root, "roundtrip", "zeta", "hello", &js)
	a, b := buildEngineStubs(t)
	jsonPath := filepath.Join(t.TempDir(), "bench.json")
	_, err := captureStdout(t, func() error {
		return run([]string{"bench", "--engine", a, "--engine", b, "--cases-dir", root, "--json", jsonPath})
	})
	if err == nil {
		t.Fatal("expected error from boom case")
	}
	if !strings.Contains(err.Error(), "error") {
		t.Fatalf("error = %v", err)
	}
	doc := readBenchJSON(t, jsonPath)
	if len(doc.Engines) != 2 {
		t.Fatalf("engines = %+v, want 2", doc.Engines)
	}
	if len(doc.Cases) != 3 {
		t.Fatalf("cases = %d, want 3 (later cases must remain after boom)", len(doc.Cases))
	}
	ids := map[string]bool{}
	for _, c := range doc.Cases {
		ids[c.ID] = true
		if len(c.Results) != 2 {
			t.Fatalf("%s results = %d, want both engines", c.ID, len(c.Results))
		}
		for eng, r := range c.Results {
			if r.DurationMs <= 0 {
				t.Fatalf("%s/%s durationMs = %d", c.ID, eng, r.DurationMs)
			}
			if r.PeakBytes <= 0 {
				t.Fatalf("%s/%s peakBytes = %d (sampler must record working set)", c.ID, eng, r.PeakBytes)
			}
			if r.StartedAt.IsZero() || r.EndedAt.IsZero() || !r.EndedAt.After(r.StartedAt) && !r.EndedAt.Equal(r.StartedAt) {
				t.Fatalf("%s/%s interval %v → %v", c.ID, eng, r.StartedAt, r.EndedAt)
			}
		}
	}
	if !ids["roundtrip/alpha"] || !ids["roundtrip/boom"] || !ids["roundtrip/zeta"] {
		t.Fatalf("cases = %v", ids)
	}
	var boomErr bool
	for _, r := range doc.Cases[indexCase(doc, "roundtrip/boom")].Results {
		if r.Error != "" {
			boomErr = true
		}
	}
	if !boomErr {
		t.Fatal("boom must record an error and still leave zeta in the JSON")
	}
	if !doc.Cases[indexCase(doc, "roundtrip/zeta")].Script {
		t.Fatal("zeta is scripted; JSON must record script=true")
	}
	assertNonOverlapping(t, doc)
}

func TestRunBenchMogOnlyWithReport(t *testing.T) {
	root := t.TempDir()
	writeNestedVerifyCase(t, root, "roundtrip", "simple", "hello", nil)
	writeNestedVerifyCase(t, root, "default", "other", "hello", nil)
	stub, _ := buildEngineStubs(t)
	out := t.TempDir()
	jsonPath := filepath.Join(out, "bench.json")
	reportDir := filepath.Join(out, "report")
	_, err := captureStdout(t, func() error {
		return run([]string{"bench", "--engine", stub, "--cases-dir", root, "--json", jsonPath, "--report", reportDir})
	})
	if err != nil {
		t.Fatal(err)
	}
	doc := readBenchJSON(t, jsonPath)
	if len(doc.Engines) != 1 {
		t.Fatalf("mog-only engines = %+v", doc.Engines)
	}
	if doc.Engines[0].Kind != "binary" {
		t.Fatalf("kind = %s, want binary", doc.Engines[0].Kind)
	}
	html := readFile(t, filepath.Join(reportDir, benchReportHTMLName))
	assertBenchHTMLCommon(t, html)
	if !strings.Contains(html, "Excel COM series was not collected") {
		t.Fatalf("mog-only HTML must say Excel was not collected:\n%s", html[:800])
	}
	if _, err := os.Stat(filepath.Join(reportDir, "speed.svg")); err != nil {
		t.Fatal(err)
	}
	if _, err := os.Stat(filepath.Join(reportDir, "memory.svg")); err != nil {
		t.Fatal(err)
	}
}

func buildEngineStubs(t *testing.T) (a, b string) {
	t.Helper()
	dir := t.TempDir()
	src := filepath.Join(dir, "main.go")
	if err := os.WriteFile(src, []byte(engineStubSrc), 0o644); err != nil {
		t.Fatal(err)
	}
	a = filepath.Join(dir, "stub-a")
	b = filepath.Join(dir, "stub-b")
	if runtime.GOOS == "windows" {
		a += ".exe"
		b += ".exe"
	}
	cmd := exec.Command("go", "build", "-o", a, src)
	cmd.Dir = dir
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("go build stub: %v\n%s", err, out)
	}
	data, err := os.ReadFile(a)
	if err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(b, data, 0o755); err != nil {
		t.Fatal(err)
	}
	return a, b
}

func readBenchJSON(t *testing.T, path string) BenchFile {
	t.Helper()
	data, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	var doc BenchFile
	if err := json.Unmarshal(data, &doc); err != nil {
		t.Fatalf("json: %v\n%s", err, data)
	}
	return doc
}

func indexCase(doc BenchFile, id string) int {
	for i, c := range doc.Cases {
		if c.ID == id {
			return i
		}
	}
	return 0
}

func assertNonOverlapping(t *testing.T, doc BenchFile) {
	t.Helper()
	type iv struct {
		start, end time.Time
		name       string
	}
	var all []iv
	for _, c := range doc.Cases {
		for eng, r := range c.Results {
			all = append(all, iv{r.StartedAt, r.EndedAt, c.ID + " " + eng})
		}
	}
	sort.Slice(all, func(i, j int) bool { return all[i].start.Before(all[j].start) })
	for i := 1; i < len(all); i++ {
		if all[i].start.Before(all[i-1].end) {
			t.Fatalf("overlapping intervals: %s %v–%v vs %s %v–%v",
				all[i-1].name, all[i-1].start, all[i-1].end,
				all[i].name, all[i].start, all[i].end)
		}
	}
}
