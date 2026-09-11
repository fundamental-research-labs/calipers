package main

import (
	"encoding/json"
	"flag"
	"fmt"
	"os"
	"path/filepath"
	"runtime"
	"strings"
	"time"

	"github.com/fundamental-research-labs/calipers/internal/cases"
)

const benchUsage = `calipers bench --engine excel --engine PATH --json OUT.json [--report DIR] [--cases-dir DIR] [--suite NAME] [--case ID]...

  Run the same committed-golden corpus as calipers verify, once per
  --engine, one case at a time. Record wall time and peak working set
  for every case×engine in one JSON file. A single case error is stored
  in the JSON and does not skip the rest of the series.

  --engine excel    Excel COM host (Windows + Microsoft Excel).
  --engine PATH     external binary (same argv as verify: save / run).
                    Repeat --engine or comma-separate to run more than
                    one series (typical Windows: excel then Mog).
                    A single --engine PATH is the Mog-only path: works
                    off Windows so you can inspect JSON/HTML before a
                    colleague runs the Excel COM series.

  --report DIR      also write report.html + SVG charts (bench-report).

  Sequential: engines and cases never overlap. The next process starts
  only after the previous one has exited, so concurrency cannot spoil
  monitoring.

  Off Windows, --engine excel errors with the same message as excel-save
  once there is work to do. Empty work succeeds on any OS.

  This command does not write config.json (see measure-budgets) and
  does not compare goldens (see verify). Render JSON with bench-report.
`

// BenchFile is the JSON document written by `calipers bench`.
type BenchFile struct {
	Version     int           `json:"version"`
	GeneratedAt time.Time     `json:"generatedAt"`
	StartedAt   time.Time     `json:"startedAt"`
	EndedAt     time.Time     `json:"endedAt"`
	Host        BenchHost     `json:"host"`
	Engines     []BenchEngine `json:"engines"`
	Cases       []BenchCase   `json:"cases"`
}

// BenchHost is the machine that collected the series.
type BenchHost struct {
	OS   string `json:"os"`
	Arch string `json:"arch"`
	GOOS string `json:"goos"`
}

// BenchEngine is one measured host (Excel COM or a save/run binary).
type BenchEngine struct {
	ID   string `json:"id"`
	Kind string `json:"kind"` // excel-com | binary
	Spec string `json:"spec"`
}

// BenchCase is one verifier-corpus case and its per-engine samples.
type BenchCase struct {
	ID      string                 `json:"id"`
	Suite   string                 `json:"suite"`
	Script  bool                   `json:"script"`
	Results map[string]BenchResult `json:"results"`
}

// BenchResult is one case×engine sample.
type BenchResult struct {
	DurationMs int64     `json:"durationMs"`
	PeakBytes  int64     `json:"peakBytes"`
	StartedAt  time.Time `json:"startedAt"`
	EndedAt    time.Time `json:"endedAt"`
	Error      string    `json:"error,omitempty"`
}

func benchCmd(args []string) error {
	fs := flag.NewFlagSet("bench", flag.ContinueOnError)
	fs.SetOutput(os.Stderr)
	jsonPath := fs.String("json", "", "output JSON path (required)")
	reportDir := fs.String("report", "", "optional directory: also write report.html + SVG")
	casesDir := fs.String("cases-dir", cases.DirName, "verification cases directory")
	suite := fs.String("suite", "", "suite directory to run (default: all suites)")
	var engineSpecs []string
	fs.Func("engine", "engine binary path, or 'excel' (repeatable or comma-separated)", func(s string) error {
		for _, p := range strings.Split(s, ",") {
			p = strings.TrimSpace(p)
			if p != "" {
				engineSpecs = append(engineSpecs, p)
			}
		}
		return nil
	})
	var caseIDs []string
	fs.Func("case", "case id to run (suite/name; repeatable or comma-separated)", func(s string) error {
		for _, id := range strings.Split(s, ",") {
			id = strings.TrimSpace(id)
			if id != "" {
				caseIDs = append(caseIDs, id)
			}
		}
		return nil
	})
	fs.Usage = func() {
		fmt.Fprint(os.Stdout, benchUsage)
	}
	if err := fs.Parse(args); err != nil {
		if err == flag.ErrHelp {
			return nil
		}
		return err
	}
	if fs.NArg() > 0 {
		return fmt.Errorf("usage: calipers bench --engine excel --engine PATH --json OUT.json [--report DIR] [--cases-dir DIR] [--suite NAME] [--case ID]...")
	}
	if len(engineSpecs) == 0 {
		return fmt.Errorf("bench requires --engine <path|excel> (Mog-only: one PATH; Windows: excel then PATH)")
	}
	if strings.TrimSpace(*jsonPath) == "" {
		return fmt.Errorf("bench requires --json OUT.json")
	}
	err := runBench(*casesDir, *suite, caseIDs, engineSpecs, *jsonPath)
	if strings.TrimSpace(*reportDir) != "" {
		if _, statErr := os.Stat(*jsonPath); statErr == nil {
			if rerr := renderBenchReport(*jsonPath, *reportDir); rerr != nil {
				if err != nil {
					return fmt.Errorf("%v; report: %w", err, rerr)
				}
				return rerr
			}
			printf("bench-report: wrote %s\n", filepath.Join(*reportDir, benchReportHTMLName))
		}
	}
	return err
}

func runBench(casesDir, suite string, caseIDs, engineSpecs []string, jsonPath string) error {
	corpus, err := cases.LoadCorpus(casesDir)
	if err != nil {
		return err
	}
	all, err := cases.FilterSuite(corpus.Cases, suite, corpus.Suites)
	if err != nil {
		return err
	}
	var selected []cases.Case
	if suite == "" && len(caseIDs) == 0 {
		selected = cases.GoldenComparePass(all)
	} else {
		selected, err = cases.Select(all, caseIDs)
		if err != nil {
			return err
		}
		if len(caseIDs) == 0 {
			selected = cases.GoldenComparePass(selected)
		}
	}
	engines := nameEngines(engineSpecs)

	if len(selected) == 0 {
		doc := newBenchFile(engines)
		doc.EndedAt = time.Now().UTC()
		if err := writeBenchJSON(jsonPath, doc); err != nil {
			return err
		}
		printf("bench: 0 cases, %d engine(s); wrote %s\n", len(engines), jsonPath)
		return nil
	}

	runners := make([]caseRunner, len(engines))
	for i, eng := range engines {
		run, err := budgetRunner(eng.Spec)
		if err != nil {
			return err
		}
		runners[i] = run
	}

	doc := newBenchFile(engines)
	doc.Cases = make([]BenchCase, len(selected))
	for i, c := range selected {
		doc.Cases[i] = BenchCase{
			ID:      c.ID,
			Suite:   c.Suite,
			Script:  c.RunScript(),
			Results: make(map[string]BenchResult, len(engines)),
		}
	}

	nErr := 0
	total := len(engines) * len(selected)
	step := 0
	for i, eng := range engines {
		printf("bench series %s (%s)\n", eng.ID, eng.Kind)
		for j, c := range selected {
			step++
			printf("[%d/%d] %s %s\n", step, total, eng.ID, c.ID)
			res := measureOne(runners[i], c)
			doc.Cases[j].Results[eng.ID] = res
			if res.Error != "" {
				fmt.Fprintf(os.Stderr, "calipers: %s %s: %s\n", eng.ID, c.ID, res.Error)
				_ = os.Stderr.Sync()
				nErr++
				printf("  ERROR %s\n", res.Error)
			} else {
				printf("  duration=%dms peak=%dB\n", res.DurationMs, res.PeakBytes)
			}
		}
	}
	doc.EndedAt = time.Now().UTC()
	if err := writeBenchJSON(jsonPath, doc); err != nil {
		return err
	}
	printf("bench: %d cases × %d engine(s), %d error(s); wrote %s\n",
		len(selected), len(engines), nErr, jsonPath)
	if nErr > 0 {
		return fmt.Errorf("%d case×engine error(s); wrote %s", nErr, jsonPath)
	}
	return nil
}

func measureOne(run caseRunner, c cases.Case) BenchResult {
	start := time.Now().UTC()
	sample, err := run(c)
	end := time.Now().UTC()
	res := BenchResult{
		DurationMs: sample.duration.Milliseconds(),
		PeakBytes:  sample.peakBytes,
		StartedAt:  start,
		EndedAt:    end,
	}
	if res.DurationMs <= 0 {
		res.DurationMs = end.Sub(start).Milliseconds()
		if res.DurationMs <= 0 {
			res.DurationMs = 1
		}
	}
	if err != nil {
		res.Error = err.Error()
	}
	return res
}

func newBenchFile(engines []BenchEngine) BenchFile {
	now := time.Now().UTC()
	return BenchFile{
		Version:     1,
		GeneratedAt: now,
		StartedAt:   now,
		Host: BenchHost{
			OS:   runtime.GOOS,
			Arch: runtime.GOARCH,
			GOOS: runtime.GOOS,
		},
		Engines: engines,
		Cases:   []BenchCase{},
	}
}

func nameEngines(specs []string) []BenchEngine {
	used := make(map[string]int, len(specs))
	out := make([]BenchEngine, 0, len(specs))
	for _, spec := range specs {
		kind := "binary"
		id := filepath.Base(spec)
		if spec == "excel" {
			kind = "excel-com"
			id = "excel"
		}
		n := used[id]
		used[id]++
		if n > 0 {
			id = fmt.Sprintf("%s-%d", id, n+1)
		}
		out = append(out, BenchEngine{ID: id, Kind: kind, Spec: spec})
	}
	return out
}

func writeBenchJSON(path string, doc BenchFile) error {
	if dir := filepath.Dir(path); dir != "." && dir != "" {
		if err := os.MkdirAll(dir, 0o755); err != nil {
			return err
		}
	}
	f, err := os.Create(path)
	if err != nil {
		return err
	}
	defer f.Close()
	enc := json.NewEncoder(f)
	enc.SetIndent("", "  ")
	return enc.Encode(doc)
}

func loadBenchFile(path string) (BenchFile, error) {
	data, err := os.ReadFile(path)
	if err != nil {
		return BenchFile{}, err
	}
	var doc BenchFile
	if err := json.Unmarshal(data, &doc); err != nil {
		return BenchFile{}, fmt.Errorf("bench JSON: %w", err)
	}
	if len(doc.Engines) == 0 {
		return BenchFile{}, fmt.Errorf("bench JSON: missing engines")
	}
	return doc, nil
}
