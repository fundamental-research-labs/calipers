package main

import (
	"context"
	"flag"
	"fmt"
	"math"
	"os"
	"os/exec"
	"path/filepath"
	"strings"
	"sync/atomic"
	"time"

	"github.com/fundamental-research-labs/calipers/internal/cases"
	"github.com/fundamental-research-labs/calipers/internal/excel"
)

const (
	defaultBudgetMargin = 1.5
	durationFloorMs     = 250
	memoryFloorBytes    = 32 << 20
	measureSampleEvery  = 25 * time.Millisecond
	measureBudgetsUsage = `calipers measure-budgets [--engine excel|PATH] [--margin 1.5] [--force] [--cases-dir DIR] [--suite NAME]

  Run each selected case, record wall time and peak working set, and write
  config.json next to init.xlsx / script.js / golden.xlsx:

      { "maxPeakMemoryBytes": N, "maxDurationMs": N }

  Values are the measured quantity times --margin, plus a small floor, so
  later engine runs can flag regressions without matching Excel exactly.

  Missing config.json is valid (no budget). This command fills those files.
  --force overwrites configs that already exist.

  --engine excel    Excel COM host (Windows + Microsoft Excel). Default.
  --engine PATH     external binary (same argv as verify: save / run).

  Goldens are not written. Off Windows, --engine excel errors with the
  same message as excel-save once there is work to do. Empty work
  (nothing to measure) succeeds on any OS.

  Run this on Windows after excel-run-pass. Do not run it in Linux CI.
`
)

func measureBudgetsCmd(args []string) error {
	fs := flag.NewFlagSet("measure-budgets", flag.ContinueOnError)
	fs.SetOutput(os.Stderr)
	engineSpec := fs.String("engine", "excel", "engine binary path, or 'excel' (default)")
	margin := fs.Float64("margin", defaultBudgetMargin, "multiplier applied to measured peak memory and duration")
	force := fs.Bool("force", false, "overwrite existing config.json files")
	casesDir := fs.String("cases-dir", cases.DirName, "verification cases directory")
	suite := fs.String("suite", "", "suite directory to measure (default: all suites)")
	fs.Usage = func() {
		fmt.Fprint(os.Stdout, measureBudgetsUsage)
	}
	if err := fs.Parse(args); err != nil {
		if err == flag.ErrHelp {
			return nil
		}
		return err
	}
	if fs.NArg() > 1 {
		return fmt.Errorf("usage: calipers measure-budgets [--engine excel|PATH] [--margin 1.5] [--force] [--cases-dir DIR] [--suite NAME]")
	}
	if fs.NArg() == 1 {
		*casesDir = fs.Arg(0)
	}
	if *margin < 1 {
		return fmt.Errorf("--margin must be >= 1, got %g", *margin)
	}
	return measureBudgets(*casesDir, *suite, strings.TrimSpace(*engineSpec), *margin, *force)
}

func measureBudgets(casesDir, suite, engineSpec string, margin float64, force bool) error {
	corpus, err := cases.LoadCorpus(casesDir)
	if err != nil {
		return err
	}
	all, err := cases.FilterSuite(corpus.Cases, suite, corpus.Suites)
	if err != nil {
		return err
	}
	targets := all
	if !force {
		targets = cases.MissingBudget(all)
	}
	skipped := len(all) - len(targets)
	if len(targets) == 0 {
		printf("measure-budgets: 0 wrote, 0 failed, %d skipped\n", skipped)
		return nil
	}

	run, err := budgetRunner(engineSpec)
	if err != nil {
		return err
	}

	failed := 0
	wrote := 0
	for i, c := range targets {
		printf("[%d/%d] %s\n", i+1, len(targets), c.ID)
		sample, err := run(c)
		if err != nil {
			fmt.Fprintf(os.Stderr, "calipers: %s: %v\n", c.ID, err)
			_ = os.Stderr.Sync()
			failed++
			continue
		}
		b := budgetFromSample(sample, margin)
		if err := cases.WriteBudget(c.Dir, b); err != nil {
			fmt.Fprintf(os.Stderr, "calipers: %s: write %s: %v\n", c.ID, cases.ConfigFile, err)
			_ = os.Stderr.Sync()
			failed++
			continue
		}
		printf("  peak=%dB duration=%dms -> maxPeakMemoryBytes=%d maxDurationMs=%d\n",
			sample.peakBytes, sample.duration.Milliseconds(), b.MaxPeakMemoryBytes, b.MaxDurationMs)
		wrote++
	}
	printf("measure-budgets: %d wrote, %d failed, %d skipped\n", wrote, failed, skipped)
	if failed > 0 {
		return fmt.Errorf("%d case(s) failed", failed)
	}
	return nil
}

type memSample struct {
	peakBytes int64
	duration  time.Duration
}

func budgetFromSample(s memSample, margin float64) cases.Budget {
	durMs := s.duration.Milliseconds()
	if durMs <= 0 {
		durMs = 1
	}
	return cases.Budget{
		MaxPeakMemoryBytes: applyMargin(s.peakBytes, margin, memoryFloorBytes),
		MaxDurationMs:      applyMargin(durMs, margin, durationFloorMs),
	}
}

func applyMargin(v int64, margin float64, floor int64) int64 {
	if v <= 0 {
		return 0
	}
	out := int64(math.Ceil(float64(v) * margin))
	if floor > 0 && out < v+floor {
		out = v + floor
	}
	return out
}

type caseRunner func(c cases.Case) (memSample, error)

func budgetRunner(engineSpec string) (caseRunner, error) {
	switch engineSpec {
	case "", "excel":
		host := excel.NewHost()
		if !host.Available() {
			return nil, excel.ErrNotWindows
		}
		return func(c cases.Case) (memSample, error) {
			out, err := os.CreateTemp("", "calipers-measure-*.xlsx")
			if err != nil {
				return memSample{}, err
			}
			outPath := out.Name()
			_ = out.Close()
			defer os.Remove(outPath)
			return measureExcel(host, c, outPath)
		}, nil
	default:
		abs, err := filepath.Abs(engineSpec)
		if err != nil {
			return nil, err
		}
		return func(c cases.Case) (memSample, error) {
			out, err := os.CreateTemp("", "calipers-measure-*.xlsx")
			if err != nil {
				return memSample{}, err
			}
			outPath := out.Name()
			_ = out.Close()
			defer os.Remove(outPath)
			args := []string{"save", c.InitPath, outPath}
			if c.RunScript() {
				args = []string{"run", c.InitPath, c.ScriptPath, outPath}
			}
			return measureChild(abs, args)
		}, nil
	}
}

func measureExcel(host excel.Host, c cases.Case, outPath string) (memSample, error) {
	var peak atomic.Int64
	ctx, cancel := context.WithCancel(context.Background())
	done := make(chan struct{})
	go func() {
		defer close(done)
		tick := time.NewTicker(measureSampleEvery)
		defer tick.Stop()
		for {
			select {
			case <-ctx.Done():
				return
			case <-tick.C:
				for _, pid := range excelPIDs() {
					if n := processPeakBytes(pid); n > peak.Load() {
						peak.Store(n)
					}
				}
			}
		}
	}()
	start := time.Now()
	var err error
	if c.RunScript() {
		err = host.RunScript(c.InitPath, c.ScriptPath, outPath)
	} else {
		err = host.OpenSave(c.InitPath, outPath)
	}
	sample := memSample{peakBytes: peak.Load(), duration: time.Since(start)}
	cancel()
	<-done
	if err != nil {
		return sample, err
	}
	return sample, nil
}

func measureChild(bin string, args []string) (memSample, error) {
	cmd := exec.Command(bin, args...)
	start := time.Now()
	if err := cmd.Start(); err != nil {
		return memSample{}, err
	}
	var peak atomic.Int64
	ctx, cancel := context.WithCancel(context.Background())
	done := make(chan struct{})
	pid := cmd.Process.Pid
	go func() {
		defer close(done)
		tick := time.NewTicker(measureSampleEvery)
		defer tick.Stop()
		for {
			select {
			case <-ctx.Done():
				return
			case <-tick.C:
				if n := processPeakBytes(pid); n > peak.Load() {
					peak.Store(n)
				}
			}
		}
	}()
	err := cmd.Wait()
	sample := memSample{peakBytes: peak.Load(), duration: time.Since(start)}
	cancel()
	<-done
	if err != nil {
		return sample, err
	}
	return sample, nil
}
