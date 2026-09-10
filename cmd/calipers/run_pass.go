package main

import (
	"fmt"
	"os"

	"github.com/fundamental-research-labs/calipers/internal/cases"
	"github.com/fundamental-research-labs/calipers/internal/excel"
	"github.com/fundamental-research-labs/calipers/internal/golden"
)

// excelRunPass generates goldens for scripted cases that still lack
// golden.xlsx. Analogous to excel-save-pass, but uses excel-run.
// Goldens must be produced on Windows + Excel; this command is not
// executed in Linux CI.
func excelRunPass(root string) error {
	all, err := cases.Load(root)
	if err != nil {
		return err
	}
	targets := cases.PendingScriptedGoldens(all)
	skipped := len(all) - len(targets)
	for _, c := range all {
		if !c.RunScript() {
			continue
		}
		st, err := os.Stat(c.GoldenPath)
		if err == nil && st.Size() > 0 {
			printf("skip %s (golden already present)\n", c.ID)
		}
	}

	if len(targets) == 0 {
		printf("excel-run-pass: 0 ok, 0 failed, %d skipped\n", skipped)
		return nil
	}

	host := excel.NewHost()
	if !host.Available() {
		return excel.ErrNotWindows
	}

	failed := map[string]error{}
	for i, c := range targets {
		printf("[%d/%d] %s\n", i+1, len(targets), c.ID)
		if err := generateGolden(host, c.InitPath, c.ScriptPath, c.GoldenPath); err != nil {
			fmt.Fprintf(os.Stderr, "calipers: %s: %v\n", c.ID, err)
			_ = os.Stderr.Sync()
			failed[c.ID] = err
		}
	}

	if err := checkUniformScriptedGoldens(targets, failed); err != nil {
		return err
	}

	printf("excel-run-pass: %d ok, %d failed, %d skipped\n",
		len(targets)-len(failed), len(failed), skipped)
	if len(failed) > 0 {
		return fmt.Errorf("%d case(s) failed", len(failed))
	}
	return nil
}

func checkUniformScriptedGoldens(targets []cases.Case, failed map[string]error) error {
	var first golden.Meta
	var firstID string
	n := 0
	for _, c := range targets {
		if _, skip := failed[c.ID]; skip {
			continue
		}
		m, err := golden.Read(c.GoldenPath)
		if err != nil {
			return fmt.Errorf("%s: %w", c.ID, err)
		}
		if m.Host != excel.HostID {
			return fmt.Errorf("%s: refuse golden host %q (want %s)", c.ID, m.Host, excel.HostID)
		}
		if m.Script != cases.ScriptFile {
			return fmt.Errorf("%s: scripted golden must record %s, got %q", c.ID, cases.ScriptFile, m.Script)
		}
		if n == 0 {
			first = m
			firstID = c.ID
			n++
			continue
		}
		if m.Host != first.Host || m.ExcelVersion != first.ExcelVersion || m.ExcelBuild != first.ExcelBuild {
			return fmt.Errorf("mixed excel goldens: %s is host=%s version=%s build=%s; %s is host=%s version=%s build=%s",
				c.ID, m.Host, m.ExcelVersion, m.ExcelBuild,
				firstID, first.Host, first.ExcelVersion, first.ExcelBuild)
		}
		n++
	}
	return nil
}
