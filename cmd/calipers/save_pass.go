package main

import (
	"fmt"
	"os"

	"github.com/fundamental-research-labs/calipers/internal/cases"
	"github.com/fundamental-research-labs/calipers/internal/excel"
	"github.com/fundamental-research-labs/calipers/internal/golden"
)

// excelSavePass generates goldens for OpenSavePass via generateGolden.
// Scripted cases and tier_c are skipped (no Office.js goldens in this pass).
// Goldens must be excel-win and share one Excel version/build.
func excelSavePass(root string) error {
	all, err := cases.Load(root)
	if err != nil {
		return err
	}
	targets := cases.OpenSavePass(all)
	skipped := skippedOpenSave(all)
	for _, c := range skipped {
		printf("skip %s (Office.js; not an open+save golden)\n", c.ID)
	}

	if len(targets) == 0 {
		printf("excel-save-pass: 0 ok, 0 failed, %d skipped\n", len(skipped))
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

	if err := checkUniformOpenSaveGoldens(targets, failed); err != nil {
		return err
	}

	printf("excel-save-pass: %d ok, %d failed, %d skipped\n",
		len(targets)-len(failed), len(failed), len(skipped))
	if len(failed) > 0 {
		return fmt.Errorf("%d case(s) failed", len(failed))
	}
	return nil
}

func skippedOpenSave(all []cases.Case) []cases.Case {
	var skipped []cases.Case
	for _, c := range cases.DefaultPass(all) {
		if c.RunScript() {
			skipped = append(skipped, c)
		}
	}
	return skipped
}

func checkUniformOpenSaveGoldens(targets []cases.Case, failed map[string]error) error {
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
		if m.Script != "" {
			return fmt.Errorf("%s: load+save golden must omit script, got %q", c.ID, m.Script)
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
