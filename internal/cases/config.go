package cases

import (
	"bytes"
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"
	"strings"

	"github.com/fundamental-research-labs/calipers/internal/xlsxmodel"
)

type resolved struct {
	Budget  Budget
	Compare xlsxmodel.Options
}

func resolveConfig(root, caseDir string) (resolved, string, error) {
	rel, err := filepath.Rel(root, caseDir)
	if err != nil {
		return resolved{}, "", fmt.Errorf("%s: %w", ConfigFile, err)
	}
	dirs := []string{root}
	if rel != "." {
		cur := root
		for _, p := range strings.Split(rel, string(filepath.Separator)) {
			if p == "" || p == "." {
				continue
			}
			cur = filepath.Join(cur, p)
			dirs = append(dirs, cur)
		}
	}
	var acc resolved
	local := ""
	for _, dir := range dirs {
		p := filepath.Join(dir, ConfigFile)
		present, err := mergeConfigFile(&acc, p)
		if err != nil {
			return resolved{}, "", err
		}
		if present && dir == caseDir {
			local = p
		}
	}
	return acc, local, nil
}

func mergeConfigFile(acc *resolved, path string) (bool, error) {
	data, err := os.ReadFile(path)
	if err != nil {
		if os.IsNotExist(err) {
			return false, nil
		}
		return false, err
	}
	if len(bytes.TrimSpace(data)) == 0 {
		return false, fmt.Errorf("%s: empty", ConfigFile)
	}
	var raw map[string]json.RawMessage
	if err := json.Unmarshal(data, &raw); err != nil {
		return false, fmt.Errorf("%s: %w", ConfigFile, err)
	}
	if v, ok := raw["maxPeakMemoryBytes"]; ok {
		var n int64
		if err := json.Unmarshal(v, &n); err != nil {
			return false, fmt.Errorf("%s: maxPeakMemoryBytes: %w", path, err)
		}
		if n < 0 {
			return false, fmt.Errorf("%s: negative budget", ConfigFile)
		}
		acc.Budget.MaxPeakMemoryBytes = n
	}
	if v, ok := raw["maxDurationMs"]; ok {
		var n int64
		if err := json.Unmarshal(v, &n); err != nil {
			return false, fmt.Errorf("%s: maxDurationMs: %w", path, err)
		}
		if n < 0 {
			return false, fmt.Errorf("%s: negative budget", ConfigFile)
		}
		acc.Budget.MaxDurationMs = n
	}
	if v, ok := raw["compare"]; ok {
		var child xlsxmodel.Options
		if err := json.Unmarshal(v, &child); err != nil {
			return false, fmt.Errorf("%s: compare: %w", path, err)
		}
		if err := validateCompare(path, child); err != nil {
			return false, err
		}
		acc.Compare = mergeCompare(acc.Compare, child)
	}
	return true, nil
}

func mergeCompare(parent, child xlsxmodel.Options) xlsxmodel.Options {
	return xlsxmodel.Options{
		Ignore: unionIgnore(parent.Ignore, child.Ignore),
		Cells:  mergeCells(parent.Cells, child.Cells),
	}
}

func unionIgnore(parent, child []string) []string {
	seen := map[string]struct{}{}
	var out []string
	for _, s := range append(append([]string{}, parent...), child...) {
		s = strings.TrimSpace(s)
		if s == "" {
			continue
		}
		if _, ok := seen[s]; ok {
			continue
		}
		seen[s] = struct{}{}
		out = append(out, s)
	}
	return out
}

func mergeCells(parent, child map[string]xlsxmodel.CellBand) map[string]xlsxmodel.CellBand {
	if len(parent) == 0 && len(child) == 0 {
		return nil
	}
	out := map[string]xlsxmodel.CellBand{}
	for k, v := range parent {
		out[k] = v
	}
	for k, v := range child {
		out[k] = v
	}
	return out
}

func validateCompare(path string, o xlsxmodel.Options) error {
	for i, s := range o.Ignore {
		s = strings.TrimSpace(s)
		if s == "" {
			return fmt.Errorf("%s: compare.ignore[%d] is empty", path, i)
		}
		if !xlsxmodel.KnownIgnore(s) {
			return fmt.Errorf("%s: unknown compare.ignore %q", path, s)
		}
	}
	for cell, band := range o.Cells {
		if strings.TrimSpace(cell) == "" {
			return fmt.Errorf("%s: compare.cells entry missing name", path)
		}
		if band.Min == nil || band.Max == nil {
			return fmt.Errorf("%s: compare.cells %q requires min and max", path, cell)
		}
		if *band.Min > *band.Max {
			return fmt.Errorf("%s: compare.cells %q min > max", path, cell)
		}
	}
	return nil
}

func budgetPtr(b Budget) *Budget {
	if b.MaxPeakMemoryBytes == 0 && b.MaxDurationMs == 0 {
		return nil
	}
	cp := b
	return &cp
}

// WriteBudget merges budget keys into dir/config.json, preserving other keys.
func WriteBudget(dir string, b Budget) error {
	if b.MaxPeakMemoryBytes < 0 || b.MaxDurationMs < 0 {
		return fmt.Errorf("%s: negative budget", ConfigFile)
	}
	p := filepath.Join(dir, ConfigFile)
	raw := map[string]any{}
	data, err := os.ReadFile(p)
	if err != nil {
		if !os.IsNotExist(err) {
			return err
		}
	} else if len(bytes.TrimSpace(data)) > 0 {
		if err := json.Unmarshal(data, &raw); err != nil {
			return fmt.Errorf("%s: %w", ConfigFile, err)
		}
	}
	raw["maxPeakMemoryBytes"] = b.MaxPeakMemoryBytes
	raw["maxDurationMs"] = b.MaxDurationMs
	out, err := json.MarshalIndent(raw, "", "  ")
	if err != nil {
		return err
	}
	out = append(out, '\n')
	return os.WriteFile(p, out, 0o644)
}
