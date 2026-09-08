// Package cases discovers verification cases on disk.
//
// A case is a directory named tier_{a|b|c}_<feature>/ containing a required
// init.xlsx, an optional script.js, and a dedicated golden.xlsx destination
// (not mixed into a flat init dump). A missing or empty script means load+save
// only: skip script execution.
package cases

import (
	"bytes"
	"fmt"
	"os"
	"path/filepath"
	"regexp"
	"sort"
)

// DirName is the in-repo corpus path relative to the module root.
const DirName = "verification/cases"

const (
	InitFile   = "init.xlsx"
	ScriptFile = "script.js"
	GoldenFile = "golden.xlsx"
)

// Tier is the case pass class encoded in the directory name.
type Tier string

const (
	TierA Tier = "a"
	TierB Tier = "b"
	TierC Tier = "c"
)

// Case is one verification case: init workbook, optional Office.js, golden dest.
type Case struct {
	ID         string // directory name, e.g. "tier_a_simple"
	Tier       Tier
	Dir        string
	InitPath   string
	ScriptPath string // non-empty only when a script should run
	GoldenPath string // destination; may not exist yet
}

// RunScript reports whether this case has a non-empty Office.js file to execute.
// False means load+save only.
func (c Case) RunScript() bool {
	return c.ScriptPath != ""
}

var tierName = regexp.MustCompile(`^tier_([abc])_.+`)

// Load walks root for case directories. Non-case entries are ignored.
// A tier_* directory without init.xlsx is an error.
func Load(root string) ([]Case, error) {
	entries, err := os.ReadDir(root)
	if err != nil {
		return nil, fmt.Errorf("cases: %w", err)
	}
	out := make([]Case, 0, len(entries))
	for _, e := range entries {
		if !e.IsDir() {
			continue
		}
		name := e.Name()
		tier, ok := parseTier(name)
		if !ok {
			continue
		}
		c, err := loadOne(filepath.Join(root, name), name, tier)
		if err != nil {
			return nil, err
		}
		out = append(out, c)
	}
	sort.Slice(out, func(i, j int) bool { return out[i].ID < out[j].ID })
	return out, nil
}

// DefaultPass is the cases to run unless asked otherwise: tier_a and tier_b.
// tier_c hostiles are excluded.
func DefaultPass(all []Case) []Case {
	out := make([]Case, 0, len(all))
	for _, c := range all {
		if c.Tier == TierA || c.Tier == TierB {
			out = append(out, c)
		}
	}
	return out
}

func parseTier(name string) (Tier, bool) {
	m := tierName.FindStringSubmatch(name)
	if m == nil {
		return "", false
	}
	return Tier(m[1]), true
}

func loadOne(dir, id string, tier Tier) (Case, error) {
	initPath := filepath.Join(dir, InitFile)
	if _, err := os.Stat(initPath); err != nil {
		return Case{}, fmt.Errorf("case %s: %s: %w", id, InitFile, err)
	}
	scriptPath, err := scriptToRun(dir)
	if err != nil {
		return Case{}, fmt.Errorf("case %s: %w", id, err)
	}
	return Case{
		ID:         id,
		Tier:       tier,
		Dir:        dir,
		InitPath:   initPath,
		ScriptPath: scriptPath,
		GoldenPath: filepath.Join(dir, GoldenFile),
	}, nil
}

func scriptToRun(dir string) (string, error) {
	p := filepath.Join(dir, ScriptFile)
	data, err := os.ReadFile(p)
	if err != nil {
		if os.IsNotExist(err) {
			return "", nil
		}
		return "", err
	}
	if len(bytes.TrimSpace(data)) == 0 {
		return "", nil
	}
	return p, nil
}
