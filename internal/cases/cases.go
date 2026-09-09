// Package cases discovers verification cases on disk.
//
// Cases live in suite directories under the corpus root:
//
//	verification/cases/<suite>/<case>/
//
// Suite names are the directory names on disk; the loader does not hardcode
// them. A case directory contains a required init.xlsx, an optional
// script.js, and a dedicated golden.xlsx destination (not mixed into a flat
// init dump). A missing or empty script means load+save only: skip script
// execution. The optional tier_{a|b|c}_ prefix classifies a case; other
// directory names are unprefixed cases when they contain init.xlsx.
//
// Load also accepts a flat directory of cases (one suite, used by tests
// and --cases-dir pointing at a single suite).
package cases

import (
	"bytes"
	"fmt"
	"os"
	"path/filepath"
	"regexp"
	"sort"
	"strings"
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
	ID         string // suite/name, e.g. "roundtrip/tier_a_simple"; name only in a flat layout
	Name       string // case directory name, e.g. "tier_a_simple"
	Suite      string // suite directory, e.g. "roundtrip"; empty in a flat layout
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

func makeID(suite, name string) string {
	if suite == "" {
		return name
	}
	return suite + "/" + name
}

var tierName = regexp.MustCompile(`^tier_([abc])_.+`)

// Load walks root for case directories across all suites. Non-case entries
// are ignored. A tier_* directory without init.xlsx is an error. An
// unprefixed directory is a case only when it contains init.xlsx.
//
// If root contains any case directories, it is treated as a single flat
// suite (Suite left empty). Otherwise every subdirectory is a suite whose
// own case children are loaded.
func Load(root string) ([]Case, error) {
	corpus, err := LoadCorpus(root)
	if err != nil {
		return nil, err
	}
	return corpus.Cases, nil
}

// Corpus is the cases and suite names discovered under a cases root.
type Corpus struct {
	Cases  []Case
	Suites []string // nested suite directory names; nil for a flat layout
}

// LoadCorpus is Load plus the suite names found under root, including empty
// suite directories.
func LoadCorpus(root string) (Corpus, error) {
	entries, err := os.ReadDir(root)
	if err != nil {
		return Corpus{}, fmt.Errorf("cases: %w", err)
	}

	var suiteDirs []string
	var hasCases bool
	for _, e := range entries {
		if !e.IsDir() {
			continue
		}
		name := e.Name()
		if _, ok := caseTier(root, name); ok {
			hasCases = true
			continue
		}
		suiteDirs = append(suiteDirs, name)
	}

	if hasCases {
		cs, err := loadSuite(root, "")
		if err != nil {
			return Corpus{}, err
		}
		return Corpus{Cases: cs}, nil
	}

	sort.Strings(suiteDirs)
	var out []Case
	for _, suite := range suiteDirs {
		cs, err := loadSuite(filepath.Join(root, suite), suite)
		if err != nil {
			return Corpus{}, err
		}
		out = append(out, cs...)
	}
	sort.Slice(out, func(i, j int) bool { return out[i].ID < out[j].ID })
	return Corpus{Cases: out, Suites: suiteDirs}, nil
}

func loadSuite(dir, suite string) ([]Case, error) {
	entries, err := os.ReadDir(dir)
	if err != nil {
		return nil, fmt.Errorf("cases: %w", err)
	}
	out := make([]Case, 0, len(entries))
	for _, e := range entries {
		if !e.IsDir() {
			continue
		}
		name := e.Name()
		tier, ok := caseTier(dir, name)
		if !ok {
			continue
		}
		c, err := loadOne(filepath.Join(dir, name), name, suite, tier)
		if err != nil {
			return nil, err
		}
		out = append(out, c)
	}
	sort.Slice(out, func(i, j int) bool { return out[i].Name < out[j].Name })
	return out, nil
}

// DefaultPass is the cases to run unless asked otherwise: every case except
// tier_c hostiles. Unprefixed cases (no tier_{a|b|c}_ name) are included.
// The default golden-comparison walk uses GoldenComparePass instead.
func DefaultPass(all []Case) []Case {
	out := make([]Case, 0, len(all))
	for _, c := range all {
		if c.Tier == TierC {
			continue
		}
		out = append(out, c)
	}
	return out
}

// GoldenComparePass is the default verify walk: DefaultPass minus cases
// with no committed golden.xlsx. Package comparison cannot score a case
// that has no oracle, regardless of suite name.
func GoldenComparePass(all []Case) []Case {
	pass := DefaultPass(all)
	out := make([]Case, 0, len(pass))
	for _, c := range pass {
		st, err := os.Stat(c.GoldenPath)
		if err != nil || st.Size() == 0 {
			continue
		}
		out = append(out, c)
	}
	return out
}

// OpenSavePass is the v1 excel-save golden set: DefaultPass minus cases that
// have Office.js to run. Scripted goldens are produced with excel-run, not
// by load+save. tier_c stays excluded.
func OpenSavePass(all []Case) []Case {
	pass := DefaultPass(all)
	out := make([]Case, 0, len(pass))
	for _, c := range pass {
		if c.RunScript() {
			continue
		}
		out = append(out, c)
	}
	return out
}

// FilterSuite restricts to one suite directory. Empty suite is a no-op.
// Unknown suite names error; a known empty suite returns no cases.
func FilterSuite(all []Case, suite string, known []string) ([]Case, error) {
	if suite == "" {
		return all, nil
	}
	var out []Case
	for _, c := range all {
		if c.Suite == suite {
			out = append(out, c)
		}
	}
	if len(out) > 0 {
		return out, nil
	}
	for _, s := range known {
		if s == suite {
			return nil, nil
		}
	}
	return nil, fmt.Errorf("unknown suite %q", suite)
}

// Select returns DefaultPass when ids is empty, otherwise the named cases
// in the given order (tier_c included only if asked). Unknown ids error.
// Ids are suite/name (the case id). A bare directory name is accepted when
// it uniquely identifies a case in the pool (e.g. after --suite).
func Select(all []Case, ids []string) ([]Case, error) {
	if len(ids) == 0 {
		return DefaultPass(all), nil
	}
	byID := make(map[string]Case, len(all))
	byName := make(map[string][]Case, len(all))
	for _, c := range all {
		byID[c.ID] = c
		byName[c.Name] = append(byName[c.Name], c)
	}
	out := make([]Case, 0, len(ids))
	for _, id := range ids {
		id = strings.TrimSpace(id)
		if c, ok := byID[id]; ok {
			out = append(out, c)
			continue
		}
		matches := byName[id]
		switch len(matches) {
		case 0:
			return nil, fmt.Errorf("unknown case %q", id)
		case 1:
			out = append(out, matches[0])
		default:
			return nil, fmt.Errorf("ambiguous case %q (present in multiple suites)", id)
		}
	}
	return out, nil
}

func parseTier(name string) (Tier, bool) {
	m := tierName.FindStringSubmatch(name)
	if m == nil {
		return "", false
	}
	return Tier(m[1]), true
}

// caseTier reports whether name under parent is a case directory and, if so,
// the tier encoded in the name. tier_* names are always cases (missing
// init.xlsx is an error in loadOne). Other names are cases only when they
// contain init.xlsx; they have no tier.
func caseTier(parent, name string) (Tier, bool) {
	if tier, ok := parseTier(name); ok {
		return tier, true
	}
	if _, err := os.Stat(filepath.Join(parent, name, InitFile)); err == nil {
		return "", true
	}
	return "", false
}

func loadOne(dir, name, suite string, tier Tier) (Case, error) {
	id := makeID(suite, name)
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
		Name:       name,
		Suite:      suite,
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
