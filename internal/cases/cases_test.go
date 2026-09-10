package cases

import (
	"archive/zip"
	"bytes"
	"io"
	"os"
	"path/filepath"
	"regexp"
	"runtime"
	"strings"
	"testing"

	"github.com/fundamental-research-labs/calipers/internal/excel"
	"github.com/fundamental-research-labs/calipers/internal/golden"
)

func TestLoadMissingScriptIsLoadSave(t *testing.T) {
	root := t.TempDir()
	writeCase(t, root, "noscript", "pk", nil)
	got, err := Load(root)
	if err != nil {
		t.Fatal(err)
	}
	if len(got) != 1 {
		t.Fatalf("len=%d, want 1", len(got))
	}
	c := got[0]
	if c.ID != "noscript" || c.Name != "noscript" {
		t.Fatalf("id/name = %s %s", c.ID, c.Name)
	}
	if c.RunScript() || c.ScriptPath != "" {
		t.Fatalf("missing script should be load+save, got ScriptPath=%q RunScript=%v", c.ScriptPath, c.RunScript())
	}
	if filepath.Base(c.InitPath) != InitFile {
		t.Fatalf("InitPath=%q", c.InitPath)
	}
	if filepath.Base(c.GoldenPath) != GoldenFile {
		t.Fatalf("GoldenPath=%q", c.GoldenPath)
	}
}

func TestLoadEmptyScriptIsLoadSave(t *testing.T) {
	root := t.TempDir()
	empty := ""
	writeCase(t, root, "emptyjs", "pk", &empty)
	got, err := Load(root)
	if err != nil {
		t.Fatal(err)
	}
	if len(got) != 1 {
		t.Fatalf("len=%d, want 1", len(got))
	}
	c := got[0]
	if c.RunScript() || c.ScriptPath != "" {
		t.Fatalf("empty script should be load+save, got ScriptPath=%q RunScript=%v", c.ScriptPath, c.RunScript())
	}
	if _, err := os.Stat(filepath.Join(c.Dir, ScriptFile)); err != nil {
		t.Fatalf("empty script.js should still exist on disk: %v", err)
	}
}

func TestLoadWhitespaceScriptIsLoadSave(t *testing.T) {
	root := t.TempDir()
	ws := "  \n\t\n"
	writeCase(t, root, "ws", "pk", &ws)
	got, err := Load(root)
	if err != nil {
		t.Fatal(err)
	}
	if len(got) != 1 {
		t.Fatalf("got %+v", got)
	}
	if got[0].RunScript() {
		t.Fatal("whitespace-only script should skip execution")
	}
}

func TestLoadMissingConfigIsOK(t *testing.T) {
	root := t.TempDir()
	js := "Excel.run(async () => {});\n"
	writeCase(t, root, "noconfig", "pk", &js)
	got, err := Load(root)
	if err != nil {
		t.Fatal(err)
	}
	if len(got) != 1 {
		t.Fatalf("len=%d", len(got))
	}
	c := got[0]
	if c.Budget != nil || c.ConfigPath != "" {
		t.Fatalf("missing config.json must leave Budget nil, got %+v path=%q", c.Budget, c.ConfigPath)
	}
}

func TestLoadConfigPresent(t *testing.T) {
	root := t.TempDir()
	js := "Excel.run(async () => {});\n"
	writeCase(t, root, "withcfg", "pk", &js)
	dir := filepath.Join(root, "withcfg")
	if err := WriteBudget(dir, Budget{MaxPeakMemoryBytes: 64 << 20, MaxDurationMs: 1500}); err != nil {
		t.Fatal(err)
	}
	got, err := Load(root)
	if err != nil {
		t.Fatal(err)
	}
	if len(got) != 1 {
		t.Fatalf("len=%d", len(got))
	}
	c := got[0]
	if c.ConfigPath != filepath.Join(dir, ConfigFile) {
		t.Fatalf("ConfigPath=%q", c.ConfigPath)
	}
	if c.Budget == nil {
		t.Fatal("Budget is nil")
	}
	if c.Budget.MaxPeakMemoryBytes != 64<<20 || c.Budget.MaxDurationMs != 1500 {
		t.Fatalf("budget = %+v", c.Budget)
	}
}

func TestLoadConfigPartialFields(t *testing.T) {
	root := t.TempDir()
	writeCase(t, root, "partial", "pk", nil)
	path := filepath.Join(root, "partial", ConfigFile)
	if err := os.WriteFile(path, []byte(`{"maxDurationMs": 800}`+"\n"), 0o644); err != nil {
		t.Fatal(err)
	}
	got, err := Load(root)
	if err != nil {
		t.Fatal(err)
	}
	if got[0].Budget == nil || got[0].Budget.MaxDurationMs != 800 || got[0].Budget.MaxPeakMemoryBytes != 0 {
		t.Fatalf("partial budget = %+v", got[0].Budget)
	}
}

func TestLoadConfigInvalidJSON(t *testing.T) {
	root := t.TempDir()
	writeCase(t, root, "bad", "pk", nil)
	if err := os.WriteFile(filepath.Join(root, "bad", ConfigFile), []byte("{"), 0o644); err != nil {
		t.Fatal(err)
	}
	_, err := Load(root)
	if err == nil || !strings.Contains(err.Error(), ConfigFile) {
		t.Fatalf("want %s parse error, got %v", ConfigFile, err)
	}
}

func TestLoadConfigEmptyFile(t *testing.T) {
	root := t.TempDir()
	writeCase(t, root, "emptycfg", "pk", nil)
	if err := os.WriteFile(filepath.Join(root, "emptycfg", ConfigFile), []byte("  \n"), 0o644); err != nil {
		t.Fatal(err)
	}
	_, err := Load(root)
	if err == nil || !strings.Contains(err.Error(), "empty") {
		t.Fatalf("want empty config error, got %v", err)
	}
}

func TestLoadConfigNegativeRejected(t *testing.T) {
	root := t.TempDir()
	writeCase(t, root, "neg", "pk", nil)
	if err := os.WriteFile(filepath.Join(root, "neg", ConfigFile), []byte(`{"maxDurationMs":-1}`), 0o644); err != nil {
		t.Fatal(err)
	}
	_, err := Load(root)
	if err == nil || !strings.Contains(err.Error(), "negative") {
		t.Fatalf("want negative budget error, got %v", err)
	}
}

func TestWriteBudgetRoundTrip(t *testing.T) {
	dir := t.TempDir()
	want := Budget{MaxPeakMemoryBytes: 1000, MaxDurationMs: 20}
	if err := WriteBudget(dir, want); err != nil {
		t.Fatal(err)
	}
	got, path, err := loadBudget(dir)
	if err != nil {
		t.Fatal(err)
	}
	if path != filepath.Join(dir, ConfigFile) {
		t.Fatalf("path=%q", path)
	}
	if got == nil || *got != want {
		t.Fatalf("got %+v, want %+v", got, want)
	}
}

func TestLoadNonEmptyScript(t *testing.T) {
	root := t.TempDir()
	js := "Excel.run(async () => { await Excel.run(); });\n"
	writeCase(t, root, "js", "pk", &js)
	got, err := Load(root)
	if err != nil {
		t.Fatal(err)
	}
	if len(got) != 1 {
		t.Fatalf("len=%d", len(got))
	}
	c := got[0]
	if !c.RunScript() {
		t.Fatal("non-empty script should run")
	}
	if filepath.Base(c.ScriptPath) != ScriptFile {
		t.Fatalf("ScriptPath=%q", c.ScriptPath)
	}
	data, err := os.ReadFile(c.ScriptPath)
	if err != nil {
		t.Fatal(err)
	}
	if string(data) != js {
		t.Fatalf("script body = %q, want %q", data, js)
	}
}

func TestOpenSavePassSkipsScripted(t *testing.T) {
	root := t.TempDir()
	js := "Excel.run(async (ctx) => {});"
	writeCase(t, root, "one", "pk", nil)
	writeCase(t, root, "js", "pk", &js)
	writeCase(t, root, "two", "pk", nil)
	all, err := Load(root)
	if err != nil {
		t.Fatal(err)
	}
	pass := OpenSavePass(all)
	if len(pass) != 2 {
		t.Fatalf("open+save pass len=%d, want 2", len(pass))
	}
	ids := map[string]bool{}
	for _, c := range pass {
		ids[c.ID] = true
		if c.RunScript() {
			t.Fatalf("open+save pass included scripted %s", c.ID)
		}
	}
	if !ids["one"] || !ids["two"] {
		t.Fatalf("open+save pass ids = %v", ids)
	}
}

func TestDefaultPassIncludesAll(t *testing.T) {
	root := t.TempDir()
	writeCase(t, root, "one", "pk", nil)
	writeCase(t, root, "two", "pk", nil)
	all, err := Load(root)
	if err != nil {
		t.Fatal(err)
	}
	if len(all) != 2 {
		t.Fatalf("len(all)=%d", len(all))
	}
	pass := DefaultPass(all)
	if len(pass) != 2 {
		t.Fatalf("default pass len=%d, want 2", len(pass))
	}
	ids := map[string]bool{pass[0].ID: true, pass[1].ID: true}
	if !ids["one"] || !ids["two"] {
		t.Fatalf("default pass ids = %v", ids)
	}
}

func TestSelectDefaultPassWhenNoIDs(t *testing.T) {
	root := t.TempDir()
	writeCase(t, root, "one", "pk", nil)
	writeCase(t, root, "two", "pk", nil)
	all, err := Load(root)
	if err != nil {
		t.Fatal(err)
	}
	got, err := Select(all, nil)
	if err != nil {
		t.Fatal(err)
	}
	if len(got) != 2 {
		t.Fatalf("len=%d, want default pass 2", len(got))
	}
}

func TestSelectByIDOrder(t *testing.T) {
	root := t.TempDir()
	writeCase(t, root, "one", "pk", nil)
	writeCase(t, root, "bomb", "pk", nil)
	all, err := Load(root)
	if err != nil {
		t.Fatal(err)
	}
	got, err := Select(all, []string{"bomb", "one"})
	if err != nil {
		t.Fatal(err)
	}
	if len(got) != 2 || got[0].ID != "bomb" || got[1].ID != "one" {
		t.Fatalf("got %+v", got)
	}
	_, err = Select(all, []string{"nope"})
	if err == nil || !bytes.Contains([]byte(err.Error()), []byte("unknown case")) {
		t.Fatalf("want unknown case, got %v", err)
	}
}

func TestLoadIgnoresNonCaseDirs(t *testing.T) {
	root := t.TempDir()
	if err := os.Mkdir(filepath.Join(root, "not_a_case"), 0o755); err != nil {
		t.Fatal(err)
	}
	writeCase(t, root, "ok", "pk", nil)
	got, err := Load(root)
	if err != nil {
		t.Fatal(err)
	}
	if len(got) != 1 || got[0].ID != "ok" {
		t.Fatalf("got %+v", got)
	}
}

func TestLoadSkipsUnderscorePrefixedSuites(t *testing.T) {
	root := t.TempDir()
	writeCase(t, filepath.Join(root, "default"), "ok", "pk", nil)
	writeCase(t, filepath.Join(root, "_disabled"), "bomb", "pk", nil)
	writeCase(t, filepath.Join(root, "_archive"), "old", "pk", nil)

	corpus, err := LoadCorpus(root)
	if err != nil {
		t.Fatal(err)
	}
	if len(corpus.Cases) != 1 || corpus.Cases[0].ID != "default/ok" {
		t.Fatalf("got %+v", corpus.Cases)
	}
	if strings.Join(corpus.Suites, ",") != "default" {
		t.Fatalf("suites=%v, want only default", corpus.Suites)
	}
}

func TestLoadIgnoresDirWithoutInit(t *testing.T) {
	root := t.TempDir()
	if err := os.Mkdir(filepath.Join(root, "missing_init"), 0o755); err != nil {
		t.Fatal(err)
	}
	got, err := Load(root)
	if err != nil {
		t.Fatal(err)
	}
	if len(got) != 0 {
		t.Fatalf("dir without init.xlsx must be ignored, got %+v", got)
	}
}

func TestLoadNestedSuites(t *testing.T) {
	root := t.TempDir()
	writeCase(t, filepath.Join(root, "roundtrip"), "rt", "pk", nil)
	writeCase(t, filepath.Join(root, "default"), "other", "pk", nil)
	writeCase(t, filepath.Join(root, "default"), "bomb", "pk", nil)
	if err := os.Mkdir(filepath.Join(root, "empty_suite"), 0o755); err != nil {
		t.Fatal(err)
	}

	corpus, err := LoadCorpus(root)
	if err != nil {
		t.Fatal(err)
	}
	if len(corpus.Cases) != 3 {
		t.Fatalf("len=%d, want 3", len(corpus.Cases))
	}
	wantSuites := []string{"default", "empty_suite", "roundtrip"}
	if strings.Join(corpus.Suites, ",") != strings.Join(wantSuites, ",") {
		t.Fatalf("suites=%v, want %v", corpus.Suites, wantSuites)
	}
	byID := map[string]Case{}
	for _, c := range corpus.Cases {
		byID[c.ID] = c
		if c.Suite == "" {
			t.Errorf("%s: suite should be set in nested layout", c.ID)
		}
		if c.ID != c.Suite+"/"+c.Name {
			t.Errorf("%s: ID must be suite/name, Name=%q Suite=%q", c.ID, c.Name, c.Suite)
		}
		if !strings.Contains(c.Dir, filepath.FromSlash(c.ID)) {
			t.Errorf("%s: Dir=%q does not include id", c.ID, c.Dir)
		}
	}
	if byID["roundtrip/rt"].Suite != "roundtrip" {
		t.Fatalf("roundtrip/rt suite=%q", byID["roundtrip/rt"].Suite)
	}
	if byID["default/other"].Suite != "default" || byID["default/bomb"].Suite != "default" {
		t.Fatalf("default-suite cases: %+v", byID)
	}

	got, err := FilterSuite(corpus.Cases, "roundtrip", corpus.Suites)
	if err != nil {
		t.Fatal(err)
	}
	if len(got) != 1 || got[0].ID != "roundtrip/rt" {
		t.Fatalf("roundtrip filter = %+v", got)
	}
	empty, err := FilterSuite(corpus.Cases, "empty_suite", corpus.Suites)
	if err != nil {
		t.Fatal(err)
	}
	if len(empty) != 0 {
		t.Fatalf("empty suite = %+v", empty)
	}
	_, err = FilterSuite(corpus.Cases, "nope", corpus.Suites)
	if err == nil || !bytes.Contains([]byte(err.Error()), []byte("unknown suite")) {
		t.Fatalf("want unknown suite, got %v", err)
	}

	pass := DefaultPass(corpus.Cases)
	if len(pass) != 3 {
		t.Fatalf("default pass across suites len=%d, want 3", len(pass))
	}
}

func TestSelectQualifiedID(t *testing.T) {
	root := t.TempDir()
	writeCase(t, filepath.Join(root, "roundtrip"), "one", "pk", nil)
	writeCase(t, filepath.Join(root, "default"), "bomb", "pk", nil)
	all, err := Load(root)
	if err != nil {
		t.Fatal(err)
	}
	got, err := Select(all, []string{"roundtrip/one", "default/bomb"})
	if err != nil {
		t.Fatal(err)
	}
	if len(got) != 2 || got[0].ID != "roundtrip/one" || got[1].ID != "default/bomb" {
		t.Fatalf("got %+v", got)
	}
	bare, err := Select(all, []string{"one"})
	if err != nil {
		t.Fatal(err)
	}
	if len(bare) != 1 || bare[0].ID != "roundtrip/one" {
		t.Fatalf("unique bare name = %+v", bare)
	}
}

func TestSelectAmbiguousID(t *testing.T) {
	root := t.TempDir()
	writeCase(t, filepath.Join(root, "roundtrip"), "one", "pk", nil)
	writeCase(t, filepath.Join(root, "default"), "one", "pk", nil)
	all, err := Load(root)
	if err != nil {
		t.Fatal(err)
	}
	_, err = Select(all, []string{"one"})
	if err == nil || !bytes.Contains([]byte(err.Error()), []byte("ambiguous case")) {
		t.Fatalf("want ambiguous case, got %v", err)
	}
	got, err := Select(all, []string{"default/one"})
	if err != nil {
		t.Fatal(err)
	}
	if len(got) != 1 || got[0].ID != "default/one" {
		t.Fatalf("got %+v", got)
	}
}

func TestPendingScriptedGoldens(t *testing.T) {
	root := t.TempDir()
	js := "await Excel.run(async (context) => { await context.sync(); });\n"
	writeCase(t, filepath.Join(root, "scratch"), "pending", "pk", &js)
	writeCase(t, filepath.Join(root, "scratch"), "done", "pk", &js)
	writeCase(t, filepath.Join(root, "roundtrip"), "rt", "pk", nil)

	all, err := Load(root)
	if err != nil {
		t.Fatal(err)
	}
	var done *Case
	for i := range all {
		if all[i].Name == "done" {
			done = &all[i]
			break
		}
	}
	if done == nil {
		t.Fatal("missing done")
	}
	if err := os.WriteFile(done.GoldenPath, []byte("pk"), 0o644); err != nil {
		t.Fatal(err)
	}
	all, err = Load(root)
	if err != nil {
		t.Fatal(err)
	}
	got := PendingScriptedGoldens(all)
	if len(got) != 1 || got[0].ID != "scratch/pending" {
		t.Fatalf("PendingScriptedGoldens = %+v, want scratch/pending", got)
	}
}

func TestMissingBudget(t *testing.T) {
	root := t.TempDir()
	writeCase(t, root, "plain", "pk", nil)
	writeCase(t, root, "capped", "pk", nil)
	if err := WriteBudget(filepath.Join(root, "capped"), Budget{MaxDurationMs: 10}); err != nil {
		t.Fatal(err)
	}
	all, err := Load(root)
	if err != nil {
		t.Fatal(err)
	}
	got := MissingBudget(all)
	if len(got) != 1 || got[0].ID != "plain" {
		t.Fatalf("MissingBudget = %+v, want plain", got)
	}
}

func TestGoldenComparePassSkipsMissingGolden(t *testing.T) {
	root := t.TempDir()
	js := "await Excel.run(async (context) => { await context.sync(); });\n"
	writeCase(t, filepath.Join(root, "roundtrip"), "rt", "pk", nil)
	writeCase(t, filepath.Join(root, "scratch"), "text", "pk", &js)

	all, err := Load(root)
	if err != nil {
		t.Fatal(err)
	}
	var rt *Case
	for i := range all {
		if all[i].ID == "roundtrip/rt" {
			rt = &all[i]
			break
		}
	}
	if rt == nil {
		t.Fatal("missing roundtrip/rt")
	}
	if err := os.WriteFile(rt.GoldenPath, []byte("pk"), 0o644); err != nil {
		t.Fatal(err)
	}
	def := DefaultPass(all)
	if len(def) != 2 {
		t.Fatalf("DefaultPass len=%d, want 2 (roundtrip + scratch without golden)", len(def))
	}
	got := GoldenComparePass(all)
	if len(got) != 1 || got[0].ID != "roundtrip/rt" {
		t.Fatalf("GoldenComparePass = %+v, want only the case with a golden", got)
	}
}

func TestLoadFlatLayoutLeavesSuiteEmpty(t *testing.T) {
	root := t.TempDir()
	writeCase(t, root, "flat", "pk", nil)
	if err := os.Mkdir(filepath.Join(root, "not_a_suite"), 0o755); err != nil {
		t.Fatal(err)
	}
	corpus, err := LoadCorpus(root)
	if err != nil {
		t.Fatal(err)
	}
	if len(corpus.Cases) != 1 || corpus.Cases[0].Suite != "" || corpus.Cases[0].ID != "flat" || corpus.Cases[0].Name != "flat" {
		t.Fatalf("got %+v", corpus)
	}
	if len(corpus.Suites) != 0 {
		t.Fatalf("flat layout should not report nested suites, got %v", corpus.Suites)
	}
}

func TestLoadRealCorpus(t *testing.T) {
	root := repoCasesDir(t)
	corpus, err := LoadCorpus(root)
	if err != nil {
		t.Fatal(err)
	}
	all := corpus.Cases
	if n := len(all); n < 193 || n > 213 {
		t.Fatalf("real corpus: got %d cases, want 193-213 (103 existing + 90-110 officejs)", n)
	}
	wantSuites := "default,officejs,roundtrip,scratch"
	if strings.Join(corpus.Suites, ",") != wantSuites {
		t.Fatalf("real corpus suites=%v, want %s", corpus.Suites, wantSuites)
	}
	disabled := filepath.Join(root, "_disabled")
	if _, err := os.Stat(disabled); err != nil {
		t.Fatalf("missing _disabled dir: %v", err)
	}
	for _, s := range corpus.Suites {
		if strings.HasPrefix(s, "_") {
			t.Errorf("underscore-prefixed directory %q must not be a suite", s)
		}
	}
	disabledEntries, err := os.ReadDir(disabled)
	if err != nil {
		t.Fatal(err)
	}
	nDisabled := 0
	for _, e := range disabledEntries {
		if !e.IsDir() {
			continue
		}
		nDisabled++
		if _, err := os.Stat(filepath.Join(disabled, e.Name(), GoldenFile)); err == nil {
			t.Errorf("_disabled/%s: must not have a golden", e.Name())
		}
	}
	if nDisabled != 8 {
		t.Fatalf("_disabled dirs: %d, want 8", nDisabled)
	}
	entries, err := os.ReadDir(root)
	if err != nil {
		t.Fatal(err)
	}
	for _, e := range entries {
		if e.IsDir() && isCaseDir(root, e.Name()) {
			t.Errorf("case %s must live under a suite directory, not %s", e.Name(), DirName)
		}
	}
	var nRT, nDef, nScratch, nOfficejs int
	for _, c := range all {
		switch c.Suite {
		case "roundtrip":
			nRT++
			if c.RunScript() {
				t.Errorf("%s: roundtrip suite must stay load+save, got script %q", c.ID, c.ScriptPath)
			}
		case "default":
			nDef++
		case "scratch":
			nScratch++
			if !c.RunScript() {
				t.Errorf("%s: scratch case must be scripted", c.ID)
			}
		case "officejs":
			nOfficejs++
			if !c.RunScript() {
				t.Errorf("%s: officejs case must be scripted", c.ID)
			}
		default:
			t.Errorf("%s: unexpected suite %q", c.ID, c.Suite)
		}
		if c.ID != c.Suite+"/"+c.Name {
			t.Errorf("%s: ID must be suite/name, Name=%q Suite=%q", c.ID, c.Name, c.Suite)
		}
		if c.Suite != "scratch" && c.Suite != "officejs" && !allowedDefaultScript(c.Name) && (c.RunScript() || c.ScriptPath != "") {
			t.Errorf("%s: unexpected script %q", c.ID, c.ScriptPath)
		}
		if _, err := os.Stat(c.InitPath); err != nil {
			t.Errorf("%s: init: %v", c.ID, err)
		}
		if filepath.Base(c.InitPath) != InitFile {
			t.Errorf("%s: init must be %s, got %s", c.ID, InitFile, c.InitPath)
		}
		if filepath.Base(c.GoldenPath) != GoldenFile {
			t.Errorf("%s: golden dest must be %s, got %s", c.ID, GoldenFile, c.GoldenPath)
		}
		wantDir := filepath.Join(root, filepath.FromSlash(c.ID))
		if c.Dir != wantDir {
			t.Errorf("%s: Dir=%q, want %q", c.ID, c.Dir, wantDir)
		}
	}
	if nRT != 83 || nDef != 5 || nScratch != 15 {
		t.Fatalf("suite counts roundtrip=%d default=%d scratch=%d, want 83/5/15", nRT, nDef, nScratch)
	}
	if nOfficejs < 90 || nOfficejs > 110 {
		t.Fatalf("officejs cases: %d, want 90-110", nOfficejs)
	}
	if len(all) != nRT+nDef+nScratch+nOfficejs {
		t.Fatalf("suite counts do not sum to corpus len %d", len(all))
	}

	var scripted []Case
	var simpleSetA1 *Case
	for i := range all {
		c := &all[i]
		if c.RunScript() {
			scripted = append(scripted, *c)
		}
		if c.ID == "default/simple_set_a1" {
			simpleSetA1 = c
		}
	}
	if len(scripted) != nDef+nScratch+nOfficejs {
		t.Fatalf("scripted cases: got %d, want %d (all default + scratch + officejs)", len(scripted), nDef+nScratch+nOfficejs)
	}
	if simpleSetA1 == nil || !simpleSetA1.RunScript() || simpleSetA1.Suite != "default" || simpleSetA1.Name != "simple_set_a1" {
		t.Fatalf("scripted case default/simple_set_a1 missing or not runnable")
	}
	s := *simpleSetA1
	if filepath.Base(s.ScriptPath) != ScriptFile {
		t.Fatalf("ScriptPath=%q", s.ScriptPath)
	}
	body, err := os.ReadFile(s.ScriptPath)
	if err != nil {
		t.Fatal(err)
	}
	if len(bytes.TrimSpace(body)) == 0 {
		t.Fatal("simple_set_a1 script.js is empty; discovery would skip it")
	}
	if !bytes.Contains(body, []byte("Excel.run")) {
		t.Fatalf("simple_set_a1 script.js is not Office.js Excel.run:\n%s", body)
	}
	var simple *Case
	for i := range all {
		if all[i].ID == "roundtrip/simple" {
			simple = &all[i]
			break
		}
	}
	if simple == nil {
		t.Fatal("missing load+save case roundtrip/simple")
	}
	named, err := Select(all, []string{"roundtrip/simple", "default/simple_set_a1"})
	if err != nil {
		t.Fatal(err)
	}
	if len(named) != 2 || named[0].ID != "roundtrip/simple" || named[1].ID != "default/simple_set_a1" {
		t.Fatalf("Select by suite/name = %+v", named)
	}
	if simple.Suite != "roundtrip" {
		t.Fatalf("simple suite=%q, want %s", simple.Suite, "roundtrip")
	}
	if simple.RunScript() || simple.ScriptPath != "" {
		t.Fatalf("simple must stay load+save, got ScriptPath=%q", simple.ScriptPath)
	}
	simpleInit, err := os.ReadFile(simple.InitPath)
	if err != nil {
		t.Fatal(err)
	}
	scriptedInit, err := os.ReadFile(s.InitPath)
	if err != nil {
		t.Fatal(err)
	}
	if !bytes.Equal(simpleInit, scriptedInit) {
		t.Fatal("simple_set_a1 init.xlsx must be a copy of simple")
	}

	pass := DefaultPass(all)
	if len(pass) != len(all) {
		t.Fatalf("default pass len=%d, want all %d", len(pass), len(all))
	}

	goldenPass := GoldenComparePass(all)
	pending := pendingGoldenIDs(all)
	if len(pending) != 0 {
		t.Fatalf("pending goldens: %d, want 0 (officejs goldens committed); first=%v", len(pending), pending)
	}
	if len(goldenPass) != len(all) {
		t.Fatalf("golden-compare pass len=%d, want all %d", len(goldenPass), len(all))
	}
	for _, c := range goldenPass {
		st, err := os.Stat(c.GoldenPath)
		if err != nil || st.Size() == 0 {
			t.Errorf("%s: GoldenComparePass included a case with no golden", c.ID)
		}
	}

	openSave := OpenSavePass(all)
	if len(openSave) != len(all)-len(scripted) {
		t.Fatalf("open+save pass len=%d, want %d (all minus scripted)", len(openSave), len(all)-len(scripted))
	}
	for _, c := range openSave {
		if c.RunScript() {
			t.Errorf("open+save pass included scripted %s", c.ID)
		}
		if c.Suite == "scratch" || c.Suite == "officejs" {
			t.Errorf("open+save pass included scripted suite %s", c.ID)
		}
	}

	byOfficejs, err := FilterSuite(all, "officejs", corpus.Suites)
	if err != nil {
		t.Fatal(err)
	}
	if len(byOfficejs) != nOfficejs {
		t.Fatalf("FilterSuite officejs len=%d, want %d", len(byOfficejs), nOfficejs)
	}

	bySuite, err := FilterSuite(all, "scratch", corpus.Suites)
	if err != nil {
		t.Fatal(err)
	}
	if len(bySuite) != nScratch {
		t.Fatalf("FilterSuite scratch len=%d, want %d", len(bySuite), nScratch)
	}
	for _, c := range bySuite {
		if c.Suite != "scratch" || !strings.HasPrefix(c.ID, "scratch/") {
			t.Errorf("FilterSuite scratch returned %s", c.ID)
		}
	}
}

func TestCommittedInitsAreWindowsExcel16Exports(t *testing.T) {
	root := repoCasesDir(t)
	all, err := Load(root)
	if err != nil {
		t.Fatal(err)
	}
	if len(all) == 0 {
		t.Fatal("no cases loaded")
	}
	for _, c := range all {
		if err := excel.CheckWindowsExcel16Export(c.InitPath); err != nil {
			t.Errorf("%s: %v", c.ID, err)
		}
	}
}

func TestCommittedOpenSaveGoldens(t *testing.T) {
	root := repoCasesDir(t)
	all, err := Load(root)
	if err != nil {
		t.Fatal(err)
	}

	var first golden.Meta
	var firstID string
	for i, c := range OpenSavePass(all) {
		st, err := os.Stat(c.GoldenPath)
		if err != nil {
			t.Errorf("%s: missing golden.xlsx", c.ID)
			continue
		}
		if st.Size() == 0 {
			t.Errorf("%s: golden.xlsx is empty", c.ID)
		}
		m, err := golden.Read(c.GoldenPath)
		if err != nil {
			t.Errorf("%s: %v", c.ID, err)
			continue
		}
		if m.Host != excel.HostID {
			t.Errorf("%s: host %q, want %s", c.ID, m.Host, excel.HostID)
		}
		if m.Script != "" {
			t.Errorf("%s: load+save golden must omit script, got %q", c.ID, m.Script)
		}
		if m.Input != InitFile {
			t.Errorf("%s: input %q, want %s", c.ID, m.Input, InitFile)
		}
		if m.ExcelVersion == "" || m.ExcelBuild == "" {
			t.Errorf("%s: missing Excel version/build in meta", c.ID)
		}
		if i == 0 {
			first = m
			firstID = c.ID
			continue
		}
		if m.Host != first.Host || m.ExcelVersion != first.ExcelVersion || m.ExcelBuild != first.ExcelBuild {
			t.Errorf("mixed excel goldens: %s is host=%s version=%s build=%s; %s is host=%s version=%s build=%s",
				c.ID, m.Host, m.ExcelVersion, m.ExcelBuild,
				firstID, first.Host, first.ExcelVersion, first.ExcelBuild)
		}
	}
}

func TestCommittedScriptedGolden(t *testing.T) {
	root := repoCasesDir(t)
	all, err := Load(root)
	if err != nil {
		t.Fatal(err)
	}

	var c *Case
	for i := range all {
		if all[i].ID == "default/simple_set_a1" {
			c = &all[i]
			break
		}
	}
	if c == nil || !c.RunScript() {
		t.Fatal("missing scripted case default/simple_set_a1")
	}

	for _, pass := range OpenSavePass(all) {
		if pass.ID == c.ID {
			t.Fatal("OpenSavePass must still skip the Office.js case")
		}
	}

	st, err := os.Stat(c.GoldenPath)
	if err != nil {
		t.Fatalf("committed scripted golden missing: %v", err)
	}
	if st.Size() == 0 {
		t.Fatal("committed scripted golden is empty")
	}
	m, err := golden.Read(c.GoldenPath)
	if err != nil {
		t.Fatal(err)
	}
	if m.Host != excel.HostID {
		t.Fatalf("host %q, want %s", m.Host, excel.HostID)
	}
	if m.Script != ScriptFile {
		t.Fatalf("script %q, want %s", m.Script, ScriptFile)
	}
	if m.Input != InitFile {
		t.Fatalf("input %q, want %s", m.Input, InitFile)
	}
	if m.ExcelVersion == "" || m.ExcelBuild == "" {
		t.Fatal("missing Excel version/build in scripted sidecar")
	}
	if m.GeneratedAt == "" {
		t.Fatal("missing generatedAt")
	}

	if !xlsxContains(t, c.GoldenPath, "calipers") {
		t.Fatal("scripted golden.xlsx must contain the Office.js value calipers")
	}
	if xlsxContains(t, c.InitPath, "calipers") {
		t.Fatal("simple_set_a1 init.xlsx must not already contain calipers")
	}
}

func TestScratchSuite(t *testing.T) {
	root := repoCasesDir(t)
	corpus, err := LoadCorpus(root)
	if err != nil {
		t.Fatal(err)
	}
	var foundScratch bool
	for _, s := range corpus.Suites {
		if s == "scratch" {
			foundScratch = true
			break
		}
	}
	if !foundScratch {
		t.Fatal("corpus is missing suite scratch")
	}

	scratch, err := FilterSuite(corpus.Cases, "scratch", corpus.Suites)
	if err != nil {
		t.Fatal(err)
	}
	if n := len(scratch); n < 15 {
		t.Fatalf("scratch cases: %d, want at least the original 15", n)
	}
	for _, c := range scratch {
		if c.Suite != "scratch" || !strings.HasPrefix(c.ID, "scratch/") {
			t.Errorf("FilterSuite scratch returned %s (suite=%q)", c.ID, c.Suite)
		}
	}

	emptyInit, err := os.ReadFile(filepath.Join(root, "roundtrip", "empty", InitFile))
	if err != nil {
		t.Fatal(err)
	}

	var bodies []byte
	for _, c := range scratch {
		if !c.RunScript() {
			t.Errorf("%s: scratch case must have a non-empty script.js", c.ID)
			continue
		}
		initData, err := os.ReadFile(c.InitPath)
		if err != nil {
			t.Errorf("%s: init: %v", c.ID, err)
			continue
		}
		if !bytes.Equal(initData, emptyInit) {
			t.Errorf("%s: init.xlsx must be a byte-identical copy of roundtrip/empty", c.ID)
		}
		if originalScratchNames[c.Name] {
			st, err := os.Stat(c.GoldenPath)
			if err != nil || st.Size() == 0 {
				t.Errorf("%s: original scratch case must have a committed golden.xlsx", c.ID)
			}
		}
		body, err := os.ReadFile(c.ScriptPath)
		if err != nil {
			t.Errorf("%s: script: %v", c.ID, err)
			continue
		}
		if len(bytes.TrimSpace(body)) == 0 {
			t.Errorf("%s: script.js is empty; discovery would skip it", c.ID)
			continue
		}
		if !bytes.Contains(body, []byte("Excel.run")) {
			t.Errorf("%s: script.js is not Office.js Excel.run:\n%s", c.ID, body)
		}
		bodies = append(bodies, body...)
		bodies = append(bodies, '\n')
	}

	joined := string(bodies)
	for _, needle := range []string{
		"pivotTables.add",
		"tables.add",
		"charts.add",
		"conditionalFormats",
		"columnWidth",
		"rowHeight",
	} {
		if !strings.Contains(joined, needle) {
			t.Errorf("scratch suite missing %s", needle)
		}
	}
	spill := regexp.MustCompile(`SEQUENCE|FILTER|UNIQUE|SORT|RANDARRAY`)
	if !spill.MatchString(joined) {
		t.Error("scratch suite missing a spilling dynamic-array formula (SEQUENCE/FILTER/UNIQUE/SORT/RANDARRAY)")
	}
	if !strings.Contains(joined, "freezePanes") && !strings.Contains(joined, "showGridlines") && !strings.Contains(joined, "tab.color") {
		t.Error("scratch suite missing a worksheet/workbook setting (freezePanes, showGridlines, or tab.color)")
	}
	if !regexp.MustCompile(`\.values\s*=\s*\[[\s\S]*?"[A-Za-z]`).Match(bodies) {
		t.Error("scratch suite missing text cell values")
	}
	if !regexp.MustCompile(`\.values\s*=\s*\[[\s\S]*?[0-9]`).Match(bodies) {
		t.Error("scratch suite missing numeric cell values")
	}

	for _, c := range OpenSavePass(corpus.Cases) {
		if c.Suite == "scratch" {
			t.Errorf("OpenSavePass included scratch %s", c.ID)
		}
	}
	golden := GoldenComparePass(corpus.Cases)
	var nScratchGolden int
	for _, c := range golden {
		if c.Suite == "scratch" {
			nScratchGolden++
		}
	}
	if nScratchGolden < 15 {
		t.Errorf("GoldenComparePass scratch count=%d, want at least the original 15 goldens", nScratchGolden)
	}
	def := DefaultPass(corpus.Cases)
	var nScratchDefault int
	for _, c := range def {
		if c.Suite == "scratch" {
			nScratchDefault++
		}
	}
	if nScratchDefault != len(scratch) {
		t.Fatalf("DefaultPass scratch count=%d, want %d (unprefixed scratch included)", nScratchDefault, len(scratch))
	}
	for name := range originalScratchNames {
		found := false
		for _, c := range scratch {
			if c.Name == name {
				found = true
				break
			}
		}
		if !found {
			t.Errorf("missing original scratch case %s", name)
		}
	}
}

func TestOfficejsSuite(t *testing.T) {
	root := repoCasesDir(t)
	corpus, err := LoadCorpus(root)
	if err != nil {
		t.Fatal(err)
	}
	var found bool
	for _, s := range corpus.Suites {
		if s == "officejs" {
			found = true
			break
		}
	}
	if !found {
		t.Fatal("corpus is missing suite officejs")
	}

	officejs, err := FilterSuite(corpus.Cases, "officejs", corpus.Suites)
	if err != nil {
		t.Fatal(err)
	}
	if n := len(officejs); n < 90 || n > 110 {
		t.Fatalf("officejs cases: %d, want 90-110", n)
	}

	emptyInit, err := os.ReadFile(filepath.Join(root, "roundtrip", "empty", InitFile))
	if err != nil {
		t.Fatal(err)
	}

	var bodies []byte
	for _, c := range officejs {
		if !c.RunScript() {
			t.Errorf("%s: officejs case must have a non-empty script.js", c.ID)
			continue
		}
		initData, err := os.ReadFile(c.InitPath)
		if err != nil {
			t.Errorf("%s: init: %v", c.ID, err)
			continue
		}
		if !bytes.Equal(initData, emptyInit) {
			t.Errorf("%s: init.xlsx must be a byte-identical copy of roundtrip/empty", c.ID)
		}
		st, err := os.Stat(c.GoldenPath)
		if err != nil || st.Size() == 0 {
			t.Errorf("%s: officejs case must have a committed golden.xlsx", c.ID)
			continue
		}
		m, err := golden.Read(c.GoldenPath)
		if err != nil {
			t.Errorf("%s: %v", c.ID, err)
			continue
		}
		if m.Host != excel.HostID {
			t.Errorf("%s: host %q, want %s", c.ID, m.Host, excel.HostID)
		}
		if m.Script != ScriptFile {
			t.Errorf("%s: script %q, want %s", c.ID, m.Script, ScriptFile)
		}
		if m.Input != InitFile {
			t.Errorf("%s: input %q, want %s", c.ID, m.Input, InitFile)
		}
		if c.Budget != nil || c.ConfigPath != "" {
			t.Errorf("%s: officejs case must not ship a budget config yet", c.ID)
		}
		if _, err := os.Stat(filepath.Join(c.Dir, ConfigFile)); err == nil {
			t.Errorf("%s: must not have %s until measure-budgets runs on Windows", c.ID, ConfigFile)
		}
		body, err := os.ReadFile(c.ScriptPath)
		if err != nil {
			t.Errorf("%s: script: %v", c.ID, err)
			continue
		}
		if len(bytes.TrimSpace(body)) == 0 {
			t.Errorf("%s: script.js is empty; discovery would skip it", c.ID)
			continue
		}
		if !bytes.Contains(body, []byte("Excel.run")) {
			t.Errorf("%s: script.js is not Office.js Excel.run:\n%s", c.ID, body)
		}
		bodies = append(bodies, body...)
		bodies = append(bodies, '\n')
	}

	joined := string(bodies)
	for _, needle := range []string{
		"tables.add",
		"pivotTables.add",
		"charts.add",
		"FILTER(",
		"UNIQUE(",
		"SORT(",
		"XLOOKUP(",
		"SEQUENCE(",
		"format.font",
		"format.fill",
		"numberFormat",
		"conditionalFormats",
		"dataValidation",
		"autoFilter",
		"comments.add",
		"hyperlink",
		"names.add",
		"freezePanes",
		"tabColor",
		".sort.apply",
		".insert(",
		"for (let i = 1; i <= 500; i++)",
		"TableStyleMedium2",
		"Excel.ChartType.pie",
		"Excel.ChartType.line",
		"Excel.ChartType.xyscatter",
		"columnHierarchies.add",
		"filterHierarchies.add",
		"showTotals",
		"style = \"Good\"",
	} {
		if !strings.Contains(joined, needle) {
			t.Errorf("officejs suite missing %s", needle)
		}
	}

	for _, c := range OpenSavePass(corpus.Cases) {
		if c.Suite == "officejs" {
			t.Errorf("OpenSavePass included officejs %s", c.ID)
		}
	}
	golden := GoldenComparePass(corpus.Cases)
	var nOfficejsGolden int
	for _, c := range golden {
		if c.Suite == "officejs" {
			nOfficejsGolden++
		}
	}
	if nOfficejsGolden != len(officejs) {
		t.Errorf("GoldenComparePass officejs count=%d, want %d", nOfficejsGolden, len(officejs))
	}
	def := DefaultPass(corpus.Cases)
	var nDefault int
	for _, c := range def {
		if c.Suite == "officejs" {
			nDefault++
		}
	}
	if nDefault != len(officejs) {
		t.Fatalf("DefaultPass officejs count=%d, want %d", nDefault, len(officejs))
	}
}

func TestLoadUnprefixedCase(t *testing.T) {
	root := t.TempDir()
	js := "await Excel.run(async (context) => { await context.sync(); });\n"
	writeCase(t, filepath.Join(root, "scratch"), "text", "pk", &js)
	if err := os.Mkdir(filepath.Join(root, "scratch", "not_a_case"), 0o755); err != nil {
		t.Fatal(err)
	}

	got, err := Load(root)
	if err != nil {
		t.Fatal(err)
	}
	if len(got) != 1 || got[0].ID != "scratch/text" || got[0].Name != "text" || got[0].Suite != "scratch" {
		t.Fatalf("got %+v", got)
	}
	if !got[0].RunScript() {
		t.Fatal("unprefixed scratch case should run its script")
	}
	pass := DefaultPass(got)
	if len(pass) != 1 || pass[0].ID != "scratch/text" {
		t.Fatalf("DefaultPass should include unprefixed scratch, got %+v", pass)
	}
}

func TestLoadFlatUnprefixedCases(t *testing.T) {
	root := t.TempDir()
	js := "Excel.run(async () => {});\n"
	writeCase(t, root, "text", "pk", &js)
	writeCase(t, root, "table", "pk", &js)
	if err := os.Mkdir(filepath.Join(root, "not_a_case"), 0o755); err != nil {
		t.Fatal(err)
	}

	corpus, err := LoadCorpus(root)
	if err != nil {
		t.Fatal(err)
	}
	if len(corpus.Suites) != 0 {
		t.Fatalf("flat unprefixed layout should not report suites, got %v", corpus.Suites)
	}
	if len(corpus.Cases) != 2 {
		t.Fatalf("len=%d, want 2", len(corpus.Cases))
	}
	byID := map[string]Case{}
	for _, c := range corpus.Cases {
		byID[c.ID] = c
		if c.Suite != "" {
			t.Errorf("%s: suite=%q, want empty", c.ID, c.Suite)
		}
	}
	if _, ok := byID["text"]; !ok {
		t.Fatalf("missing text: %+v", byID)
	}
	if _, ok := byID["table"]; !ok {
		t.Fatalf("missing table: %+v", byID)
	}
}

var originalScratchNames = map[string]bool{
	"text":                   true,
	"numbers":                true,
	"column_row_sizes":       true,
	"freeze_panes":           true,
	"spill":                  true,
	"named_range":            true,
	"merge":                  true,
	"number_format":          true,
	"add_sheet":              true,
	"table":                  true,
	"chart":                  true,
	"conditional_formatting": true,
	"pivot_table":            true,
	"data_validation":        true,
	"autofilter":             true,
}

func allowedDefaultScript(name string) bool {
	switch name {
	case "simple_set_a1",
		"names_add_defined_names_order",
		"add_sheet_sparse_ids",
		"chart_titles",
		"multi_row_formulas":
		return true
	default:
		return false
	}
}

func pendingGoldenIDs(all []Case) []string {
	var out []string
	for _, c := range all {
		st, err := os.Stat(c.GoldenPath)
		if err != nil || st.Size() == 0 {
			out = append(out, c.ID)
		}
	}
	return out
}

func xlsxContains(t *testing.T, path, needle string) bool {
	t.Helper()
	r, err := zip.OpenReader(path)
	if err != nil {
		t.Fatalf("zip %s: %v", path, err)
	}
	defer r.Close()
	want := []byte(needle)
	for _, f := range r.File {
		rc, err := f.Open()
		if err != nil {
			t.Fatal(err)
		}
		data, err := io.ReadAll(rc)
		_ = rc.Close()
		if err != nil {
			t.Fatal(err)
		}
		if bytes.Contains(data, want) {
			return true
		}
	}
	return false
}

func writeCase(t *testing.T, root, id, init string, script *string) {
	t.Helper()
	dir := filepath.Join(root, id)
	if err := os.MkdirAll(dir, 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(dir, InitFile), []byte(init), 0o644); err != nil {
		t.Fatal(err)
	}
	if script != nil {
		if err := os.WriteFile(filepath.Join(dir, ScriptFile), []byte(*script), 0o644); err != nil {
			t.Fatal(err)
		}
	}
}

func repoCasesDir(t *testing.T) string {
	t.Helper()
	_, file, _, ok := runtime.Caller(0)
	if !ok {
		t.Fatal("runtime.Caller failed")
	}
	dir := filepath.Join(filepath.Dir(file), "..", "..", DirName)
	abs, err := filepath.Abs(dir)
	if err != nil {
		t.Fatal(err)
	}
	if _, err := os.Stat(abs); err != nil {
		t.Fatalf("corpus %s: %v", abs, err)
	}
	return abs
}
