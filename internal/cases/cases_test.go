package cases

import (
	"archive/zip"
	"bytes"
	"io"
	"os"
	"path/filepath"
	"runtime"
	"strings"
	"testing"

	"github.com/fundamental-research-labs/calipers/internal/excel"
	"github.com/fundamental-research-labs/calipers/internal/golden"
)

func TestLoadMissingScriptIsLoadSave(t *testing.T) {
	root := t.TempDir()
	writeCase(t, root, "tier_a_noscript", "pk", nil)
	got, err := Load(root)
	if err != nil {
		t.Fatal(err)
	}
	if len(got) != 1 {
		t.Fatalf("len=%d, want 1", len(got))
	}
	c := got[0]
	if c.ID != "tier_a_noscript" || c.Name != "tier_a_noscript" || c.Tier != TierA {
		t.Fatalf("id/name/tier = %s %s %s", c.ID, c.Name, c.Tier)
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
	writeCase(t, root, "tier_a_emptyjs", "pk", &empty)
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
	writeCase(t, root, "tier_b_ws", "pk", &ws)
	got, err := Load(root)
	if err != nil {
		t.Fatal(err)
	}
	if len(got) != 1 || got[0].Tier != TierB {
		t.Fatalf("got %+v", got)
	}
	if got[0].RunScript() {
		t.Fatal("whitespace-only script should skip execution")
	}
}

func TestLoadNonEmptyScript(t *testing.T) {
	root := t.TempDir()
	js := "Excel.run(async () => { await Excel.run(); });\n"
	writeCase(t, root, "tier_a_js", "pk", &js)
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

func TestOpenSavePassSkipsScriptedAndTierC(t *testing.T) {
	root := t.TempDir()
	js := "Excel.run(async (ctx) => {});"
	writeCase(t, root, "tier_a_one", "pk", nil)
	writeCase(t, root, "tier_a_js", "pk", &js)
	writeCase(t, root, "tier_b_two", "pk", nil)
	writeCase(t, root, "tier_c_bomb", "pk", nil)
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
		if c.Tier == TierC {
			t.Fatalf("open+save pass included hostile %s", c.ID)
		}
		if c.RunScript() {
			t.Fatalf("open+save pass included scripted %s", c.ID)
		}
	}
	if !ids["tier_a_one"] || !ids["tier_b_two"] {
		t.Fatalf("open+save pass ids = %v", ids)
	}
}

func TestDefaultPassExcludesTierC(t *testing.T) {
	root := t.TempDir()
	writeCase(t, root, "tier_a_one", "pk", nil)
	writeCase(t, root, "tier_b_two", "pk", nil)
	writeCase(t, root, "tier_c_bomb", "pk", nil)
	all, err := Load(root)
	if err != nil {
		t.Fatal(err)
	}
	if len(all) != 3 {
		t.Fatalf("len(all)=%d", len(all))
	}
	pass := DefaultPass(all)
	if len(pass) != 2 {
		t.Fatalf("default pass len=%d, want 2", len(pass))
	}
	for _, c := range pass {
		if c.Tier == TierC {
			t.Fatalf("default pass included %s", c.ID)
		}
	}
	ids := map[string]bool{pass[0].ID: true, pass[1].ID: true}
	if !ids["tier_a_one"] || !ids["tier_b_two"] {
		t.Fatalf("default pass ids = %v", ids)
	}
}

func TestSelectDefaultPassWhenNoIDs(t *testing.T) {
	root := t.TempDir()
	writeCase(t, root, "tier_a_one", "pk", nil)
	writeCase(t, root, "tier_b_two", "pk", nil)
	writeCase(t, root, "tier_c_bomb", "pk", nil)
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
	for _, c := range got {
		if c.Tier == TierC {
			t.Fatalf("default select included %s", c.ID)
		}
	}
}

func TestSelectByIDIncludesTierCAndOrder(t *testing.T) {
	root := t.TempDir()
	writeCase(t, root, "tier_a_one", "pk", nil)
	writeCase(t, root, "tier_c_bomb", "pk", nil)
	all, err := Load(root)
	if err != nil {
		t.Fatal(err)
	}
	got, err := Select(all, []string{"tier_c_bomb", "tier_a_one"})
	if err != nil {
		t.Fatal(err)
	}
	if len(got) != 2 || got[0].ID != "tier_c_bomb" || got[1].ID != "tier_a_one" {
		t.Fatalf("got %+v", got)
	}
	_, err = Select(all, []string{"tier_a_nope"})
	if err == nil || !bytes.Contains([]byte(err.Error()), []byte("unknown case")) {
		t.Fatalf("want unknown case, got %v", err)
	}
}

func TestLoadIgnoresNonCaseDirs(t *testing.T) {
	root := t.TempDir()
	if err := os.Mkdir(filepath.Join(root, "not_a_case"), 0o755); err != nil {
		t.Fatal(err)
	}
	writeCase(t, root, "tier_a_ok", "pk", nil)
	got, err := Load(root)
	if err != nil {
		t.Fatal(err)
	}
	if len(got) != 1 || got[0].ID != "tier_a_ok" {
		t.Fatalf("got %+v", got)
	}
}

func TestLoadRequiresInit(t *testing.T) {
	root := t.TempDir()
	if err := os.Mkdir(filepath.Join(root, "tier_a_missing_init"), 0o755); err != nil {
		t.Fatal(err)
	}
	_, err := Load(root)
	if err == nil {
		t.Fatal("expected error for tier dir without init.xlsx")
	}
}

func TestLoadNestedSuites(t *testing.T) {
	root := t.TempDir()
	writeCase(t, filepath.Join(root, SuiteRoundtrip), "tier_a_rt", "pk", nil)
	writeCase(t, filepath.Join(root, SuiteDefault), "tier_a_other", "pk", nil)
	writeCase(t, filepath.Join(root, SuiteDefault), "tier_c_bomb", "pk", nil)
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
	wantSuites := []string{SuiteDefault, "empty_suite", SuiteRoundtrip}
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
	if byID["roundtrip/tier_a_rt"].Suite != SuiteRoundtrip {
		t.Fatalf("roundtrip/tier_a_rt suite=%q", byID["roundtrip/tier_a_rt"].Suite)
	}
	if byID["default/tier_a_other"].Suite != SuiteDefault || byID["default/tier_c_bomb"].Suite != SuiteDefault {
		t.Fatalf("default-suite cases: %+v", byID)
	}

	got, err := FilterSuite(corpus.Cases, SuiteRoundtrip, corpus.Suites)
	if err != nil {
		t.Fatal(err)
	}
	if len(got) != 1 || got[0].ID != "roundtrip/tier_a_rt" {
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
	if len(pass) != 2 {
		t.Fatalf("default pass across suites len=%d, want 2", len(pass))
	}
}

func TestSelectQualifiedID(t *testing.T) {
	root := t.TempDir()
	writeCase(t, filepath.Join(root, SuiteRoundtrip), "tier_a_one", "pk", nil)
	writeCase(t, filepath.Join(root, SuiteDefault), "tier_c_bomb", "pk", nil)
	all, err := Load(root)
	if err != nil {
		t.Fatal(err)
	}
	got, err := Select(all, []string{"roundtrip/tier_a_one", "default/tier_c_bomb"})
	if err != nil {
		t.Fatal(err)
	}
	if len(got) != 2 || got[0].ID != "roundtrip/tier_a_one" || got[1].ID != "default/tier_c_bomb" {
		t.Fatalf("got %+v", got)
	}
	bare, err := Select(all, []string{"tier_a_one"})
	if err != nil {
		t.Fatal(err)
	}
	if len(bare) != 1 || bare[0].ID != "roundtrip/tier_a_one" {
		t.Fatalf("unique bare name = %+v", bare)
	}
}

func TestSelectAmbiguousID(t *testing.T) {
	root := t.TempDir()
	writeCase(t, filepath.Join(root, SuiteRoundtrip), "tier_a_one", "pk", nil)
	writeCase(t, filepath.Join(root, SuiteDefault), "tier_a_one", "pk", nil)
	all, err := Load(root)
	if err != nil {
		t.Fatal(err)
	}
	_, err = Select(all, []string{"tier_a_one"})
	if err == nil || !bytes.Contains([]byte(err.Error()), []byte("ambiguous case")) {
		t.Fatalf("want ambiguous case, got %v", err)
	}
	got, err := Select(all, []string{"default/tier_a_one"})
	if err != nil {
		t.Fatal(err)
	}
	if len(got) != 1 || got[0].ID != "default/tier_a_one" {
		t.Fatalf("got %+v", got)
	}
}

func TestLoadFlatLayoutLeavesSuiteEmpty(t *testing.T) {
	root := t.TempDir()
	writeCase(t, root, "tier_a_flat", "pk", nil)
	if err := os.Mkdir(filepath.Join(root, "not_a_suite"), 0o755); err != nil {
		t.Fatal(err)
	}
	corpus, err := LoadCorpus(root)
	if err != nil {
		t.Fatal(err)
	}
	if len(corpus.Cases) != 1 || corpus.Cases[0].Suite != "" || corpus.Cases[0].ID != "tier_a_flat" || corpus.Cases[0].Name != "tier_a_flat" {
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
	if len(all) != 90 {
		t.Fatalf("real corpus: got %d cases, want 90", len(all))
	}
	if strings.Join(corpus.Suites, ",") != SuiteDefault+","+SuiteRoundtrip {
		t.Fatalf("real corpus suites=%v, want [%s %s]", corpus.Suites, SuiteDefault, SuiteRoundtrip)
	}
	entries, err := os.ReadDir(root)
	if err != nil {
		t.Fatal(err)
	}
	for _, e := range entries {
		if e.IsDir() {
			if _, ok := parseTier(e.Name()); ok {
				t.Errorf("case %s must live under a suite directory, not %s", e.Name(), DirName)
			}
		}
	}
	var nA, nB, nC, nRT, nDef int
	for _, c := range all {
		switch c.Tier {
		case TierA:
			nA++
		case TierB:
			nB++
		case TierC:
			nC++
		default:
			t.Errorf("%s: unexpected tier %q", c.ID, c.Tier)
		}
		switch c.Suite {
		case SuiteRoundtrip:
			nRT++
			if c.RunScript() {
				t.Errorf("%s: roundtrip suite must stay load+save, got script %q", c.ID, c.ScriptPath)
			}
		case SuiteDefault:
			nDef++
		default:
			t.Errorf("%s: unexpected suite %q", c.ID, c.Suite)
		}
		if c.ID != c.Suite+"/"+c.Name {
			t.Errorf("%s: ID must be suite/name, Name=%q Suite=%q", c.ID, c.Name, c.Suite)
		}
		if c.Name != "tier_a_simple_set_a1" && (c.RunScript() || c.ScriptPath != "") {
			t.Errorf("%s: unexpected script %q (only default/tier_a_simple_set_a1 is scripted)", c.ID, c.ScriptPath)
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
	if nA != 57 || nB != 25 || nC != 8 {
		t.Fatalf("tier counts a=%d b=%d c=%d, want 57/25/8", nA, nB, nC)
	}
	if nRT != 81 || nDef != 9 {
		t.Fatalf("suite counts roundtrip=%d default=%d, want 81/9", nRT, nDef)
	}

	var scripted []Case
	for _, c := range all {
		if c.RunScript() {
			scripted = append(scripted, c)
		}
	}
	if len(scripted) != 1 {
		t.Fatalf("scripted cases: got %d, want 1", len(scripted))
	}
	s := scripted[0]
	if s.ID != "default/tier_a_simple_set_a1" || s.Suite != SuiteDefault || s.Name != "tier_a_simple_set_a1" {
		t.Fatalf("scripted case = %s, want default/tier_a_simple_set_a1", s.ID)
	}
	if filepath.Base(s.ScriptPath) != ScriptFile {
		t.Fatalf("ScriptPath=%q", s.ScriptPath)
	}
	body, err := os.ReadFile(s.ScriptPath)
	if err != nil {
		t.Fatal(err)
	}
	if len(bytes.TrimSpace(body)) == 0 {
		t.Fatal("tier_a_simple_set_a1 script.js is empty; discovery would skip it")
	}
	if !bytes.Contains(body, []byte("Excel.run")) {
		t.Fatalf("tier_a_simple_set_a1 script.js is not Office.js Excel.run:\n%s", body)
	}
	var simple *Case
	for i := range all {
		if all[i].ID == "roundtrip/tier_a_simple" {
			simple = &all[i]
			break
		}
	}
	if simple == nil {
		t.Fatal("missing load+save case roundtrip/tier_a_simple")
	}
	named, err := Select(all, []string{"roundtrip/tier_a_simple", "default/tier_a_simple_set_a1"})
	if err != nil {
		t.Fatal(err)
	}
	if len(named) != 2 || named[0].ID != "roundtrip/tier_a_simple" || named[1].ID != "default/tier_a_simple_set_a1" {
		t.Fatalf("Select by suite/name = %+v", named)
	}
	if simple.Suite != SuiteRoundtrip {
		t.Fatalf("tier_a_simple suite=%q, want %s", simple.Suite, SuiteRoundtrip)
	}
	if simple.RunScript() || simple.ScriptPath != "" {
		t.Fatalf("tier_a_simple must stay load+save, got ScriptPath=%q", simple.ScriptPath)
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
		t.Fatal("tier_a_simple_set_a1 init.xlsx must be a copy of tier_a_simple")
	}

	pass := DefaultPass(all)
	if len(pass) != nA+nB {
		t.Fatalf("default pass len=%d, want %d (exclude tier_c)", len(pass), nA+nB)
	}
	for _, c := range pass {
		if c.Tier == TierC {
			t.Errorf("default pass included hostile %s", c.ID)
		}
	}
	if len(pass) != len(all)-nC {
		t.Fatalf("default pass %d vs all-c %d", len(pass), len(all)-nC)
	}

	openSave := OpenSavePass(all)
	if len(openSave) != nA+nB-len(scripted) {
		t.Fatalf("open+save pass len=%d, want %d (default minus scripted)", len(openSave), nA+nB-len(scripted))
	}
	for _, c := range openSave {
		if c.RunScript() {
			t.Errorf("open+save pass included scripted %s", c.ID)
		}
		if c.Tier == TierC {
			t.Errorf("open+save pass included hostile %s", c.ID)
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

	for _, c := range all {
		if c.Tier != TierC {
			continue
		}
		if _, err := os.Stat(c.GoldenPath); err == nil {
			t.Errorf("%s: must not have a golden (tier_c)", c.ID)
		}
	}
}

func TestCommittedScriptedGolden(t *testing.T) {
	root := repoCasesDir(t)
	all, err := Load(root)
	if err != nil {
		t.Fatal(err)
	}

	var scripted []Case
	for _, c := range all {
		if c.RunScript() {
			scripted = append(scripted, c)
		}
	}
	if len(scripted) != 1 || scripted[0].ID != "default/tier_a_simple_set_a1" {
		t.Fatalf("scripted = %v, want only default/tier_a_simple_set_a1", scripted)
	}
	c := scripted[0]

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
		t.Fatal("tier_a_simple_set_a1 init.xlsx must not already contain calipers")
	}
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
