package excel

import (
	"os"
	"path/filepath"
	"strings"
	"testing"
)

func TestRenderManifest(t *testing.T) {
	man, err := RenderManifest("http://127.0.0.1:3456/")
	if err != nil {
		t.Fatal(err)
	}
	s := string(man)
	for _, need := range []string{
		AddinID,
		"http://127.0.0.1:3456/taskpane.html",
		"ReadWriteDocument",
		"Office.AutoShowTaskpaneWithDocument",
		"Office.js",
	} {
		if !strings.Contains(s, need) {
			t.Fatalf("manifest missing %q:\n%s", need, s)
		}
	}
	if strings.Contains(s, "{{.") {
		t.Fatal("unexpanded template")
	}
	if !strings.Contains(s, "<Version>1.0.0</Version>") {
		t.Fatal("RenderManifest should fill Version")
	}
	if !strings.Contains(s, "Runs corpus Office.js") {
		t.Fatal("manifest should describe Office.js runner")
	}
}

func TestWriteSideloadCatalog(t *testing.T) {
	dir := t.TempDir()
	if err := WriteSideloadCatalog(dir, "http://127.0.0.1:9"); err != nil {
		t.Fatal(err)
	}
	data, err := os.ReadFile(filepath.Join(dir, catalogManifestName))
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(string(data), "http://127.0.0.1:9/taskpane.html") {
		t.Fatalf("catalog manifest: %s", data)
	}
	reg, err := NewSideloadReg(dir)
	if err != nil {
		t.Fatal(err)
	}
	if reg.DeveloperKey != wefDeveloperKey {
		t.Fatalf("DeveloperKey = %q", reg.DeveloperKey)
	}
	if reg.DeveloperName != AddinID {
		t.Fatalf("DeveloperName = %q, want add-in Id %q", reg.DeveloperName, AddinID)
	}
	if filepath.Base(reg.DeveloperValue) != catalogManifestName {
		t.Fatalf("DeveloperValue = %q", reg.DeveloperValue)
	}
	wantKey := wefCatalogsRoot + `\{` + CatalogGUID + `}`
	if reg.CatalogKey != wantKey {
		t.Fatalf("CatalogKey = %q, want braced %q", reg.CatalogKey, wantKey)
	}
	if reg.CatalogId != `{`+CatalogGUID+`}` {
		t.Fatalf("CatalogId = %q", reg.CatalogId)
	}
	if strings.HasPrefix(reg.CatalogURL, "file:") || !strings.HasPrefix(reg.CatalogURL, `\\`) {
		t.Fatalf("CatalogURL must be UNC, got %q", reg.CatalogURL)
	}
	if reg.CatalogFlags != 1 {
		t.Fatalf("Flags = %d", reg.CatalogFlags)
	}
}

func TestPersistentCatalogDir(t *testing.T) {
	if os.Getenv("LOCALAPPDATA") == "" {
		t.Skip("LOCALAPPDATA unset")
	}
	dir, err := PersistentCatalogDir()
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(strings.ToLower(dir), `calipers\wef`) {
		t.Fatalf("PersistentCatalogDir = %q", dir)
	}
	st, err := os.Stat(dir)
	if err != nil || !st.IsDir() {
		t.Fatalf("catalog dir: %v", err)
	}
}

func TestCatalogUNCWindowsPath(t *testing.T) {
	got := CatalogUNC(`C:\Users\foo\calipers-wef`)
	want := `\\localhost\C$\Users\foo\calipers-wef\`
	if got != want {
		t.Fatalf("CatalogUNC = %q, want %q", got, want)
	}
	if strings.Contains(got, "file:") {
		t.Fatal("UNC must not be a file:// URL")
	}
}
