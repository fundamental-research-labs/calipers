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
	url := CatalogFileURL(dir)
	if !strings.HasPrefix(url, "file://") || !strings.HasSuffix(url, "/") {
		t.Fatalf("CatalogFileURL = %q", url)
	}
}
