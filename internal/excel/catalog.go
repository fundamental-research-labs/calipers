package excel

import (
	"bytes"
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"text/template"
)

const catalogManifestName = "calipers-runner.xml"

// CatalogGUID is the WEF TrustedCatalogs subkey for this add-in.
const CatalogGUID = "5c1a1e15-0000-4000-a000-c01a1e1500ca"

type manifestData struct {
	ID      string
	BaseURL string
}

// RenderManifest fills the sideload manifest with the local add-in base URL.
func RenderManifest(baseURL string) ([]byte, error) {
	baseURL = strings.TrimRight(baseURL, "/")
	tmplBytes, err := webFS.ReadFile("web/manifest.xml.tmpl")
	if err != nil {
		return nil, fmt.Errorf("add-in manifest: %w", err)
	}
	tmpl, err := template.New("manifest").Parse(string(tmplBytes))
	if err != nil {
		return nil, fmt.Errorf("add-in manifest: %w", err)
	}
	var buf bytes.Buffer
	if err := tmpl.Execute(&buf, manifestData{ID: AddinID, BaseURL: baseURL}); err != nil {
		return nil, fmt.Errorf("add-in manifest: %w", err)
	}
	return buf.Bytes(), nil
}

// WriteSideloadCatalog writes the WEF catalog folder (manifest XML).
func WriteSideloadCatalog(dir, baseURL string) error {
	if err := os.MkdirAll(dir, 0o755); err != nil {
		return err
	}
	man, err := RenderManifest(baseURL)
	if err != nil {
		return err
	}
	return os.WriteFile(filepath.Join(dir, catalogManifestName), man, 0o644)
}

// CatalogFileURL is the WEF TrustedCatalogs Url value for a local folder.
func CatalogFileURL(dir string) string {
	abs, err := filepath.Abs(dir)
	if err != nil {
		abs = dir
	}
	abs = filepath.ToSlash(abs)
	if !strings.HasPrefix(abs, "/") {
		abs = "/" + abs
	}
	if !strings.HasSuffix(abs, "/") {
		abs += "/"
	}
	return "file://" + abs
}
