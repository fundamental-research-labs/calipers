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

const (
	wefDeveloperKey = `Software\Microsoft\Office\16.0\WEF\Developer`
	wefCatalogsRoot = `Software\Microsoft\Office\16.0\WEF\TrustedCatalogs`
)

// SideloadReg is the WEF registry payload for the sideloaded add-in.
// Developer matches webextension store=developer storeType=Registry.
// TrustedCatalogs matches Microsoft's Shared Folder catalog script ({GUID} key, Id, UNC Url).
type SideloadReg struct {
	DeveloperKey   string // HKCU\...\WEF\Developer
	DeveloperName  string // add-in GUID (manifest Id)
	DeveloperValue string // absolute path to calipers-runner.xml
	CatalogKey     string // HKCU\...\TrustedCatalogs\{CatalogGUID}
	CatalogId      string // {CatalogGUID}
	CatalogURL     string // UNC folder, e.g. \\localhost\C$\...\
	CatalogFlags   uint32
}

// NewSideloadReg builds registry values for a catalog directory that already
// contains calipers-runner.xml.
func NewSideloadReg(catalogDir string) (SideloadReg, error) {
	abs, err := filepath.Abs(catalogDir)
	if err != nil {
		return SideloadReg{}, err
	}
	return SideloadReg{
		DeveloperKey:   wefDeveloperKey,
		DeveloperName:  AddinID,
		DeveloperValue: filepath.Join(abs, catalogManifestName),
		CatalogKey:     wefCatalogsRoot + `\{` + CatalogGUID + `}`,
		CatalogId:      `{` + CatalogGUID + `}`,
		CatalogURL:     CatalogUNC(abs),
		CatalogFlags:   1,
	}, nil
}

// CatalogUNC is the TrustedCatalogs Url: a UNC share path, not file://.
// Local NT paths become \\localhost\<drive>$\<rest>\ (ADMIN$ style).
func CatalogUNC(dir string) string {
	s := strings.ReplaceAll(dir, "/", `\`)
	if len(s) >= 2 && s[1] == ':' {
		drive := strings.ToUpper(s[:1])
		rest := strings.TrimPrefix(s[2:], `\`)
		if rest != "" && !strings.HasSuffix(rest, `\`) {
			rest += `\`
		}
		if rest == "" {
			return `\\localhost\` + drive + `$\`
		}
		return `\\localhost\` + drive + `$\` + rest
	}
	abs, err := filepath.Abs(dir)
	if err == nil {
		s = strings.ReplaceAll(abs, "/", `\`)
	}
	s = strings.TrimPrefix(s, `\`)
	if !strings.HasSuffix(s, `\`) {
		s += `\`
	}
	return `\\localhost\` + s
}
