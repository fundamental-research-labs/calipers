//go:build windows

package excel

import (
	"fmt"
	"path/filepath"

	"golang.org/x/sys/windows/registry"
)

const wefCatalogKey = `Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\`

func registerTrustedCatalog(catalogDir string) error {
	abs, err := filepath.Abs(catalogDir)
	if err != nil {
		return err
	}
	keyPath := wefCatalogKey + CatalogGUID
	k, _, err := registry.CreateKey(registry.CURRENT_USER, keyPath, registry.ALL_ACCESS)
	if err != nil {
		return fmt.Errorf("WEF trusted catalog: %w", err)
	}
	defer k.Close()
	if err := k.SetStringValue("Url", CatalogFileURL(abs)); err != nil {
		return fmt.Errorf("WEF trusted catalog Url: %w", err)
	}
	if err := k.SetDWordValue("Flags", 1); err != nil {
		return err
	}
	if err := k.SetDWordValue("ShowInMenu", 1); err != nil {
		return err
	}
	return nil
}
