//go:build windows

package excel

import (
	"fmt"

	"golang.org/x/sys/windows/registry"
)

func registerSideload(catalogDir string) error {
	reg, err := NewSideloadReg(catalogDir)
	if err != nil {
		return err
	}
	dev, _, err := registry.CreateKey(registry.CURRENT_USER, reg.DeveloperKey, registry.ALL_ACCESS)
	if err != nil {
		return fmt.Errorf("WEF Developer: %w", err)
	}
	defer dev.Close()
	if err := dev.SetStringValue(reg.DeveloperName, reg.DeveloperValue); err != nil {
		return fmt.Errorf("WEF Developer manifest: %w", err)
	}

	cat, _, err := registry.CreateKey(registry.CURRENT_USER, reg.CatalogKey, registry.ALL_ACCESS)
	if err != nil {
		return fmt.Errorf("WEF trusted catalog: %w", err)
	}
	defer cat.Close()
	if err := cat.SetStringValue("Id", reg.CatalogId); err != nil {
		return fmt.Errorf("WEF trusted catalog Id: %w", err)
	}
	if err := cat.SetStringValue("Url", reg.CatalogURL); err != nil {
		return fmt.Errorf("WEF trusted catalog Url: %w", err)
	}
	if err := cat.SetDWordValue("Flags", reg.CatalogFlags); err != nil {
		return err
	}
	if err := cat.SetDWordValue("ShowInMenu", 1); err != nil {
		return err
	}
	return nil
}
