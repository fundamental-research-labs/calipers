//go:build !windows

package excel

func registerTrustedCatalog(string) error {
	return ErrNotWindows
}
