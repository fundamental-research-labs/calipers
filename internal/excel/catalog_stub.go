//go:build !windows

package excel

func registerSideload(string) error {
	return ErrNotWindows
}
