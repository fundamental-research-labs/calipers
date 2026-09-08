//go:build !windows

package excel

func registerSideload(string) error {
	return ErrNotWindows
}

func enableRuntimeLogging(string) error {
	return ErrNotWindows
}
