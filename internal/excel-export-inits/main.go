// excel-export-inits rewrites verification-case init.xlsx files by opening
// each unique workbook in Windows Excel and Save As xlsx (format 51) to a
// new path, then replacing the inits. Byte-identical inits are exported
// once and copied (roundtrip/simple → default/simple_set_a1,
// roundtrip/empty → scratch/* and officejs/*).
//
//	go run ./internal/excel-export-inits [cases-dir]
//
// Requires Windows + Microsoft Excel. Off Windows this command errors with
// the same message as calipers excel-save.
package main

import (
	"fmt"
	"os"

	"github.com/fundamental-research-labs/calipers/internal/cases"
	"github.com/fundamental-research-labs/calipers/internal/excel"
)

func main() {
	if err := run(excel.NewHost(), os.Args[1:]); err != nil {
		fmt.Fprintf(os.Stderr, "excel-export-inits: %v\n", err)
		os.Exit(1)
	}
}

func run(host excel.Host, args []string) error {
	if len(args) > 0 && (args[0] == "-h" || args[0] == "--help") {
		fmt.Fprint(os.Stdout, usage)
		return nil
	}
	if !host.Available() {
		return excel.ErrNotWindows
	}
	root := cases.DirName
	switch len(args) {
	case 0:
	case 1:
		root = args[0]
	default:
		return fmt.Errorf("usage: excel-export-inits [cases-dir]")
	}
	loaded, err := cases.Load(root)
	if err != nil {
		return err
	}
	return rewriteInits(host, loaded, os.Stdout)
}

const usage = `excel-export-inits — rewrite case init.xlsx files via Windows Excel Save As

Usage:
  excel-export-inits [cases-dir]

Opens each unique loaded init.xlsx in Excel and Save As xlsx format 51 to a
temp path, then replaces the inits. Identical files are exported once and
copied. Default cases-dir is verification/cases.

Requires Windows + Microsoft Excel. Off Windows this errors:
  excel-save requires Windows + Excel (COM)
`
