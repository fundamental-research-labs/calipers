// gen-officejs-cases writes verification/cases/officejs/<name>/{init.xlsx,script.js}
// from a spec table. init.xlsx is a byte copy of roundtrip/empty.
//
//	go run ./scripts/gen-officejs-cases [cases-dir]
//
// Re-run after editing the spec files. Does not write golden.xlsx or config.json.
package main

import (
	"fmt"
	"os"
	"path/filepath"
)

const usage = `gen-officejs-cases — write blank-init Office.js verification cases

Usage:
  gen-officejs-cases [cases-dir]

Copies roundtrip/empty/init.xlsx into cases-dir/officejs/<name>/init.xlsx
and writes script.js from the spec table. Default cases-dir is
verification/cases. Does not write golden.xlsx (Windows excel-run-pass)
or config.json (Windows measure-budgets).
`

func main() {
	if err := run(os.Args[1:]); err != nil {
		fmt.Fprintf(os.Stderr, "gen-officejs-cases: %v\n", err)
		os.Exit(1)
	}
}

func run(args []string) error {
	if len(args) > 0 && (args[0] == "-h" || args[0] == "--help") {
		fmt.Fprint(os.Stdout, usage)
		return nil
	}
	root := "verification/cases"
	switch len(args) {
	case 0:
	case 1:
		root = args[0]
	default:
		return fmt.Errorf("usage: gen-officejs-cases [cases-dir]")
	}
	empty := filepath.Join(root, "roundtrip", "empty", "init.xlsx")
	dest := filepath.Join(root, "officejs")
	n, err := generate(empty, dest)
	if err != nil {
		return err
	}
	fmt.Printf("wrote %d cases to %s\n", n, dest)
	return nil
}
