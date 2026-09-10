package main

import (
	"crypto/sha256"
	"encoding/hex"
	"errors"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"runtime"
	"strings"

	"github.com/fundamental-research-labs/calipers/internal/cases"
	"github.com/fundamental-research-labs/calipers/internal/excel"
)

type initGroup struct {
	hash  string
	src   string
	ids   []string
	dests []string
}

func rewriteInits(host excel.Host, loaded []cases.Case, w io.Writer) error {
	if !host.Available() {
		return excel.ErrNotWindows
	}
	if len(loaded) == 0 {
		return fmt.Errorf("no cases loaded")
	}

	groups, err := groupInits(loaded)
	if err != nil {
		return err
	}

	// Excel writes this directory into xl/workbook.xml x15ac:absPath.
	// The prefix must not contain "calipers": default/simple_set_a1's
	// Office.js sets A1 to that string, and the corpus test forbids it
	// already appearing anywhere in init.xlsx.
	tmpRoot, err := os.MkdirTemp("", "xlsx-export-inits-*")
	if err != nil {
		return err
	}
	defer os.RemoveAll(tmpRoot)

	exported := make([][]byte, len(groups))
	for i, g := range groups {
		fmt.Fprintf(w, "[%d/%d] %s", i+1, len(groups), g.ids[0])
		if len(g.ids) > 1 {
			fmt.Fprintf(w, " (+ %d copies: %s)", len(g.ids)-1, strings.Join(g.ids[1:], ", "))
		}
		fmt.Fprintln(w)
		if f, ok := w.(interface{ Sync() error }); ok {
			_ = f.Sync()
		}

		out := filepath.Join(tmpRoot, g.hash+".xlsx")
		if samePath(g.src, out) {
			return fmt.Errorf("%s: Save As dest must not be the source path", g.ids[0])
		}
		if err := host.OpenSave(g.src, out); err != nil {
			return fmt.Errorf("%s: OpenSave: %w", g.ids[0], err)
		}
		if err := excel.CheckWindowsExcel16Export(out); err != nil {
			return fmt.Errorf("%s: exported workbook is not Windows Excel 16: %w", g.ids[0], err)
		}
		data, err := os.ReadFile(out)
		if err != nil {
			return fmt.Errorf("%s: read export: %w", g.ids[0], err)
		}
		exported[i] = data
	}

	for i, g := range groups {
		for _, dest := range g.dests {
			if err := replaceFile(dest, exported[i]); err != nil {
				return fmt.Errorf("replace %s: %w", dest, err)
			}
		}
	}
	fmt.Fprintf(w, "excel-export-inits: %d unique, %d files\n", len(groups), len(loaded))
	return nil
}

func groupInits(loaded []cases.Case) ([]initGroup, error) {
	byHash := make(map[string]*initGroup, len(loaded))
	var order []string
	for _, c := range loaded {
		data, err := os.ReadFile(c.InitPath)
		if err != nil {
			return nil, fmt.Errorf("%s: read init: %w", c.ID, err)
		}
		sum := sha256.Sum256(data)
		h := hex.EncodeToString(sum[:])
		g, ok := byHash[h]
		if !ok {
			g = &initGroup{hash: h, src: c.InitPath}
			byHash[h] = g
			order = append(order, h)
		}
		g.ids = append(g.ids, c.ID)
		g.dests = append(g.dests, c.InitPath)
	}
	out := make([]initGroup, 0, len(order))
	for _, h := range order {
		out = append(out, *byHash[h])
	}
	return out, nil
}

func replaceFile(path string, data []byte) error {
	dir := filepath.Dir(path)
	tmp, err := os.CreateTemp(dir, ".init.xlsx.tmp-*")
	if err != nil {
		return err
	}
	tmpName := tmp.Name()
	if _, err := tmp.Write(data); err != nil {
		_ = tmp.Close()
		_ = os.Remove(tmpName)
		return err
	}
	if err := tmp.Close(); err != nil {
		_ = os.Remove(tmpName)
		return err
	}
	if err := os.Remove(path); err != nil && !errors.Is(err, os.ErrNotExist) {
		_ = os.Remove(tmpName)
		return err
	}
	if err := os.Rename(tmpName, path); err != nil {
		return err
	}
	return nil
}

func samePath(a, b string) bool {
	a, b = filepath.Clean(a), filepath.Clean(b)
	if runtime.GOOS == "windows" {
		return strings.EqualFold(a, b)
	}
	return a == b
}
