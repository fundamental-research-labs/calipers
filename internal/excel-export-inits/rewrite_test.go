package main

import (
	"archive/zip"
	"bytes"
	"errors"
	"io"
	"os"
	"path/filepath"
	"strings"
	"testing"
	"time"

	"github.com/fundamental-research-labs/calipers/internal/cases"
	"github.com/fundamental-research-labs/calipers/internal/excel"
)

type fakeHost struct {
	available bool
	openSave  func(in, out string) error
	calls     [][2]string
}

func (h *fakeHost) OpenSave(in, out string) error {
	h.calls = append(h.calls, [2]string{in, out})
	if h.openSave != nil {
		return h.openSave(in, out)
	}
	return nil
}

func (h *fakeHost) RunScript(_, _, _ string) error { return nil }

func (h *fakeHost) Info() (excel.HostInfo, error) {
	return excel.HostInfo{ID: excel.HostID}, nil
}

func (h *fakeHost) Available() bool { return h.available }

func TestRewriteRequiresAvailableHost(t *testing.T) {
	h := &fakeHost{available: false}
	err := rewriteInits(h, []cases.Case{{ID: "x", InitPath: "init.xlsx"}}, io.Discard)
	if !errors.Is(err, excel.ErrNotWindows) {
		t.Fatalf("error = %v, want ErrNotWindows", err)
	}
	if len(h.calls) != 0 {
		t.Fatalf("OpenSave called off-host: %v", h.calls)
	}
}

func TestRewriteSaveAsDistinctPathAndDedupCopies(t *testing.T) {
	root := t.TempDir()
	mac := namespacedAppXML("Microsoft Macintosh Excel", "16.0300")
	simple := writeZipBytes(t, map[string]string{"docProps/app.xml": mac, "xl/marker": "simple"})
	empty := writeZipBytes(t, map[string]string{"docProps/app.xml": mac, "xl/marker": "empty"})
	other := writeZipBytes(t, map[string]string{"docProps/app.xml": mac, "xl/marker": "other"})

	simpleRT := writeCaseInit(t, root, "roundtrip", "simple", simple)
	simpleDef := writeCaseInit(t, root, "default", "simple_set_a1", simple)
	emptyRT := writeCaseInit(t, root, "roundtrip", "empty", empty)
	scratchA := writeCaseInit(t, root, "scratch", "text", empty)
	scratchB := writeCaseInit(t, root, "scratch", "table", empty)
	otherRT := writeCaseInit(t, root, "roundtrip", "other", other)

	loaded, err := cases.Load(root)
	if err != nil {
		t.Fatal(err)
	}
	if len(loaded) != 6 {
		t.Fatalf("loaded %d, want 6", len(loaded))
	}

	h := &fakeHost{
		available: true,
		openSave: func(in, out string) error {
			if samePath(in, out) {
				t.Errorf("OpenSave in-place: %s", in)
			}
			src, err := os.ReadFile(in)
			if err != nil {
				return err
			}
			marker := zipPart(t, src, "xl/marker")
			data := writeZipBytes(t, map[string]string{
				"docProps/app.xml": namespacedAppXML(excel.WindowsExcelApplication, "16.0300"),
				"xl/marker":        "exported-" + marker,
			})
			if err := os.MkdirAll(filepath.Dir(out), 0o755); err != nil {
				return err
			}
			return os.WriteFile(out, data, 0o644)
		},
	}
	var buf bytes.Buffer
	if err := rewriteInits(h, loaded, &buf); err != nil {
		t.Fatal(err)
	}
	if n := len(h.calls); n != 3 {
		t.Fatalf("OpenSave calls = %d, want 3 unique (simple, empty, other); log:\n%s", n, buf.String())
	}
	for _, c := range h.calls {
		if samePath(c[0], c[1]) {
			t.Fatalf("Save As used the source path: %q -> %q", c[0], c[1])
		}
		if !strings.HasSuffix(strings.ToLower(c[1]), ".xlsx") {
			t.Fatalf("Save As dest is not xlsx: %q", c[1])
		}
		if strings.Contains(strings.ToLower(c[1]), "calipers") {
			t.Fatalf("Save As dest embeds calipers (Excel records absPath): %q", c[1])
		}
	}

	gotSimple, err := os.ReadFile(simpleRT)
	if err != nil {
		t.Fatal(err)
	}
	gotDef, err := os.ReadFile(simpleDef)
	if err != nil {
		t.Fatal(err)
	}
	if !bytes.Equal(gotSimple, gotDef) {
		t.Fatal("simple_set_a1 init must stay a byte copy of roundtrip/simple")
	}
	if err := excel.CheckWindowsExcel16Export(simpleRT); err != nil {
		t.Fatal(err)
	}
	if zipPart(t, gotSimple, "xl/marker") != "exported-simple" {
		t.Fatalf("simple marker = %q", zipPart(t, gotSimple, "xl/marker"))
	}

	gotEmpty, err := os.ReadFile(emptyRT)
	if err != nil {
		t.Fatal(err)
	}
	for _, p := range []string{scratchA, scratchB} {
		got, err := os.ReadFile(p)
		if err != nil {
			t.Fatal(err)
		}
		if !bytes.Equal(got, gotEmpty) {
			t.Fatalf("%s init must stay a byte copy of roundtrip/empty", p)
		}
	}
	if zipPart(t, gotEmpty, "xl/marker") != "exported-empty" {
		t.Fatalf("empty marker = %q", zipPart(t, gotEmpty, "xl/marker"))
	}

	if err := excel.CheckWindowsExcel16Export(otherRT); err != nil {
		t.Fatal(err)
	}
}

func TestRewriteLeavesOriginalsWhenExportIsNotExcel16(t *testing.T) {
	root := t.TempDir()
	mac := namespacedAppXML("Microsoft Macintosh Excel", "16.0300")
	orig := writeZipBytes(t, map[string]string{"docProps/app.xml": mac, "xl/marker": "keep"})
	initPath := writeCaseInit(t, root, "roundtrip", "simple", orig)
	loaded, err := cases.Load(root)
	if err != nil {
		t.Fatal(err)
	}

	h := &fakeHost{
		available: true,
		openSave: func(in, out string) error {
			if err := os.MkdirAll(filepath.Dir(out), 0o755); err != nil {
				return err
			}
			return os.WriteFile(out, orig, 0o644)
		},
	}
	err = rewriteInits(h, loaded, io.Discard)
	if err == nil || !strings.Contains(err.Error(), "not Windows Excel 16") {
		t.Fatalf("error = %v, want not Windows Excel 16", err)
	}
	got, err := os.ReadFile(initPath)
	if err != nil {
		t.Fatal(err)
	}
	if !bytes.Equal(got, orig) {
		t.Fatal("failed export must leave the original init in place")
	}
}

func TestRewriteLeavesOriginalsWhenLaterOpenSaveFails(t *testing.T) {
	root := t.TempDir()
	mac := namespacedAppXML("Microsoft Macintosh Excel", "16.0300")
	a := writeZipBytes(t, map[string]string{"docProps/app.xml": mac, "xl/marker": "a"})
	b := writeZipBytes(t, map[string]string{"docProps/app.xml": mac, "xl/marker": "b"})
	pathA := writeCaseInit(t, root, "roundtrip", "a", a)
	pathB := writeCaseInit(t, root, "roundtrip", "b", b)
	loaded, err := cases.Load(root)
	if err != nil {
		t.Fatal(err)
	}

	h := &fakeHost{
		available: true,
		openSave: func(in, out string) error {
			marker := zipPart(t, mustRead(t, in), "xl/marker")
			if marker == "b" {
				return errors.New("excel boom")
			}
			data := writeZipBytes(t, map[string]string{
				"docProps/app.xml": namespacedAppXML(excel.WindowsExcelApplication, "16.0300"),
				"xl/marker":        "exported-" + marker,
			})
			if err := os.MkdirAll(filepath.Dir(out), 0o755); err != nil {
				return err
			}
			return os.WriteFile(out, data, 0o644)
		},
	}
	err = rewriteInits(h, loaded, io.Discard)
	if err == nil || !strings.Contains(err.Error(), "excel boom") {
		t.Fatalf("error = %v, want excel boom", err)
	}
	if !bytes.Equal(mustRead(t, pathA), a) {
		t.Fatal("partial failure must not replace the first init")
	}
	if !bytes.Equal(mustRead(t, pathB), b) {
		t.Fatal("partial failure must not replace the second init")
	}
}

func writeCaseInit(t *testing.T, root, suite, name string, init []byte) string {
	t.Helper()
	dir := filepath.Join(root, suite, name)
	if err := os.MkdirAll(dir, 0o755); err != nil {
		t.Fatal(err)
	}
	p := filepath.Join(dir, cases.InitFile)
	if err := os.WriteFile(p, init, 0o644); err != nil {
		t.Fatal(err)
	}
	return p
}

func namespacedAppXML(app, ver string) string {
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
		`<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">` +
		`<Application>` + app + `</Application>` +
		`<AppVersion>` + ver + `</AppVersion>` +
		`</Properties>`
}

func writeZipBytes(t *testing.T, parts map[string]string) []byte {
	t.Helper()
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	for name, body := range parts {
		h := &zip.FileHeader{Name: name, Method: zip.Deflate, Modified: time.Time{}}
		w, err := zw.CreateHeader(h)
		if err != nil {
			t.Fatal(err)
		}
		if _, err := w.Write([]byte(body)); err != nil {
			t.Fatal(err)
		}
	}
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	return buf.Bytes()
}

func zipPart(t *testing.T, data []byte, name string) string {
	t.Helper()
	zr, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatal(err)
	}
	for _, f := range zr.File {
		if f.Name != name {
			continue
		}
		rc, err := f.Open()
		if err != nil {
			t.Fatal(err)
		}
		b, err := io.ReadAll(rc)
		_ = rc.Close()
		if err != nil {
			t.Fatal(err)
		}
		return string(b)
	}
	t.Fatalf("missing zip part %s", name)
	return ""
}

func mustRead(t *testing.T, path string) []byte {
	t.Helper()
	b, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	return b
}
