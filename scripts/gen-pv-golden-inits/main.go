// Derive cache-free init.xlsx files solely from the committed PV goldens.
// Run from the repository root; -check verifies without writing anything.
package main

import (
	"archive/zip"
	"bytes"
	"crypto/sha256"
	"encoding/json"
	"encoding/xml"
	"flag"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strconv"
	"strings"
)

const suiteDir = "verification/cases/pv_goldens"
const spreadsheetNS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

type fixture struct {
	ID     string `json:"id"`
	Golden string `json:"golden"`
	SHA256 string `json:"sha256"`
}
type manifest struct {
	Repository string    `json:"repository"`
	Revision   string    `json:"revision"`
	Root       string    `json:"root"`
	Fixtures   []fixture `json:"fixtures"`
}

func main() {
	check := flag.Bool("check", false, "check committed goldens and derived inputs without writing")
	flag.Parse()
	if err := run(suiteDir, *check); err != nil {
		fmt.Fprintln(os.Stderr, err)
		os.Exit(1)
	}
}

func readManifest(root string) (manifest, error) {
	var m manifest
	data, err := os.ReadFile(filepath.Join(root, "manifest.json"))
	if err != nil {
		return m, err
	}
	err = json.Unmarshal(data, &m)
	return m, err
}

func run(root string, check bool) error {
	m, err := readManifest(root)
	if err != nil {
		return err
	}
	if len(m.Fixtures) != 107 {
		return fmt.Errorf("want 107 goldens, got %d", len(m.Fixtures))
	}
	seen := map[string]bool{}
	for _, f := range m.Fixtures {
		if f.ID == "." || f.ID == ".." || strings.ContainsAny(f.ID, `/\`) || f.ID == "" || seen[f.ID] {
			return fmt.Errorf("invalid or duplicate fixture ID %q", f.ID)
		}
		seen[f.ID] = true
		if err := generate(root, f, check); err != nil {
			return fmt.Errorf("%s: %w", f.ID, err)
		}
	}
	action := "generated"
	if check {
		action = "checked"
	}
	fmt.Printf("pv_goldens: %d golden hashes and cache-free inputs %s\n", len(seen), action)
	return nil
}

func generate(root string, f fixture, check bool) error {
	dir := filepath.Join(root, f.ID)
	golden, err := os.ReadFile(filepath.Join(dir, "golden.xlsx"))
	if err != nil {
		return err
	}
	if got := fmt.Sprintf("%x", sha256.Sum256(golden)); got != f.SHA256 {
		return fmt.Errorf("golden SHA-256 %s differs from source %s", got, f.SHA256)
	}
	init, err := withoutCaches(golden)
	if err != nil {
		return err
	}
	dest := filepath.Join(dir, "init.xlsx")
	if check {
		committed, err := os.ReadFile(dest)
		if err != nil {
			return err
		}
		// Compare ZIP payloads, independent of compressor/Go version.
		return equalParts(init, committed)
	}
	return os.WriteFile(dest, init, 0644)
}

func withoutCaches(data []byte) ([]byte, error) {
	r, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		return nil, err
	}
	var out bytes.Buffer
	w := zip.NewWriter(&out)
	if err := w.SetComment(r.Comment); err != nil {
		return nil, err
	}
	for _, f := range r.File {
		if strings.HasPrefix(f.Name, "xl/worksheets/") && strings.HasSuffix(f.Name, ".xml") {
			data, err := readPart(f)
			if err != nil {
				return nil, err
			}
			stripped, err := stripWorksheet(data)
			if err != nil {
				return nil, fmt.Errorf("%s: %w", f.Name, err)
			}
			if !bytes.Equal(data, stripped) {
				h := f.FileHeader
				dest, err := w.CreateHeader(&h)
				if err != nil {
					return nil, err
				}
				if _, err := dest.Write(stripped); err != nil {
					return nil, err
				}
				continue
			}
		}
		// Unchanged ZIP entries retain even their compressed bytes and metadata.
		if err := w.Copy(f); err != nil {
			return nil, err
		}
	}
	if err := w.Close(); err != nil {
		return nil, err
	}
	return out.Bytes(), nil
}

type span struct{ start, end int64 }
type point struct{ col, row int }
type rectangle struct{ first, last point }
type cell struct {
	address string
	formula bool
	caches  []span
}

func attr(s xml.StartElement, name string) string {
	for _, a := range s.Attr {
		if a.Name.Local == name && a.Name.Space == "" {
			return a.Value
		}
	}
	return ""
}

// Use XML offsets to delete only value elements. Re-encoding XML can change
// namespace prefixes, formulas, and unrelated markup in these real workbooks.
func stripWorksheet(data []byte) ([]byte, error) {
	d := xml.NewDecoder(bytes.NewReader(data))
	var cells []cell
	var ranges []rectangle
	var current *cell
	var stack []xml.Name
	for {
		start := d.InputOffset()
		token, err := d.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return nil, err
		}
		switch s := token.(type) {
		case xml.StartElement:
			if s.Name.Space == spreadsheetNS && s.Name.Local == "c" && len(stack) == 3 &&
				stack[0].Local == "worksheet" && stack[1].Local == "sheetData" && stack[2].Local == "row" {
				cells = append(cells, cell{address: attr(s, "r")})
				current = &cells[len(cells)-1]
			}
			if current != nil && len(stack) == 4 && s.Name.Space == spreadsheetNS {
				switch s.Name.Local {
				case "f":
					current.formula = true
					// Array/spill and data-table followers can contain caches without <f>.
					// Shared-formula followers have their own <f>, and are handled directly.
					if ref := attr(s, "ref"); ref != "" && (attr(s, "t") == "array" || attr(s, "t") == "dataTable") {
						r, err := parseRange(ref)
						if err != nil {
							return nil, err
						}
						ranges = append(ranges, r)
					}
				case "v", "is":
					if err := d.Skip(); err != nil {
						return nil, err
					}
					current.caches = append(current.caches, span{start, d.InputOffset()})
					continue
				}
			}
			stack = append(stack, s.Name)
		case xml.EndElement:
			if current != nil && len(stack) == 4 && s.Name.Local == "c" {
				current = nil
			}
			stack = stack[:len(stack)-1]
		}
	}
	var out bytes.Buffer
	var last int64
	for _, c := range cells {
		derived := c.formula
		if !derived && len(ranges) > 0 {
			p, err := parsePoint(c.address)
			if err != nil {
				return nil, err
			}
			for _, r := range ranges {
				if p.col >= r.first.col && p.col <= r.last.col && p.row >= r.first.row && p.row <= r.last.row {
					derived = true
					break
				}
			}
		}
		if derived {
			for _, s := range c.caches {
				out.Write(data[last:s.start])
				last = s.end
			}
		}
	}
	out.Write(data[last:])
	return out.Bytes(), nil
}

func parsePoint(ref string) (point, error) {
	ref = strings.ReplaceAll(ref, "$", "")
	p := point{}
	i := 0
	for i < len(ref) && ref[i] >= 'A' && ref[i] <= 'Z' {
		p.col = p.col*26 + int(ref[i]-'A'+1)
		i++
	}
	row, err := strconv.Atoi(ref[i:])
	if err != nil || p.col < 1 || p.col > 16384 || row < 1 || row > 1048576 {
		return p, fmt.Errorf("invalid cell reference %q", ref)
	}
	p.row = row
	return p, nil
}

func parseRange(ref string) (rectangle, error) {
	ends := strings.Split(ref, ":")
	if len(ends) > 2 {
		return rectangle{}, fmt.Errorf("invalid formula range %q", ref)
	}
	first, err := parsePoint(ends[0])
	if err != nil {
		return rectangle{}, err
	}
	last, err := parsePoint(ends[len(ends)-1])
	if err != nil {
		return rectangle{}, err
	}
	if first.col > last.col || first.row > last.row {
		return rectangle{}, fmt.Errorf("reversed formula range %q", ref)
	}
	return rectangle{first, last}, nil
}

func readPart(f *zip.File) ([]byte, error) {
	r, err := f.Open()
	if err != nil {
		return nil, err
	}
	defer r.Close()
	return io.ReadAll(r)
}

func equalParts(expected, actual []byte) error {
	a, err := zip.NewReader(bytes.NewReader(expected), int64(len(expected)))
	if err != nil {
		return err
	}
	b, err := zip.NewReader(bytes.NewReader(actual), int64(len(actual)))
	if err != nil {
		return err
	}
	if len(a.File) != len(b.File) || a.Comment != b.Comment {
		return fmt.Errorf("init ZIP package differs")
	}
	for i, f := range a.File {
		if f.Name != b.File[i].Name {
			return fmt.Errorf("init ZIP member %d differs", i)
		}
		want, err := readPart(f)
		if err != nil {
			return err
		}
		got, err := readPart(b.File[i])
		if err != nil {
			return err
		}
		if !bytes.Equal(want, got) {
			return fmt.Errorf("%s: init is not golden with only formula caches removed", f.Name)
		}
	}
	return nil
}
