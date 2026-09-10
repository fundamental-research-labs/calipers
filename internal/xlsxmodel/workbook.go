// Package xlsxmodel parses xlsx archives into workbook semantics (sheets,
// cell values, date1904) and diffs those models. It is not ZIP-part or XML
// string equality. Types, formulas, styles, names, merges, and freeze are
// later PRs. calipers verify still gates on compare.Files.
package xlsxmodel

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"math"
	"os"
	"path"
	"sort"
	"strconv"
	"strings"
)

// Workbook is the PR1 semantic model.
type Workbook struct {
	Date1904 bool
	// Cells is sheet name → A1 → value.
	Cells map[string]map[string]Value
}

// Value is a resolved cell value. Numbers compare as IEEE-754 binary64.
type Value struct {
	Number bool
	F      float64
	S      string
}

// Diff is one semantic mismatch (not a ZIP part).
type Diff struct {
	Axis     string
	Location string
	Detail   string
}

// Result is the outcome of comparing two parsed workbooks.
type Result struct {
	Equal bool
	Diffs []Diff
}

// CompareFiles parses two xlsx files and diffs values / date1904.
func CompareFiles(exportPath, goldenPath string) (Result, error) {
	export, err := os.ReadFile(exportPath)
	if err != nil {
		return Result{}, fmt.Errorf("export: %w", err)
	}
	golden, err := os.ReadFile(goldenPath)
	if err != nil {
		return Result{}, fmt.Errorf("golden: %w", err)
	}
	return Compare(export, golden)
}

// Compare parses two xlsx blobs and diffs values / date1904.
func Compare(export, golden []byte) (Result, error) {
	a, err := Parse(export)
	if err != nil {
		return Result{}, fmt.Errorf("export xlsx: %w", err)
	}
	b, err := Parse(golden)
	if err != nil {
		return Result{}, fmt.Errorf("golden xlsx: %w", err)
	}
	return DiffWorkbooks(a, b), nil
}

// Parse reads an xlsx archive into sheets, A1 values, and date1904.
// Absent workbookPr@date1904 means the 1900 date system (not 1904).
func Parse(data []byte) (Workbook, error) {
	parts, err := readParts(data)
	if err != nil {
		return Workbook{}, err
	}
	sst := parseSST(parts["xl/sharedStrings.xml"])
	wb := Workbook{Cells: map[string]map[string]Value{}}
	if raw, ok := parts["xl/workbook.xml"]; ok {
		meta := parseWorkbook(raw)
		wb.Date1904 = meta.date1904
		paths := sheetPaths(meta.sheets, parts)
		for i, s := range meta.sheets {
			cells := parseSheet(parts[paths[i]], sst)
			wb.Cells[s.name] = cells
		}
		return wb, nil
	}
	names := worksheetNames(parts)
	if len(names) == 0 {
		return wb, nil
	}
	wb.Cells["Sheet1"] = parseSheet(parts[names[0]], sst)
	for i := 1; i < len(names); i++ {
		wb.Cells[fmt.Sprintf("Sheet%d", i+1)] = parseSheet(parts[names[i]], sst)
	}
	return wb, nil
}

// DiffWorkbooks compares date1904 and resolved cell values.
func DiffWorkbooks(a, b Workbook) Result {
	var diffs []Diff
	if a.Date1904 != b.Date1904 {
		diffs = append(diffs, Diff{
			Axis:   "date1904",
			Detail: fmt.Sprintf("expected %v got %v", b.Date1904, a.Date1904),
		})
	}
	type key struct{ sheet, ref string }
	seen := map[key]struct{}{}
	for sheet, cells := range a.Cells {
		for ref, av := range cells {
			seen[key{sheet, ref}] = struct{}{}
			bv, ok := b.Cells[sheet][ref]
			loc := sheet + "!" + ref
			if !ok {
				diffs = append(diffs, Diff{Axis: "values", Location: loc, Detail: "present in export, missing in golden"})
				continue
			}
			if d := valueDiff(loc, av, bv); d != nil {
				diffs = append(diffs, *d)
			}
		}
	}
	for sheet, cells := range b.Cells {
		for ref := range cells {
			k := key{sheet, ref}
			if _, ok := seen[k]; ok {
				continue
			}
			loc := sheet + "!" + ref
			diffs = append(diffs, Diff{Axis: "values", Location: loc, Detail: "present in golden, missing in export"})
		}
	}
	sort.Slice(diffs, func(i, j int) bool {
		if diffs[i].Axis != diffs[j].Axis {
			return diffs[i].Axis < diffs[j].Axis
		}
		return diffs[i].Location < diffs[j].Location
	})
	return Result{Equal: len(diffs) == 0, Diffs: diffs}
}

func valueDiff(loc string, a, b Value) *Diff {
	if a.Number && b.Number {
		if math.Float64bits(a.F) == math.Float64bits(b.F) {
			return nil
		}
		return &Diff{Axis: "values", Location: loc, Detail: fmt.Sprintf("expected %s got %s", b.S, a.S)}
	}
	if !a.Number && !b.Number && a.S == b.S {
		return nil
	}
	return &Diff{Axis: "values", Location: loc, Detail: fmt.Sprintf("expected %q got %q", b.S, a.S)}
}

func readParts(data []byte) (map[string][]byte, error) {
	r, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		return nil, err
	}
	out := make(map[string][]byte, len(r.File))
	for _, f := range r.File {
		name := strings.ReplaceAll(f.Name, "\\", "/")
		if strings.HasSuffix(name, "/") {
			continue
		}
		rc, err := f.Open()
		if err != nil {
			return nil, fmt.Errorf("%s: %w", name, err)
		}
		body, err := io.ReadAll(rc)
		_ = rc.Close()
		if err != nil {
			return nil, fmt.Errorf("%s: %w", name, err)
		}
		out[name] = body
	}
	return out, nil
}

type sheetRef struct {
	name string
	rid  string
}

type workbookMeta struct {
	date1904 bool
	sheets   []sheetRef
}

func parseWorkbook(raw []byte) workbookMeta {
	var meta workbookMeta
	dec := xml.NewDecoder(bytes.NewReader(raw))
	for {
		tok, err := dec.Token()
		if err != nil {
			break
		}
		se, ok := tok.(xml.StartElement)
		if !ok {
			continue
		}
		switch se.Name.Local {
		case "workbookPr":
			meta.date1904 = is1904(attr(se, "date1904"))
		case "sheet":
			meta.sheets = append(meta.sheets, sheetRef{name: attr(se, "name"), rid: attr(se, "id")})
		}
	}
	return meta
}

func is1904(v string) bool {
	switch strings.ToLower(strings.TrimSpace(v)) {
	case "1", "true":
		return true
	default:
		return false
	}
}

func sheetPaths(sheets []sheetRef, parts map[string][]byte) []string {
	rels := parseRels(parts["xl/_rels/workbook.xml.rels"])
	out := make([]string, len(sheets))
	files := worksheetNames(parts)
	for i, s := range sheets {
		if p, ok := rels[s.rid]; ok {
			if _, exists := parts[p]; exists {
				out[i] = p
				continue
			}
		}
		if i < len(files) {
			out[i] = files[i]
		}
	}
	return out
}

func parseRels(raw []byte) map[string]string {
	out := map[string]string{}
	if len(raw) == 0 {
		return out
	}
	dec := xml.NewDecoder(bytes.NewReader(raw))
	for {
		tok, err := dec.Token()
		if err != nil {
			break
		}
		se, ok := tok.(xml.StartElement)
		if !ok || se.Name.Local != "Relationship" {
			continue
		}
		id, typ, target := attr(se, "Id"), attr(se, "Type"), attr(se, "Target")
		if !strings.Contains(typ, "/worksheet") || strings.Contains(typ, "chartsheet") {
			continue
		}
		out[id] = xlPath(target)
	}
	return out
}

func xlPath(target string) string {
	t := strings.ReplaceAll(target, "\\", "/")
	t = strings.TrimPrefix(t, "/")
	if strings.HasPrefix(t, "xl/") {
		return t
	}
	return path.Join("xl", t)
}

func worksheetNames(parts map[string][]byte) []string {
	var names []string
	for n := range parts {
		if !strings.HasPrefix(n, "xl/worksheets/") || !strings.HasSuffix(n, ".xml") {
			continue
		}
		if strings.Contains(n, "/_rels/") {
			continue
		}
		names = append(names, n)
	}
	sort.Strings(names)
	return names
}

func parseSST(raw []byte) []string {
	if len(raw) == 0 {
		return nil
	}
	var sst []string
	var cur strings.Builder
	inSI, inT := false, false
	dec := xml.NewDecoder(bytes.NewReader(raw))
	for {
		tok, err := dec.Token()
		if err != nil {
			break
		}
		switch e := tok.(type) {
		case xml.StartElement:
			switch e.Name.Local {
			case "si":
				inSI = true
				cur.Reset()
			case "t":
				inT = true
			}
		case xml.EndElement:
			switch e.Name.Local {
			case "t":
				inT = false
			case "si":
				sst = append(sst, cur.String())
				inSI = false
			}
		case xml.CharData:
			if inSI && inT {
				cur.Write(e)
			}
		}
	}
	return sst
}

func parseSheet(raw []byte, sst []string) map[string]Value {
	cells := map[string]Value{}
	if len(raw) == 0 {
		return cells
	}
	dec := xml.NewDecoder(bytes.NewReader(raw))
	for {
		tok, err := dec.Token()
		if err != nil {
			break
		}
		se, ok := tok.(xml.StartElement)
		if !ok || se.Name.Local != "c" {
			continue
		}
		ref := attr(se, "r")
		if ref == "" {
			skip(dec)
			continue
		}
		t := attr(se, "t")
		vtext, ttext := readCell(dec)
		if val, ok := cellValue(t, vtext, ttext, sst); ok {
			cells[ref] = val
		}
	}
	return cells
}

func readCell(dec *xml.Decoder) (vtext, ttext string) {
	depth := 1
	inV, inT := false, false
	for depth > 0 {
		tok, err := dec.Token()
		if err != nil {
			return vtext, ttext
		}
		switch e := tok.(type) {
		case xml.StartElement:
			depth++
			switch e.Name.Local {
			case "v":
				inV = true
			case "t":
				inT = true
			}
		case xml.EndElement:
			switch e.Name.Local {
			case "v":
				inV = false
			case "t":
				inT = false
			}
			depth--
		case xml.CharData:
			if inV {
				vtext += string(e)
			}
			if inT {
				ttext += string(e)
			}
		}
	}
	return vtext, ttext
}

func skip(dec *xml.Decoder) {
	depth := 1
	for depth > 0 {
		tok, err := dec.Token()
		if err != nil {
			return
		}
		switch tok.(type) {
		case xml.StartElement:
			depth++
		case xml.EndElement:
			depth--
		}
	}
}

func cellValue(t, vtext, ttext string, sst []string) (Value, bool) {
	switch t {
	case "inlineStr":
		return Value{S: ttext}, true
	case "s":
		i, err := strconv.Atoi(strings.TrimSpace(vtext))
		if err != nil || i < 0 || i >= len(sst) {
			return Value{S: vtext}, true
		}
		return Value{S: sst[i]}, true
	case "str":
		return Value{S: vtext}, true
	default:
		// number (t missing or t="n"), or other v-bearing cells
		if vtext == "" && ttext == "" {
			return Value{}, false
		}
		if t == "n" || t == "" || t == "b" {
			if f, err := strconv.ParseFloat(strings.TrimSpace(vtext), 64); err == nil {
				return Value{Number: true, F: f, S: strings.TrimSpace(vtext)}, true
			}
		}
		if ttext != "" {
			return Value{S: ttext}, true
		}
		return Value{S: vtext}, true
	}
}

func attr(se xml.StartElement, local string) string {
	for _, a := range se.Attr {
		if a.Name.Local == local {
			return a.Value
		}
	}
	return ""
}
