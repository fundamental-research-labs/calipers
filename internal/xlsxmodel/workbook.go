// Package xlsxmodel parses xlsx archives into workbook semantics (sheets,
// cell values, types, formulas, styles, names, merges, freeze, date1904) and
// diffs those models. It is not ZIP-part or XML string equality. calipers
// verify still gates on compare.Files.
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

const spreadsheetML = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

// Workbook is the semantic model.
type Workbook struct {
	Date1904 bool
	Sheets   []SheetInfo // name + visibility, workbook order (not rId/sheetId)
	Names    []DefinedName
	// Cells is sheet name → A1 → value.
	Cells  map[string]map[string]Value
	Merges map[string][]string // sheet → merge refs
	Freeze map[string]Freeze   // sheet → pane
}

// SheetInfo is sheet identity: name and visibility in workbook order.
type SheetInfo struct {
	Name  string
	State string // "", "hidden", "veryHidden"
}

// DefinedName is a workbook or sheet-local name.
type DefinedName struct {
	Name       string
	Formula    string
	LocalSheet string // localSheetId; empty = workbook scope
}

// Freeze is a frozen pane on one sheet.
type Freeze struct {
	State      string // frozen / frozenSplit
	XSplit     string
	YSplit     string
	TopLeft    string
	ActivePane string
}

// CellType is the resolved OOXML cell type.
type CellType string

const (
	TypeNumber CellType = "number"
	TypeString CellType = "string"
	TypeBool   CellType = "boolean"
	TypeError  CellType = "error"
)

// Value is a resolved cell: type, value, and formula text (if any).
// Numbers compare as IEEE-754 binary64. Formula is shared-expanded A1 text.
type Value struct {
	Type    CellType
	F       float64
	S       string
	Formula string
	FKind   string // "array" when <f t="array">; empty after shared expansion
	FRef    string
	Style   Style
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

// CompareFiles parses two xlsx files and diffs the semantic model.
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

// Compare parses two xlsx blobs and diffs the semantic model.
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

// Parse reads an xlsx archive into sheets, A1 values/types/formulas/styles, and date1904.
// Absent workbookPr@date1904 means the 1900 date system (not 1904).
func Parse(data []byte) (Workbook, error) {
	parts, err := readParts(data)
	if err != nil {
		return Workbook{}, err
	}
	sst := parseSST(parts["xl/sharedStrings.xml"])
	styles := parseStyleBook(parts["xl/styles.xml"], parts["xl/theme/theme1.xml"])
	wb := Workbook{
		Cells:  map[string]map[string]Value{},
		Merges: map[string][]string{},
		Freeze: map[string]Freeze{},
	}
	if raw, ok := parts["xl/workbook.xml"]; ok {
		meta := parseWorkbook(raw)
		wb.Date1904 = meta.date1904
		wb.Names = meta.names
		paths := sheetPaths(meta.sheets, parts)
		for i, s := range meta.sheets {
			wb.Sheets = append(wb.Sheets, SheetInfo{Name: s.name, State: s.state})
			ps := parseSheet(parts[paths[i]], sst, styles)
			wb.Cells[s.name] = ps.cells
			if len(ps.merges) > 0 {
				wb.Merges[s.name] = ps.merges
			}
			if ps.freeze != (Freeze{}) {
				wb.Freeze[s.name] = ps.freeze
			}
		}
		return wb, nil
	}
	files := worksheetNames(parts)
	if len(files) == 0 {
		return wb, nil
	}
	for i, path := range files {
		name := "Sheet1"
		if i > 0 {
			name = fmt.Sprintf("Sheet%d", i+1)
		}
		wb.Sheets = append(wb.Sheets, SheetInfo{Name: name})
		ps := parseSheet(parts[path], sst, styles)
		wb.Cells[name] = ps.cells
		if len(ps.merges) > 0 {
			wb.Merges[name] = ps.merges
		}
		if ps.freeze != (Freeze{}) {
			wb.Freeze[name] = ps.freeze
		}
	}
	return wb, nil
}

// DiffWorkbooks compares the semantic model.
func DiffWorkbooks(a, b Workbook) Result {
	var diffs []Diff
	if a.Date1904 != b.Date1904 {
		diffs = append(diffs, Diff{
			Axis:   "date1904",
			Detail: fmt.Sprintf("expected %v got %v", b.Date1904, a.Date1904),
		})
	}
	diffs = append(diffs, sheetDiffs(a.Sheets, b.Sheets)...)
	diffs = append(diffs, nameDiffs(a.Names, b.Names)...)
	type key struct{ sheet, ref string }
	seen := map[key]struct{}{}
	for sheet, cells := range a.Cells {
		for ref, av := range cells {
			seen[key{sheet, ref}] = struct{}{}
			bv, ok := b.Cells[sheet][ref]
			loc := sheet + "!" + ref
			if !ok {
				diffs = append(diffs, missingDiffs(loc, av, "present in export, missing in golden")...)
				continue
			}
			diffs = append(diffs, cellDiffs(loc, av, bv)...)
		}
	}
	for sheet, cells := range b.Cells {
		for ref, bv := range cells {
			k := key{sheet, ref}
			if _, ok := seen[k]; ok {
				continue
			}
			loc := sheet + "!" + ref
			diffs = append(diffs, missingDiffs(loc, bv, "present in golden, missing in export")...)
		}
	}
	diffs = append(diffs, mergeDiffs(a.Merges, b.Merges)...)
	diffs = append(diffs, freezeDiffs(a.Freeze, b.Freeze)...)
	sort.Slice(diffs, func(i, j int) bool {
		if diffs[i].Axis != diffs[j].Axis {
			return diffs[i].Axis < diffs[j].Axis
		}
		return diffs[i].Location < diffs[j].Location
	})
	return Result{Equal: len(diffs) == 0, Diffs: diffs}
}

func missingDiffs(loc string, v Value, detail string) []Diff {
	var diffs []Diff
	if v.Type != "" {
		diffs = append(diffs, Diff{Axis: "values", Location: loc, Detail: detail})
	}
	if v.Formula != "" || v.FKind != "" {
		diffs = append(diffs, Diff{Axis: "formulas", Location: loc, Detail: detail})
	}
	if len(diffs) == 0 {
		diffs = append(diffs, Diff{Axis: "values", Location: loc, Detail: detail})
	}
	return diffs
}

func cellDiffs(loc string, a, b Value) []Diff {
	var diffs []Diff
	if a.Type != b.Type {
		diffs = append(diffs, Diff{Axis: "types", Location: loc, Detail: fmt.Sprintf("expected %s got %s", b.Type, a.Type)})
	} else if d := valueDiff(loc, a, b); d != nil {
		diffs = append(diffs, *d)
	}
	if a.Formula != b.Formula || a.FKind != b.FKind || a.FRef != b.FRef {
		diffs = append(diffs, Diff{Axis: "formulas", Location: loc, Detail: fmt.Sprintf("expected %q got %q", formulaKey(b), formulaKey(a))})
	}
	if a.Style != b.Style {
		diffs = append(diffs, Diff{Axis: "styles", Location: loc, Detail: fmt.Sprintf("expected %s got %s", b.Style, a.Style)})
	}
	return diffs
}

func sheetDiffs(a, b []SheetInfo) []Diff {
	if sheetKey(a) == sheetKey(b) {
		return nil
	}
	return []Diff{{Axis: "sheets", Detail: fmt.Sprintf("expected %s got %s", sheetKey(b), sheetKey(a))}}
}

func sheetKey(ss []SheetInfo) string {
	parts := make([]string, len(ss))
	for i, s := range ss {
		st := s.State
		if st == "" {
			st = "visible"
		}
		parts[i] = s.Name + ":" + st
	}
	return strings.Join(parts, ",")
}

func nameDiffs(a, b []DefinedName) []Diff {
	am, bm := nameMap(a), nameMap(b)
	var diffs []Diff
	for k, af := range am {
		bf, ok := bm[k]
		if !ok {
			diffs = append(diffs, Diff{Axis: "names", Location: k, Detail: "present in export, missing in golden"})
			continue
		}
		if af != bf {
			diffs = append(diffs, Diff{Axis: "names", Location: k, Detail: fmt.Sprintf("expected %q got %q", bf, af)})
		}
	}
	for k := range bm {
		if _, ok := am[k]; !ok {
			diffs = append(diffs, Diff{Axis: "names", Location: k, Detail: "present in golden, missing in export"})
		}
	}
	return diffs
}

func nameMap(ns []DefinedName) map[string]string {
	out := map[string]string{}
	for _, n := range ns {
		k := n.Name
		if n.LocalSheet != "" {
			k += "@" + n.LocalSheet
		}
		out[k] = n.Formula
	}
	return out
}

func mergeDiffs(a, b map[string][]string) []Diff {
	am, bm := mergeSet(a), mergeSet(b)
	var diffs []Diff
	for p := range am {
		if !bm[p] {
			diffs = append(diffs, Diff{Axis: "merges", Location: p.sheet + "!" + p.ref, Detail: "present in export, missing in golden"})
		}
	}
	for p := range bm {
		if !am[p] {
			diffs = append(diffs, Diff{Axis: "merges", Location: p.sheet + "!" + p.ref, Detail: "present in golden, missing in export"})
		}
	}
	return diffs
}

func mergeSet(m map[string][]string) map[struct{ sheet, ref string }]bool {
	out := map[struct{ sheet, ref string }]bool{}
	for sheet, refs := range m {
		for _, ref := range refs {
			out[struct{ sheet, ref string }{sheet, ref}] = true
		}
	}
	return out
}

func freezeDiffs(a, b map[string]Freeze) []Diff {
	seen := map[string]struct{}{}
	var diffs []Diff
	for sheet, af := range a {
		seen[sheet] = struct{}{}
		bf := b[sheet]
		if af != bf {
			diffs = append(diffs, Diff{Axis: "freeze", Location: sheet, Detail: fmt.Sprintf("expected %+v got %+v", bf, af)})
		}
	}
	for sheet, bf := range b {
		if _, ok := seen[sheet]; ok {
			continue
		}
		if bf != (Freeze{}) {
			diffs = append(diffs, Diff{Axis: "freeze", Location: sheet, Detail: fmt.Sprintf("expected %+v got none", bf)})
		}
	}
	return diffs
}

func formulaKey(v Value) string {
	if v.FKind == "" && v.FRef == "" {
		return v.Formula
	}
	return v.FKind + ":" + v.FRef + ":" + v.Formula
}

func valueDiff(loc string, a, b Value) *Diff {
	if a.Type == TypeNumber && b.Type == TypeNumber {
		if math.Float64bits(a.F) == math.Float64bits(b.F) {
			return nil
		}
		return &Diff{Axis: "values", Location: loc, Detail: fmt.Sprintf("expected %s got %s", b.S, a.S)}
	}
	if a.Type == TypeBool && b.Type == TypeBool {
		if a.F == b.F {
			return nil
		}
		return &Diff{Axis: "values", Location: loc, Detail: fmt.Sprintf("expected %s got %s", b.S, a.S)}
	}
	if a.S == b.S {
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
	name, rid, state string
}

type workbookMeta struct {
	date1904 bool
	sheets   []sheetRef
	names    []DefinedName
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
			if se.Name.Space != "" && se.Name.Space != spreadsheetML {
				continue
			}
			if v := attr(se, "date1904"); v != "" {
				meta.date1904 = is1904(v)
			}
		case "sheet":
			meta.sheets = append(meta.sheets, sheetRef{name: attr(se, "name"), rid: attr(se, "id"), state: attr(se, "state")})
		case "definedName":
			meta.names = append(meta.names, DefinedName{
				Name:       attr(se, "name"),
				LocalSheet: attr(se, "localSheetId"),
				Formula:    readElementText(dec),
			})
		}
	}
	return meta
}

func readElementText(dec *xml.Decoder) string {
	var b strings.Builder
	depth := 1
	for depth > 0 {
		tok, err := dec.Token()
		if err != nil {
			break
		}
		switch e := tok.(type) {
		case xml.StartElement:
			depth++
		case xml.EndElement:
			depth--
		case xml.CharData:
			if depth == 1 {
				b.Write(e)
			}
		}
	}
	return b.String()
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

type rawCell struct {
	ref, t, s, v, inline string
	f, fType, fRef, fSi  string
}

type parsedSheet struct {
	cells  map[string]Value
	merges []string
	freeze Freeze
}

func parseSheet(raw []byte, sst []string, styles *styleBook) parsedSheet {
	out := parsedSheet{cells: map[string]Value{}}
	raws, merges, freeze := collectSheet(raw)
	out.merges = merges
	out.freeze = freeze
	masters := map[string]rawCell{}
	for _, c := range raws {
		if c.fType == "shared" && c.f != "" {
			masters[c.fSi] = c
		}
	}
	for _, c := range raws {
		val, ok := cellValue(c.t, c.v, c.inline, sst)
		formula, kind, aref := resolveFormula(c, masters)
		st := styles.resolve(c.s)
		if !ok && formula == "" && kind == "" && st.zero() {
			continue
		}
		val.Formula = formula
		val.FKind = kind
		val.FRef = aref
		val.Style = st
		out.cells[c.ref] = val
	}
	return out
}

func collectSheet(raw []byte) (cells []rawCell, merges []string, freeze Freeze) {
	if len(raw) == 0 {
		return nil, nil, Freeze{}
	}
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
		case "c":
			ref := attr(se, "r")
			if ref == "" {
				skip(dec)
				continue
			}
			c := rawCell{ref: ref, t: attr(se, "t"), s: attr(se, "s")}
			c.v, c.inline, c.f, c.fType, c.fRef, c.fSi = readCell(dec)
			cells = append(cells, c)
		case "mergeCell":
			if ref := attr(se, "ref"); ref != "" {
				merges = append(merges, ref)
			}
		case "pane":
			st := attr(se, "state")
			if st == "frozen" || st == "frozenSplit" {
				freeze = Freeze{
					State:      st,
					XSplit:     attr(se, "xSplit"),
					YSplit:     attr(se, "ySplit"),
					TopLeft:    attr(se, "topLeftCell"),
					ActivePane: attr(se, "activePane"),
				}
			}
		}
	}
	return cells, merges, freeze
}

func resolveFormula(c rawCell, masters map[string]rawCell) (formula, kind, aref string) {
	switch c.fType {
	case "shared":
		m, ok := masters[c.fSi]
		src := c.f
		from := c.ref
		if src == "" && ok {
			src = m.f
			from = m.ref
		}
		dCol, dRow := deltaA1(from, c.ref)
		return shiftFormula(src, dCol, dRow), "", ""
	case "array":
		return c.f, "array", c.fRef
	default:
		return c.f, "", ""
	}
}

func readCell(dec *xml.Decoder) (vtext, ttext, ftext, fType, fRef, fSi string) {
	depth := 1
	inV, inT, inF := false, false, false
	for depth > 0 {
		tok, err := dec.Token()
		if err != nil {
			return
		}
		switch e := tok.(type) {
		case xml.StartElement:
			depth++
			switch e.Name.Local {
			case "v":
				inV = true
			case "t":
				inT = true
			case "f":
				inF = true
				fType = attr(e, "t")
				fRef = attr(e, "ref")
				fSi = attr(e, "si")
			}
		case xml.EndElement:
			switch e.Name.Local {
			case "v":
				inV = false
			case "t":
				inT = false
			case "f":
				inF = false
			}
			depth--
		case xml.CharData:
			if inV {
				vtext += string(e)
			}
			if inT && !inF {
				ttext += string(e)
			}
			if inF {
				ftext += string(e)
			}
		}
	}
	return
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
		return Value{Type: TypeString, S: ttext}, true
	case "s":
		i, err := strconv.Atoi(strings.TrimSpace(vtext))
		if err != nil || i < 0 || i >= len(sst) {
			return Value{Type: TypeString, S: vtext}, true
		}
		return Value{Type: TypeString, S: sst[i]}, true
	case "str":
		return Value{Type: TypeString, S: vtext}, true
	case "b":
		s := strings.TrimSpace(vtext)
		f, _ := strconv.ParseFloat(s, 64)
		return Value{Type: TypeBool, F: f, S: s}, true
	case "e":
		return Value{Type: TypeError, S: vtext}, true
	default:
		if vtext == "" && ttext == "" {
			return Value{}, false
		}
		if t == "n" || t == "" {
			if f, err := strconv.ParseFloat(strings.TrimSpace(vtext), 64); err == nil {
				return Value{Type: TypeNumber, F: f, S: strings.TrimSpace(vtext)}, true
			}
		}
		if ttext != "" {
			return Value{Type: TypeString, S: ttext}, true
		}
		return Value{Type: TypeString, S: vtext}, true
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
