// Package compare semantically diffs two xlsx workbooks.
//
// Inspired by mog's xlsx-roundtrip archive/XML compare: ZIP contents, not
// byte-identity. Volatile Excel package bits are ignored; date1904, sheet
// order, values, types, formulas, styles, merges, names, and freeze are not.
package compare

import (
	"archive/zip"
	"bytes"
	"fmt"
	"io"
	"os"
	"regexp"
	"strings"
)

// Diff is one semantic difference between a mog export and a golden.
type Diff struct {
	Part   string
	Detail string
}

// Result is the outcome of comparing two xlsx archives.
type Result struct {
	Equal bool
	Diffs []Diff
}

// Files compares a mog export xlsx to a golden xlsx.
func Files(exportPath, goldenPath string) (Result, error) {
	export, err := os.ReadFile(exportPath)
	if err != nil {
		return Result{}, fmt.Errorf("export: %w", err)
	}
	golden, err := os.ReadFile(goldenPath)
	if err != nil {
		return Result{}, fmt.Errorf("golden: %w", err)
	}
	return Archives(export, golden)
}

// Archives compares two xlsx blobs (export, golden). ZIP mtimes are ignored
// because only entry contents are read.
func Archives(export, golden []byte) (Result, error) {
	expParts, err := readParts(export)
	if err != nil {
		return Result{}, fmt.Errorf("export xlsx: %w", err)
	}
	goldParts, err := readParts(golden)
	if err != nil {
		return Result{}, fmt.Errorf("golden xlsx: %w", err)
	}

	var diffs []Diff
	seen := make(map[string]bool, len(expParts)+len(goldParts))
	for name := range goldParts {
		seen[name] = true
		if _, ok := expParts[name]; !ok {
			diffs = append(diffs, Diff{Part: name, Detail: "present in golden, missing in export"})
			continue
		}
		if d := partDiff(name, expParts[name], goldParts[name]); d != nil {
			diffs = append(diffs, *d)
		}
	}
	for name := range expParts {
		if seen[name] {
			continue
		}
		diffs = append(diffs, Diff{Part: name, Detail: "not in golden, extra in export"})
	}
	return Result{Equal: len(diffs) == 0, Diffs: diffs}, nil
}

func readParts(data []byte) (map[string][]byte, error) {
	r, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		return nil, err
	}
	out := make(map[string][]byte, len(r.File))
	for _, f := range r.File {
		name := strings.ReplaceAll(f.Name, "\\", "/")
		if strings.HasSuffix(name, "/") || ignoredPart(name) {
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

func ignoredPart(name string) bool {
	n := strings.ToLower(name)
	if n == "xl/calcchain.xml" {
		return true
	}
	return strings.Contains(n, "printersettings")
}

func partDiff(name string, export, golden []byte) *Diff {
	if bytes.Equal(export, golden) {
		return nil
	}
	if isXML(name) {
		a := normalizePart(name, export)
		b := normalizePart(name, golden)
		if a == b {
			return nil
		}
		return &Diff{Part: name, Detail: "xml content differs"}
	}
	return &Diff{Part: name, Detail: fmt.Sprintf("binary content differs (%d vs %d bytes)", len(export), len(golden))}
}

func isXML(name string) bool {
	return strings.HasSuffix(name, ".xml") || strings.HasSuffix(name, ".rels")
}

func normalizePart(name string, raw []byte) string {
	s := string(raw)
	s = stripVolatile(name, s)
	return normalizeXML(s)
}

func stripVolatile(name, xml string) string {
	switch {
	case name == "xl/workbook.xml":
		xml = calcIDRe.ReplaceAllString(xml, "")
	case name == "docProps/core.xml":
		for _, tag := range []string{"dc:creator", "cp:lastModifiedBy", "dcterms:created", "dcterms:modified"} {
			xml = stripElem(xml, tag)
		}
	case name == "docProps/app.xml":
		xml = stripElem(xml, "Application")
		xml = stripElem(xml, "AppVersion")
	case name == "[Content_Types].xml":
		xml = stripContentTypeOverrides(xml)
	case strings.HasSuffix(name, ".rels"):
		xml = stripVolatileRels(xml)
	}
	return xml
}

var calcIDRe = regexp.MustCompile(`\s+calcId="[^"]*"`)

func stripElem(xml, tag string) string {
	pair := regexp.MustCompile(`(?s)<` + regexp.QuoteMeta(tag) + `\b[^>]*>.*?</` + regexp.QuoteMeta(tag) + `>`)
	xml = pair.ReplaceAllString(xml, "")
	self := regexp.MustCompile(`<` + regexp.QuoteMeta(tag) + `\b[^>]*/>`)
	return self.ReplaceAllString(xml, "")
}

func stripContentTypeOverrides(xml string) string {
	re := regexp.MustCompile(`<Override\b[^>]*/>`)
	return re.ReplaceAllStringFunc(xml, func(tag string) string {
		low := strings.ToLower(tag)
		if strings.Contains(low, "calcchain.xml") || strings.Contains(low, "printersettings") {
			return ""
		}
		return tag
	})
}

func stripVolatileRels(xml string) string {
	re := regexp.MustCompile(`<Relationship\b[^>]*/>`)
	return re.ReplaceAllStringFunc(xml, func(tag string) string {
		low := strings.ToLower(tag)
		if strings.Contains(low, "calcchain") || strings.Contains(low, "printersettings") {
			return ""
		}
		return tag
	})
}

var betweenTags = regexp.MustCompile(`>\s+<`)

func normalizeXML(s string) string {
	s = strings.ReplaceAll(s, "\r\n", "\n")
	s = strings.ReplaceAll(s, "\r", "\n")
	s = betweenTags.ReplaceAllString(s, "><")
	return strings.TrimSpace(s)
}
