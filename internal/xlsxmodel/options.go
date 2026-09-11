package xlsxmodel

import (
	"regexp"
	"strings"
)

// Named compare.ignore tokens from hierarchical config.json.
const (
	IgnoreFillBgColor         = "fillBgColor"
	IgnoreAnchorArraySpelling = "anchorArraySpelling"
	IgnoreVolatileValues      = "volatileValues"
	IgnoreCellFilenamePrefix  = "cellFilenamePrefix"
)

// Options are per-case compare exceptions. Parse stays option-free;
// DiffWorkbooks / Compare apply these.
type Options struct {
	Ignore []string            `json:"ignore,omitempty"`
	Cells  map[string]CellBand `json:"cells,omitempty"`
}

// CellBand is an inclusive numeric range for one cell (Sheet!A1).
// An in-range number is not a values FAIL; outside is.
type CellBand struct {
	Min *float64 `json:"min"`
	Max *float64 `json:"max"`
}

// KnownIgnore reports whether name is a supported compare.ignore token.
func KnownIgnore(name string) bool {
	switch name {
	case IgnoreFillBgColor, IgnoreAnchorArraySpelling, IgnoreVolatileValues, IgnoreCellFilenamePrefix:
		return true
	default:
		return false
	}
}

func (o Options) has(name string) bool {
	for _, s := range o.Ignore {
		if s == name {
			return true
		}
	}
	return false
}

var (
	anchorArrayRe  = regexp.MustCompile(`(?i)_xlfn\.ANCHORARRAY\(([^)]+)\)`)
	volatileFnRe   = regexp.MustCompile(`(?i)(?:^|[^A-Za-z.])(?:_xlfn\.)?(?:RANDARRAY|RAND|NOW|TODAY)\s*\(`)
	cellFilenameRe = regexp.MustCompile(`(?i)CELL\s*\(\s*"filename"`)
)

func normalizeAnchorArray(formula string) string {
	return anchorArrayRe.ReplaceAllString(formula, `$1#`)
}

func isVolatileFormula(formula string) bool {
	return formula != "" && volatileFnRe.MatchString(formula)
}

func isCellFilenameFormula(formula string) bool {
	return formula != "" && cellFilenameRe.MatchString(formula)
}

func filenameSuffix(s string) string {
	if i := strings.LastIndex(s, "["); i >= 0 {
		return s[i:]
	}
	return s
}

// unused Excel leftover on solid fills: indexed palette slot 64 vs omitted.
func normalizeFillBg(fill string) string {
	i := strings.LastIndex(fill, ",")
	if i < 0 {
		return fill
	}
	bg := fill[i+1:]
	if bg == "" || bg == "indexed:64" {
		return fill[:i] + ","
	}
	return fill
}

func volatileLocs(wb Workbook) map[string]bool {
	out := map[string]bool{}
	for sheet, cells := range wb.Cells {
		for ref, v := range cells {
			if !isVolatileFormula(v.Formula) {
				continue
			}
			out[sheet+"!"+ref] = true
			if v.FKind == "array" && v.FRef != "" {
				for _, r := range expandA1Range(v.FRef) {
					out[sheet+"!"+r] = true
				}
			}
		}
	}
	return out
}

func formulaCmpKey(v Value, opt Options) string {
	if opt.has(IgnoreAnchorArraySpelling) {
		v.Formula = normalizeAnchorArray(v.Formula)
	}
	return formulaKey(v)
}
