package xlsxmodel

import (
	"bytes"
	"encoding/xml"
	"math"
	"strconv"
	"strings"
)

// Style is resolved cell formatting (not style index, not theme display names).
type Style struct {
	NumFmt string
	Font   string
	Fill   string
	Border string
	Align  string
}

func (s Style) zero() bool {
	return s == Style{}
}

func (s Style) String() string {
	return strings.Join([]string{s.NumFmt, s.Font, s.Fill, s.Border, s.Align}, "|")
}

// ECMA-376 built-in numFmtId codes.
var builtinNumFmt = map[int]string{
	0: "General", 1: "0", 2: "0.00", 3: "#,##0", 4: "#,##0.00",
	9: "0%", 10: "0.00%", 11: "0.00E+00", 12: "# ?/?", 13: "# ??/??",
	14: "mm-dd-yy", 15: "d-mmm-yy", 16: "d-mmm", 17: "mmm-yy",
	18: "h:mm AM/PM", 19: "h:mm:ss AM/PM", 20: "h:mm", 21: "h:mm:ss", 22: "m/d/yy h:mm",
	37: "#,##0 ;(#,##0)", 38: "#,##0 ;[Red](#,##0)", 39: "#,##0.00;(#,##0.00)", 40: "#,##0.00;[Red](#,##0.00)",
	45: "mm:ss", 46: "[h]:mm:ss", 47: "mmss.0", 48: "##0.0E+0", 49: "@",
}

// SpreadsheetML color@theme index → DrawingML clrScheme slot (lt/dk swapped).
var themeIndexSlot = []string{
	"lt1", "dk1", "lt2", "dk2",
	"accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
	"hlink", "folHlink",
}

type styleBook struct {
	numFmts map[int]string
	fonts   []string
	fills   []string
	borders []string
	xfs     []rawXf
	theme   map[string]string // slot → RGB
}

type rawXf struct {
	numFmtId, fontId, fillId, borderId int
	align                              string
}

func parseStyleBook(stylesXML, themeXML []byte) *styleBook {
	sb := &styleBook{numFmts: map[int]string{}, theme: parseTheme(themeXML)}
	if len(stylesXML) == 0 {
		return sb
	}
	dec := xml.NewDecoder(bytes.NewReader(stylesXML))
	section := ""
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
		case "numFmts", "fonts", "fills", "borders", "cellXfs", "dxfs":
			section = se.Name.Local
			if se.Name.Local == "dxfs" {
				skip(dec)
				section = ""
			}
		case "numFmt":
			if section == "numFmts" {
				id, _ := strconv.Atoi(attr(se, "numFmtId"))
				sb.numFmts[id] = attr(se, "formatCode")
			}
		case "font":
			if section == "fonts" {
				sb.fonts = append(sb.fonts, readFont(dec, sb.theme))
			} else {
				skip(dec)
			}
		case "fill":
			if section == "fills" {
				sb.fills = append(sb.fills, readFill(dec, sb.theme))
			} else {
				skip(dec)
			}
		case "border":
			if section == "borders" {
				sb.borders = append(sb.borders, readBorder(dec, sb.theme))
			} else {
				skip(dec)
			}
		case "xf":
			if section == "cellXfs" {
				sb.xfs = append(sb.xfs, readXf(dec, se))
			} else {
				skip(dec)
			}
		}
	}
	return sb
}

func (sb *styleBook) resolve(sIndex string) Style {
	if len(sb.xfs) == 0 {
		return Style{}
	}
	i := 0
	if sIndex != "" {
		i, _ = strconv.Atoi(sIndex)
	}
	if i < 0 || i >= len(sb.xfs) {
		i = 0
	}
	xf := sb.xfs[i]
	return Style{
		NumFmt: sb.numFmt(xf.numFmtId),
		Font:   at(sb.fonts, xf.fontId),
		Fill:   at(sb.fills, xf.fillId),
		Border: at(sb.borders, xf.borderId),
		Align:  xf.align,
	}
}

func (sb *styleBook) numFmt(id int) string {
	if s, ok := sb.numFmts[id]; ok {
		return s
	}
	if s, ok := builtinNumFmt[id]; ok {
		return s
	}
	return strconv.Itoa(id)
}

func at(ss []string, i int) string {
	if i >= 0 && i < len(ss) {
		return ss[i]
	}
	return ""
}

func readXf(dec *xml.Decoder, se xml.StartElement) rawXf {
	xf := rawXf{
		numFmtId: atoi(attr(se, "numFmtId")),
		fontId:   atoi(attr(se, "fontId")),
		fillId:   atoi(attr(se, "fillId")),
		borderId: atoi(attr(se, "borderId")),
	}
	depth := 1
	for depth > 0 {
		tok, err := dec.Token()
		if err != nil {
			break
		}
		switch e := tok.(type) {
		case xml.StartElement:
			depth++
			if e.Name.Local == "alignment" {
				xf.align = strings.Join([]string{
					attr(e, "horizontal"), attr(e, "vertical"), attr(e, "wrapText"),
				}, ",")
			}
		case xml.EndElement:
			depth--
		}
	}
	return xf
}

func readFont(dec *xml.Decoder, theme map[string]string) string {
	var name, sz, u, color string
	bold, italic := false, false
	walkLeaves(dec, func(se xml.StartElement) {
		switch se.Name.Local {
		case "name":
			name = attr(se, "val")
		case "sz":
			sz = attr(se, "val")
		case "b":
			bold = attr(se, "val") != "0"
		case "i":
			italic = attr(se, "val") != "0"
		case "u":
			u = attr(se, "val")
			if u == "" {
				u = "single"
			}
		case "color":
			color = resolveColor(se, theme)
		}
	})
	return strings.Join([]string{name, sz, bool01(bold), bool01(italic), u, color}, ",")
}

func readFill(dec *xml.Decoder, theme map[string]string) string {
	var pat, fg, bg string
	walkLeaves(dec, func(se xml.StartElement) {
		switch se.Name.Local {
		case "patternFill":
			pat = attr(se, "patternType")
		case "fgColor":
			fg = resolveColor(se, theme)
		case "bgColor":
			bg = resolveColor(se, theme)
		}
	})
	return strings.Join([]string{pat, fg, bg}, ",")
}

func readBorder(dec *xml.Decoder, theme map[string]string) string {
	var parts []string
	sideName, style, color := "", "", ""
	walkDepth(dec, func(se xml.StartElement) {
		switch se.Name.Local {
		case "left", "right", "top", "bottom", "diagonal":
			sideName = se.Name.Local
			style = attr(se, "style")
			color = ""
		case "color":
			if sideName != "" {
				color = resolveColor(se, theme)
			}
		}
	}, func(local string) {
		if local == "left" || local == "right" || local == "top" || local == "bottom" || local == "diagonal" {
			// Empty <left/> (Excel default) is the same as an omitted side.
			if style != "" {
				part := sideName + ":" + style
				if color != "" {
					part += ":" + color
				}
				parts = append(parts, part)
			}
			sideName, style, color = "", "", ""
		}
	})
	return strings.Join(parts, ";")
}

func walkLeaves(dec *xml.Decoder, start func(xml.StartElement)) {
	walkDepth(dec, start, nil)
}

func walkDepth(dec *xml.Decoder, start func(xml.StartElement), end func(string)) {
	depth := 1
	for depth > 0 {
		tok, err := dec.Token()
		if err != nil {
			return
		}
		switch e := tok.(type) {
		case xml.StartElement:
			depth++
			if start != nil {
				start(e)
			}
		case xml.EndElement:
			if end != nil {
				end(e.Name.Local)
			}
			depth--
		}
	}
}

func resolveColor(se xml.StartElement, theme map[string]string) string {
	if rgb := attr(se, "rgb"); rgb != "" {
		return strings.ToUpper(rgb)
	}
	if th := attr(se, "theme"); th != "" {
		i, _ := strconv.Atoi(th)
		slot := ""
		if i >= 0 && i < len(themeIndexSlot) {
			slot = themeIndexSlot[i]
		}
		rgb := theme[slot]
		if rgb == "" {
			rgb = "theme:" + th
		}
		if tint := attr(se, "tint"); tint != "" {
			return rgb + "@" + normalizeTint(tint)
		}
		return rgb
	}
	if ix := attr(se, "indexed"); ix != "" {
		return "indexed:" + ix
	}
	return ""
}

func parseTheme(raw []byte) map[string]string {
	out := map[string]string{}
	if len(raw) == 0 {
		return out
	}
	dec := xml.NewDecoder(bytes.NewReader(raw))
	slot := ""
	inScheme := false
	for {
		tok, err := dec.Token()
		if err != nil {
			break
		}
		switch e := tok.(type) {
		case xml.StartElement:
			switch e.Name.Local {
			case "clrScheme":
				inScheme = true
			case "dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink":
				if inScheme {
					slot = e.Name.Local
				}
			case "srgbClr":
				if slot != "" {
					out[slot] = strings.ToUpper(attr(e, "val"))
				}
			case "sysClr":
				if slot != "" {
					if last := attr(e, "lastClr"); last != "" {
						out[slot] = strings.ToUpper(last)
					}
				}
			}
		case xml.EndElement:
			switch e.Name.Local {
			case "clrScheme":
				inScheme = false
			case "dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink":
				slot = ""
			}
		}
	}
	return out
}

func normalizeTint(s string) string {
	f, err := strconv.ParseFloat(s, 64)
	if err != nil {
		return s
	}
	// Excel ST_Percentage often serializes 0.2 as 0.19998779259620961.
	return strconv.FormatFloat(math.Round(f*1e4)/1e4, 'f', 4, 64)
}

func atoi(s string) int {
	n, _ := strconv.Atoi(s)
	return n
}

func bool01(b bool) string {
	if b {
		return "1"
	}
	return "0"
}
