package xlsxmodel

import (
	"regexp"
	"strconv"
	"strings"
)

var a1RefRe = regexp.MustCompile(`(\$?)([A-Za-z]{1,3})(\$?)([1-9][0-9]{0,6})`)

func shiftFormula(formula string, dCol, dRow int) string {
	if dCol == 0 && dRow == 0 {
		return formula
	}
	return a1RefRe.ReplaceAllStringFunc(formula, func(tok string) string {
		m := a1RefRe.FindStringSubmatch(tok)
		colAbs, colName, rowAbs := m[1] == "$", m[2], m[3] == "$"
		row, _ := strconv.Atoi(m[4])
		col := parseCol(colName)
		if !colAbs {
			col += dCol
		}
		if !rowAbs {
			row += dRow
		}
		if col < 1 || row < 1 {
			return tok
		}
		out := ""
		if colAbs {
			out += "$"
		}
		out += formatCol(col)
		if rowAbs {
			out += "$"
		}
		return out + strconv.Itoa(row)
	})
}

func parseA1(ref string) (col, row int, ok bool) {
	m := a1RefRe.FindStringSubmatch(ref)
	if m == nil {
		return 0, 0, false
	}
	row, err := strconv.Atoi(m[4])
	if err != nil {
		return 0, 0, false
	}
	return parseCol(m[2]), row, true
}

func parseCol(s string) int {
	n := 0
	for _, c := range strings.ToUpper(s) {
		if c < 'A' || c > 'Z' {
			return 0
		}
		n = n*26 + int(c-'A'+1)
	}
	return n
}

func formatCol(n int) string {
	var b []byte
	for n > 0 {
		n--
		b = append([]byte{byte('A' + n%26)}, b...)
		n /= 26
	}
	return string(b)
}

func deltaA1(from, to string) (dCol, dRow int) {
	fc, fr, ok1 := parseA1(from)
	tc, tr, ok2 := parseA1(to)
	if !ok1 || !ok2 {
		return 0, 0
	}
	return tc - fc, tr - fr
}
