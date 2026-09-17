package xlsxmodel

import (
	"fmt"
	"math"
	"strconv"
	"strings"
)

// Canonicalize the displayed color. Theme colors use RGB, while explicit
// spreadsheet colors normally use opaque ARGB. Preserve non-opaque alpha.
func canonicalColor(color, tint string) string {
	color = strings.ToUpper(color)
	if len(color) == 8 && strings.HasPrefix(color, "FF") {
		color = color[2:]
	}
	if tint == "" {
		return color
	}
	t, err := strconv.ParseFloat(normalizeTint(tint), 64)
	if err != nil || math.IsNaN(t) || math.IsInf(t, 0) || t < -1 || t > 1 {
		return color + "@" + tint
	}
	if t == 0 {
		return color
	}
	alpha, rgb := "", color
	if len(rgb) == 8 {
		alpha, rgb = rgb[:2], rgb[2:]
	}
	v, err := strconv.ParseUint(rgb, 16, 24)
	if len(rgb) != 6 || err != nil {
		return color + "@" + normalizeTint(tint)
	}
	r, g, b := int(v>>16), int(v>>8&255), int(v&255)
	h, l, s := rgbToHLS(r, g, b)
	if t < 0 {
		l = int(math.Floor(float64(l) * (1 + t)))
	} else {
		l = int(math.Floor(float64(l)*(1-t) + 240*t))
	}
	r, g, b = hlsToRGB(h, l, s)
	return fmt.Sprintf("%s%02X%02X%02X", alpha, r, g, b)
}

// Excel applies tint in integer HLS space, with HLSMAX=240 and RGBMAX=255.
func rgbToHLS(r, g, b int) (h, l, s int) {
	hi, lo := max(r, g, b), min(r, g, b)
	l = ((hi+lo)*240 + 255) / 510
	if hi == lo {
		return 0, l, 0
	}
	if l <= 120 {
		s = ((hi-lo)*240 + (hi+lo)/2) / (hi + lo)
	} else {
		s = ((hi-lo)*240 + (510-hi-lo)/2) / (510 - hi - lo)
	}
	delta := func(v int) int { return ((hi-v)*40 + (hi-lo)/2) / (hi - lo) }
	switch hi {
	case r:
		h = delta(b) - delta(g)
	case g:
		h = 80 + delta(r) - delta(b)
	default:
		h = 160 + delta(g) - delta(r)
	}
	if h < 0 {
		h += 240
	}
	if h > 240 {
		h -= 240
	}
	return
}

func hlsToRGB(h, l, s int) (int, int, int) {
	if s == 0 {
		v := (l*255 + 120) / 240
		return v, v, v
	}
	m2 := l + s - (l*s+120)/240
	if l <= 120 {
		m2 = (l*(240+s) + 120) / 240
	}
	m1 := 2*l - m2
	channel := func(hue int) int {
		if hue < 0 {
			hue += 240
		}
		if hue > 240 {
			hue -= 240
		}
		v := m1
		switch {
		case hue < 40:
			v = m1 + ((m2-m1)*hue+20)/40
		case hue < 120:
			v = m2
		case hue < 160:
			v = m1 + ((m2-m1)*(160-hue)+20)/40
		}
		return max(0, min(255, (v*255+120)/240))
	}
	return channel(h + 80), channel(h), channel(h - 80)
}
