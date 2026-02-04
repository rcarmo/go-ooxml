package utils

import (
	"fmt"
	"regexp"
	"strconv"
	"strings"
)

// Color represents an ARGB color.
type Color struct {
	A, R, G, B uint8
}

var hexColorRegex = regexp.MustCompile(`^#?([0-9A-Fa-f]{6}|[0-9A-Fa-f]{8})$`)

// ParseHexColor parses a hex color string (with or without #).
// Supports RGB (#RRGGBB) and ARGB (#AARRGGBB) formats.
func ParseHexColor(hex string) (Color, error) {
	hex = strings.TrimPrefix(hex, "#")
	hex = strings.ToUpper(hex)

	if !hexColorRegex.MatchString(hex) {
		return Color{}, fmt.Errorf("invalid hex color: %s", hex)
	}

	var c Color
	if len(hex) == 6 {
		c.A = 255
		val, _ := strconv.ParseUint(hex, 16, 32)
		c.R = uint8(val >> 16)
		c.G = uint8(val >> 8)
		c.B = uint8(val)
	} else {
		val, _ := strconv.ParseUint(hex, 16, 32)
		c.A = uint8(val >> 24)
		c.R = uint8(val >> 16)
		c.G = uint8(val >> 8)
		c.B = uint8(val)
	}
	return c, nil
}

// ToHex returns the color as a hex string without # (RRGGBB or AARRGGBB).
func (c Color) ToHex() string {
	if c.A == 255 {
		return fmt.Sprintf("%02X%02X%02X", c.R, c.G, c.B)
	}
	return fmt.Sprintf("%02X%02X%02X%02X", c.A, c.R, c.G, c.B)
}

// ToHexWithHash returns the color as a hex string with # prefix.
func (c Color) ToHexWithHash() string {
	return "#" + c.ToHex()
}

// ToARGB returns the color as an ARGB hex string (always 8 chars).
func (c Color) ToARGB() string {
	return fmt.Sprintf("%02X%02X%02X%02X", c.A, c.R, c.G, c.B)
}

// Standard colors.
var (
	ColorBlack   = Color{255, 0, 0, 0}
	ColorWhite   = Color{255, 255, 255, 255}
	ColorRed     = Color{255, 255, 0, 0}
	ColorGreen   = Color{255, 0, 255, 0}
	ColorBlue    = Color{255, 0, 0, 255}
	ColorYellow  = Color{255, 255, 255, 0}
	ColorCyan    = Color{255, 0, 255, 255}
	ColorMagenta = Color{255, 255, 0, 255}
)

// ThemeColor represents a theme color reference.
type ThemeColor struct {
	Theme int     // Theme color index (0-11)
	Tint  float64 // Tint value (-1.0 to 1.0)
}

// Common theme color indices.
const (
	ThemeColorLight1      = 0  // Usually background
	ThemeColorDark1       = 1  // Usually text
	ThemeColorLight2      = 2  // Background variant
	ThemeColorDark2       = 3  // Text variant
	ThemeColorAccent1     = 4
	ThemeColorAccent2     = 5
	ThemeColorAccent3     = 6
	ThemeColorAccent4     = 7
	ThemeColorAccent5     = 8
	ThemeColorAccent6     = 9
	ThemeColorHyperlink   = 10
	ThemeColorFollowedHyp = 11
)

// HighlightColors maps Word highlight names to RGB colors.
var HighlightColors = map[string]Color{
	"yellow":      {255, 255, 255, 0},
	"green":       {255, 0, 255, 0},
	"cyan":        {255, 0, 255, 255},
	"magenta":     {255, 255, 0, 255},
	"blue":        {255, 0, 0, 255},
	"red":         {255, 255, 0, 0},
	"darkBlue":    {255, 0, 0, 139},
	"darkCyan":    {255, 0, 139, 139},
	"darkGreen":   {255, 0, 100, 0},
	"darkMagenta": {255, 139, 0, 139},
	"darkRed":     {255, 139, 0, 0},
	"darkYellow":  {255, 128, 128, 0},
	"darkGray":    {255, 169, 169, 169},
	"lightGray":   {255, 211, 211, 211},
	"black":       {255, 0, 0, 0},
}
