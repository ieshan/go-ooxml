package common

import (
	"encoding/hex"
	"fmt"
)

// Color represents an RGB color value. Each component ranges from 0 to 255.
//
// Create a Color using [RGB] for component values or [HexColor] for a
// 6-character hex string (e.g., "FF0000" for red).
//
// Color is a value type and is safe for concurrent use.
type Color struct {
	R, G, B uint8
}

// RGB creates a [Color] from red, green, and blue components (0–255 each).
func RGB(r, g, b uint8) Color {
	return Color{R: r, G: g, B: b}
}

// HexColor parses a 6-character hexadecimal color string (e.g., "FF0000")
// into a [Color]. The string must be exactly 6 hex characters without a
// leading "#". Both uppercase and lowercase hex digits are accepted.
func HexColor(s string) (Color, error) {
	if len(s) != 6 {
		return Color{}, fmt.Errorf("ooxml: hex color must be 6 characters, got %d", len(s))
	}
	b, err := hex.DecodeString(s)
	if err != nil {
		return Color{}, fmt.Errorf("ooxml: invalid hex color %q: %w", s, err)
	}
	return Color{R: b[0], G: b[1], B: b[2]}, nil
}

// Hex returns the color as a 6-character uppercase hexadecimal string
// (e.g., "FF0000" for red).
func (c Color) Hex() string {
	return fmt.Sprintf("%02X%02X%02X", c.R, c.G, c.B)
}
