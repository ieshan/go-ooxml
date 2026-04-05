package common

import "testing"

func TestRGB(t *testing.T) {
	c := RGB(0x3C, 0x2F, 0x80)
	if c.R != 0x3C || c.G != 0x2F || c.B != 0x80 {
		t.Errorf("RGB() = {%d, %d, %d}, want {60, 47, 128}", c.R, c.G, c.B)
	}
}

func TestColorHex(t *testing.T) {
	c := RGB(0x3C, 0x2F, 0x80)
	if got := c.Hex(); got != "3C2F80" {
		t.Errorf("Hex() = %q, want %q", got, "3C2F80")
	}
}

func TestColorHexBlack(t *testing.T) {
	c := RGB(0, 0, 0)
	if got := c.Hex(); got != "000000" {
		t.Errorf("Hex() = %q, want %q", got, "000000")
	}
}

func TestHexColor(t *testing.T) {
	tests := []struct {
		hex     string
		wantR   uint8
		wantG   uint8
		wantB   uint8
		wantErr bool
	}{
		{"FF0000", 255, 0, 0, false},
		{"00ff00", 0, 255, 0, false},
		{"0000FF", 0, 0, 255, false},
		{"3C2F80", 0x3C, 0x2F, 0x80, false},
		{"ZZZZZZ", 0, 0, 0, true},
		{"FF00", 0, 0, 0, true},
		{"FF00FF00", 0, 0, 0, true},
		{"", 0, 0, 0, true},
	}
	for _, tt := range tests {
		c, err := HexColor(tt.hex)
		if tt.wantErr {
			if err == nil {
				t.Errorf("HexColor(%q) should error", tt.hex)
			}
			continue
		}
		if err != nil {
			t.Errorf("HexColor(%q) unexpected error: %v", tt.hex, err)
			continue
		}
		if c.R != tt.wantR || c.G != tt.wantG || c.B != tt.wantB {
			t.Errorf("HexColor(%q) = {%d, %d, %d}, want {%d, %d, %d}",
				tt.hex, c.R, c.G, c.B, tt.wantR, tt.wantG, tt.wantB)
		}
	}
}

func ExampleRGB() {
	c := RGB(255, 0, 0)
	_ = c.Hex() // "FF0000"
}

func ExampleHexColor() {
	c, _ := HexColor("3C2F80")
	_ = c.R // 60
	_ = c.G // 47
	_ = c.B // 128
}
