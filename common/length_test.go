package common

import (
	"testing"
)

func TestEMU(t *testing.T) {
	l := EMU(914400)
	if l.EMU() != 914400 {
		t.Errorf("EMU() = %d, want 914400", l.EMU())
	}
}

func TestInches(t *testing.T) {
	l := Inches(1.0)
	if l.EMU() != 914400 {
		t.Errorf("Inches(1.0).EMU() = %d, want 914400", l.EMU())
	}
	if got := l.Inches(); got != 1.0 {
		t.Errorf("Inches() = %f, want 1.0", got)
	}
}

func TestCm(t *testing.T) {
	l := Cm(2.54)
	if l.EMU() != 914400 {
		t.Errorf("Cm(2.54).EMU() = %d, want 914400", l.EMU())
	}
	if got := l.Cm(); got < 2.539 || got > 2.541 {
		t.Errorf("Cm() = %f, want ~2.54", got)
	}
}

func TestPt(t *testing.T) {
	l := Pt(72.0)
	if l.EMU() != 914400 {
		t.Errorf("Pt(72.0).EMU() = %d, want 914400", l.EMU())
	}
	if got := l.Pt(); got != 72.0 {
		t.Errorf("Pt() = %f, want 72.0", got)
	}
}

func TestTwips(t *testing.T) {
	l := Twips(1440)
	if l.EMU() != 914400 {
		t.Errorf("Twips(1440).EMU() = %d, want 914400", l.EMU())
	}
	if got := l.Twips(); got != 1440 {
		t.Errorf("Twips() = %d, want 1440", got)
	}
}

func TestLengthZero(t *testing.T) {
	var l Length
	if l.EMU() != 0 {
		t.Errorf("zero Length.EMU() = %d, want 0", l.EMU())
	}
	if l.Inches() != 0.0 {
		t.Errorf("zero Length.Inches() = %f, want 0.0", l.Inches())
	}
}

func ExampleInches() {
	l := Inches(1.5)
	_ = l.EMU()   // 1371600
	_ = l.Pt()    // 108.0
	_ = l.Twips() // 2160
}

func ExamplePt() {
	l := Pt(12.0)
	_ = l.EMU()    // 152400
	_ = l.Inches() // 0.1666...
}
