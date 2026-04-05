package docx

import (
	"testing"

	"github.com/ieshan/go-ooxml/common"
)

func TestDocument_Sections(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	count := 0
	for range doc.Sections() {
		count++
	}
	// New document should have at least the default section
	if count < 1 {
		t.Errorf("sections = %d, want >= 1", count)
	}
}

func TestSection_PageSize(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	for s := range doc.Sections() {
		s.SetPageSize(common.Inches(8.5), common.Inches(11))
		w, h := s.PageSize()
		if w != common.Inches(8.5) {
			t.Errorf("width = %d, want %d", w, common.Inches(8.5))
		}
		if h != common.Inches(11) {
			t.Errorf("height = %d, want %d", h, common.Inches(11))
		}
	}
}

func TestSection_Orientation(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	for s := range doc.Sections() {
		s.SetOrientation(OrientLandscape)
		if s.Orientation() != OrientLandscape {
			t.Error("expected landscape orientation")
		}
		s.SetOrientation(OrientPortrait)
		if s.Orientation() != OrientPortrait {
			t.Error("expected portrait orientation")
		}
	}
}

func TestSection_Margins(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	for s := range doc.Sections() {
		m := SectionMargins{
			Top:    common.Inches(1),
			Right:  common.Inches(1),
			Bottom: common.Inches(1),
			Left:   common.Inches(1),
		}
		s.SetMargins(m)
		got := s.Margins()
		if got.Top != common.Inches(1) {
			t.Errorf("Top margin = %d, want %d", got.Top, common.Inches(1))
		}
		if got.Right != common.Inches(1) {
			t.Errorf("Right margin = %d, want %d", got.Right, common.Inches(1))
		}
		if got.Bottom != common.Inches(1) {
			t.Errorf("Bottom margin = %d, want %d", got.Bottom, common.Inches(1))
		}
		if got.Left != common.Inches(1) {
			t.Errorf("Left margin = %d, want %d", got.Left, common.Inches(1))
		}
	}
}

func TestSection_Columns(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	for s := range doc.Sections() {
		s.SetColumns(&ColumnLayout{Count: 2, Space: common.Inches(0.5), EqualWidth: true})
		cols := s.Columns()
		if cols == nil {
			t.Fatal("expected non-nil columns")
		}
		if cols.Count != 2 {
			t.Errorf("columns count = %d, want 2", cols.Count)
		}
		if cols.Space != common.Inches(0.5) {
			t.Errorf("columns space = %d, want %d", cols.Space, common.Inches(0.5))
		}
		if !cols.EqualWidth {
			t.Error("expected EqualWidth = true")
		}
	}
}

func TestSection_Columns_Nil(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	for s := range doc.Sections() {
		// Setting then clearing columns.
		s.SetColumns(&ColumnLayout{Count: 3})
		s.SetColumns(nil)
		if cols := s.Columns(); cols != nil {
			t.Errorf("expected nil columns after clearing, got %+v", cols)
		}
	}
}

func TestSection_BreakType(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	for s := range doc.Sections() {
		breaks := []SectionBreak{BreakNextPage, BreakContinuous, BreakEvenPage, BreakOddPage}
		for _, b := range breaks {
			s.SetBreakType(b)
			if got := s.BreakType(); got != b {
				t.Errorf("BreakType() = %d, want %d", got, b)
			}
		}
	}
}

func TestSection_Text_Markdown_Stubs(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	for s := range doc.Sections() {
		if s.Text() != "" {
			t.Errorf("Text() = %q, want empty", s.Text())
		}
		if s.Markdown() != "" {
			t.Errorf("Markdown() = %q, want empty", s.Markdown())
		}
	}
}

func ExampleSection_SetColumns() {
	doc, _ := New(nil)
	defer doc.Close()
	for s := range doc.Sections() {
		s.SetColumns(&ColumnLayout{Count: 2, Space: common.Inches(0.5), EqualWidth: true})
	}
}
