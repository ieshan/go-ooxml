package docx

import (
	"testing"

	"github.com/ieshan/go-ooxml/docx/wml"
)

func TestParagraph_ListInfo_NotList(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	p := doc.Body().AddParagraph()
	p.AddRun("normal")
	if p.ListInfo() != nil {
		t.Error("should be nil for non-list")
	}
	if p.IsListItem() {
		t.Error("should not be list item")
	}
}

func TestParagraph_ListInfo_WithNumPr(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	p := doc.Body().AddParagraph()
	p.AddRun("list item")

	// Set up NumPr via internal field.
	p.el.PPr = &wml.CT_PPr{
		NumPr: &wml.CT_NumPr{
			ILvl:  new(0),
			NumID: new(3),
		},
	}

	li := p.ListInfo()
	if li == nil {
		t.Fatal("expected non-nil ListInfo")
	}
	if li.Level != 0 {
		t.Errorf("level = %d, want 0", li.Level)
	}
	if li.NumID != 3 {
		t.Errorf("numID = %d, want 3", li.NumID)
	}
	if li.Format != ListBullet {
		t.Errorf("format = %v, want ListBullet", li.Format)
	}
	if !p.IsListItem() {
		t.Error("expected IsListItem() = true")
	}
}

func TestParagraph_ListInfo_Level2(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	p := doc.Body().AddParagraph()

	p.el.PPr = &wml.CT_PPr{
		NumPr: &wml.CT_NumPr{
			ILvl:  new(2),
			NumID: new(1),
		},
	}

	li := p.ListInfo()
	if li == nil {
		t.Fatal("expected non-nil ListInfo")
	}
	if li.Level != 2 {
		t.Errorf("level = %d, want 2", li.Level)
	}
}

func ExampleParagraph_IsListItem() {
	doc, _ := New(nil)
	defer doc.Close()
	doc.Body().AddParagraph().AddRun("Normal paragraph")
	for p := range doc.Body().Paragraphs() {
		if p.IsListItem() {
			li := p.ListInfo()
			_ = li.Level // zero-based indentation level
		}
	}
}

func TestParagraph_ListInfo_NilILvlDefaultsToZero(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	p := doc.Body().AddParagraph()

	p.el.PPr = &wml.CT_PPr{
		NumPr: &wml.CT_NumPr{
			NumID: new(1),
			// ILvl intentionally nil
		},
	}

	li := p.ListInfo()
	if li == nil {
		t.Fatal("expected non-nil ListInfo")
	}
	if li.Level != 0 {
		t.Errorf("level = %d, want 0 when ILvl is nil", li.Level)
	}
}
