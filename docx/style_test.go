package docx

import (
	"encoding/xml"
	"testing"

	"github.com/ieshan/go-ooxml/common"
	"github.com/ieshan/go-ooxml/docx/wml"
)

func TestDocument_Styles(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	// A new document may or may not have default styles
	// Just verify the iterator doesn't crash
	for s := range doc.Styles() {
		_ = s.ID()
		_ = s.Name()
	}
}

func TestDocument_StyleByID(t *testing.T) {
	doc := createDocWithStyles(t)
	defer doc.Close()
	s := doc.StyleByID("Heading1")
	if s == nil {
		t.Fatal("Heading1 style not found")
	}
	if s.Name() != "heading 1" {
		t.Errorf("name = %q", s.Name())
	}
	if s.Type() != "paragraph" {
		t.Errorf("type = %q", s.Type())
	}
}

func TestDocument_StyleByName(t *testing.T) {
	doc := createDocWithStyles(t)
	defer doc.Close()
	s := doc.StyleByName("heading 1")
	if s == nil {
		t.Fatal("style not found by name")
	}
	if s.ID() != "Heading1" {
		t.Errorf("id = %q", s.ID())
	}
}

func TestStyle_IsHeading(t *testing.T) {
	doc := createDocWithStyles(t)
	defer doc.Close()
	h1 := doc.StyleByID("Heading1")
	if h1 == nil || !h1.IsHeading() {
		t.Error("Heading1 should be heading")
	}
	if h1.HeadingLevel() != 1 {
		t.Errorf("level = %d", h1.HeadingLevel())
	}

	normal := doc.StyleByID("Normal")
	if normal != nil && normal.IsHeading() {
		t.Error("Normal should not be heading")
	}
}

func TestStyle_BasedOn(t *testing.T) {
	doc := createDocWithStyles(t)
	defer doc.Close()
	h1 := doc.StyleByID("Heading1")
	if h1 == nil {
		t.Fatal("nil")
	}
	if h1.BasedOn() != "Normal" {
		t.Errorf("basedOn = %q", h1.BasedOn())
	}
}

func TestDocument_StyleByID_NotFound(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	if doc.StyleByID("NonExistent") != nil {
		t.Error("should be nil")
	}
}

func ExampleDocument_Styles() {
	doc, _ := New(nil)
	defer doc.Close()
	// Iterate over all styles (requires a document with styles.xml).
	for s := range doc.Styles() {
		if s.IsHeading() {
			_ = s.HeadingLevel() // 1 for "Heading1", 2 for "Heading2", etc.
		}
	}
}

// Helper: create a document with styles injected
func createDocWithStyles(t *testing.T) *Document {
	t.Helper()
	doc, err := New(nil)
	if err != nil {
		t.Fatal(err)
	}

	// Inject styles directly into the document's styles field
	normalName := "Normal"
	h1Name := "heading 1"
	basedOn := "Normal"
	bold := true
	doc.styles = &wml.CT_Styles{
		XMLName: xml.Name{Space: wml.Ns, Local: "styles"},
		Styles: []*wml.CT_Style{
			{
				XMLName: xml.Name{Space: wml.Ns, Local: "style"},
				Type:    "paragraph", StyleID: "Normal", Name: normalName,
			},
			{
				XMLName: xml.Name{Space: wml.Ns, Local: "style"},
				Type:    "paragraph", StyleID: "Heading1", Name: h1Name,
				BasedOn: &basedOn,
				RPr:     &wml.CT_RPr{Bold: &bold},
			},
		},
	}
	return doc
}

func TestAddStyle(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()

	s := doc.AddStyle("MyCustom", "paragraph")
	if s == nil {
		t.Fatal("AddStyle returned nil")
	}
	if s.ID() != "MyCustom" {
		t.Errorf("ID = %q, want MyCustom", s.ID())
	}
	if s.Type() != "paragraph" {
		t.Errorf("Type = %q, want paragraph", s.Type())
	}

	found := doc.StyleByID("MyCustom")
	if found == nil {
		t.Error("AddStyle'd style not found by StyleByID")
	}
}

func TestStyleSetters(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()

	s := doc.AddStyle("Test", "paragraph")
	s.SetBasedOn("Normal")
	s.SetNext("Normal")
	s.SetFont("Georgia")
	s.SetFontSize(14)
	s.SetColor(common.RGB(0, 100, 0))
	s.SetBold(true)
	s.SetItalic(true)
	s.SetSpacingBefore(240)
	s.SetSpacingAfter(120)
	s.SetOutlineLevel(2)
	s.SetKeepNext(true)
	s.SetKeepLines(true)
	s.SetIndentLeft(720)
	s.SetIndentRight(360)
	s.SetAlignment("center")

	if s.BasedOn() != "Normal" {
		t.Errorf("BasedOn = %q", s.BasedOn())
	}
	if s.el.RPr == nil || s.el.RPr.FontName == nil || *s.el.RPr.FontName != "Georgia" {
		t.Error("font not set")
	}
	if s.el.RPr.Bold == nil || !*s.el.RPr.Bold {
		t.Error("bold not set")
	}
	if s.el.PPr == nil || s.el.PPr.OutlineLvl == nil || *s.el.PPr.OutlineLvl != 2 {
		t.Error("outlineLevel not set")
	}
	if s.el.PPr.Alignment == nil || *s.el.PPr.Alignment != "center" {
		t.Error("alignment not set")
	}
}

func TestEnsureBuiltinStyle(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()

	s := doc.EnsureBuiltinStyle("Heading1")
	if s == nil {
		t.Fatal("EnsureBuiltinStyle returned nil for Heading1")
	}
	if s.ID() != "Heading1" {
		t.Errorf("ID = %q", s.ID())
	}

	// Call again — should return same, no duplicate
	s2 := doc.EnsureBuiltinStyle("Heading1")
	if s2 == nil {
		t.Fatal("second call returned nil")
	}

	// Unknown style
	s3 := doc.EnsureBuiltinStyle("UnknownFoo")
	if s3 != nil {
		t.Error("unknown style should return nil")
	}
}

func ExampleDocument_AddStyle() {
	doc, _ := New(nil)
	defer doc.Close()

	// Create a custom paragraph style with specific formatting.
	s := doc.AddStyle("CustomNote", "paragraph")
	s.SetBasedOn("Normal")
	s.SetNext("Normal")
	s.SetFont("Georgia")
	s.SetFontSize(10)
	s.SetItalic(true)
	s.SetColor(common.RGB(100, 100, 100))
	s.SetSpacingBefore(120)
	s.SetSpacingAfter(120)
	s.SetIndentLeft(720)

	// Use the custom style on a paragraph.
	p := doc.Body().AddParagraph()
	p.SetStyle("CustomNote")
	p.AddRun("This is a styled note.")
}

func ExampleDocument_EnsureBuiltinStyle() {
	doc, _ := New(nil)
	defer doc.Close()

	// Ensure the Heading2 built-in style exists before using it.
	// If the document already contains the style, the existing definition is returned.
	// If not, the built-in definition is created automatically.
	doc.EnsureBuiltinStyle("Heading2")

	p := doc.Body().AddParagraph()
	p.SetStyle("Heading2")
	p.AddRun("Section Title")
}

func ExampleStyle_SetBold() {
	doc, _ := New(nil)
	defer doc.Close()

	s := doc.AddStyle("Strong", "paragraph")
	s.SetBold(true)
	s.SetFontSize(14)
	s.SetKeepNext(true)
	s.SetAlignment("center")
}

func ExampleStyle_SetLineSpacing() {
	doc, _ := New(nil)
	defer doc.Close()

	s := doc.AddStyle("DoubleSpaced", "paragraph")
	// 480 in 240ths of a line = double spacing with the "auto" rule.
	s.SetLineSpacing(480, "auto")
	s.SetSpacingBefore(240)
	s.SetSpacingAfter(240)
}
