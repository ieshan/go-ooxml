package docx

import (
	"encoding/xml"
	"testing"

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
