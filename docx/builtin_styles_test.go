package docx

import (
	"bytes"
	"testing"

	"github.com/ieshan/go-ooxml/docx/wml"
)

func TestInitDefaultStyles(t *testing.T) {
	doc, err := New(nil)
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()

	if doc.styles == nil {
		t.Fatal("styles should not be nil after New()")
	}

	if len(doc.styles.Extra) == 0 {
		t.Error("expected docDefaults/latentStyles in Extra")
	}

	required := []string{"Normal", "Heading1", "Heading2", "Heading3", "Heading4", "Heading5", "Heading6",
		"Title", "Subtitle", "ListParagraph", "Quote", "DefaultParagraphFont", "TableNormal", "NoList"}
	for _, id := range required {
		found := false
		for _, s := range doc.styles.Styles {
			if s.StyleID == id {
				found = true
				break
			}
		}
		if !found {
			t.Errorf("missing required built-in style %q", id)
		}
	}
}

func TestBuiltinStyleProperties(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()

	s := doc.StyleByID("Heading1")
	if s == nil {
		t.Fatal("Heading1 not found")
	}
	if s.Name() != "heading 1" {
		t.Errorf("name = %q, want %q", s.Name(), "heading 1")
	}
	if s.BasedOn() != "Normal" {
		t.Errorf("basedOn = %q, want Normal", s.BasedOn())
	}
	if s.el.Next == nil || *s.el.Next != "Normal" {
		t.Error("next should be Normal")
	}
	if s.el.PPr == nil || s.el.PPr.OutlineLvl == nil || *s.el.PPr.OutlineLvl != 0 {
		t.Error("outlineLvl should be 0")
	}
	if s.el.RPr == nil || s.el.RPr.FontSize == nil || *s.el.RPr.FontSize != "32" {
		t.Errorf("fontSize should be 32 half-pts, got %v", s.el.RPr.FontSize)
	}
	if s.el.RPr.Color == nil || *s.el.RPr.Color != "2F5496" {
		t.Errorf("color should be 2F5496, got %v", s.el.RPr.Color)
	}
}

func TestEnsureStyleExists_Known(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()

	// Remove Heading1 to test auto-creation
	var filtered []*wml.CT_Style
	for _, s := range doc.styles.Styles {
		if s.StyleID != "Heading1" {
			filtered = append(filtered, s)
		}
	}
	doc.styles.Styles = filtered

	if doc.StyleByID("Heading1") != nil {
		t.Fatal("Heading1 should be removed")
	}

	doc.ensureStyleExists("Heading1")

	s := doc.StyleByID("Heading1")
	if s == nil {
		t.Fatal("Heading1 should be auto-created")
	}
	if s.Name() != "heading 1" {
		t.Errorf("auto-created Heading1 name = %q", s.Name())
	}
}

func TestEnsureStyleExists_Unknown(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()

	countBefore := len(doc.styles.Styles)
	doc.ensureStyleExists("CustomFoo")
	countAfter := len(doc.styles.Styles)

	if countAfter != countBefore {
		t.Error("unknown style should not be auto-created")
	}
}

func TestSetStyle_AutoEnsures(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()

	// Remove Heading3 to test that SetStyle re-creates it
	var filtered []*wml.CT_Style
	for _, s := range doc.styles.Styles {
		if s.StyleID != "Heading3" {
			filtered = append(filtered, s)
		}
	}
	doc.styles.Styles = filtered

	if doc.StyleByID("Heading3") != nil {
		t.Fatal("Heading3 should be removed")
	}

	p := doc.Body().AddParagraph()
	p.SetStyle("Heading3")

	if p.Style() != "Heading3" {
		t.Errorf("style = %q, want Heading3", p.Style())
	}

	s := doc.StyleByID("Heading3")
	if s == nil {
		t.Error("Heading3 should be auto-created by SetStyle")
	}
}

func TestSetStyle_CustomNoAutoCreate(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()

	countBefore := len(doc.styles.Styles)

	p := doc.Body().AddParagraph()
	p.SetStyle("MyCustomStyle")

	if p.Style() != "MyCustomStyle" {
		t.Errorf("style = %q, want MyCustomStyle", p.Style())
	}

	if len(doc.styles.Styles) != countBefore {
		t.Error("custom style should not be auto-created")
	}
}

func TestNewDoc_SaveHasStylesPart(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()

	doc.Body().AddParagraph().AddRun("test")

	var buf bytes.Buffer
	if err := doc.Write(&buf); err != nil {
		t.Fatalf("Write: %v", err)
	}

	doc2, err := OpenReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()), nil)
	if err != nil {
		t.Fatalf("OpenReader: %v", err)
	}
	defer doc2.Close()

	if doc2.styles == nil {
		t.Error("reopened doc should have styles")
	}

	s := doc2.StyleByID("Heading1")
	if s == nil {
		t.Error("reopened doc should have Heading1 style")
	}
}

func TestStyles_FullRoundTrip(t *testing.T) {
	doc, _ := New(nil)

	h1 := doc.Body().AddParagraph()
	h1.SetStyle("Heading1")
	h1.AddRun("Chapter One")

	h2 := doc.Body().AddParagraph()
	h2.SetStyle("Heading2")
	h2.AddRun("Section A")

	body := doc.Body().AddParagraph()
	body.SetStyle("Normal")
	body.AddRun("Body text here.")

	title := doc.Body().AddParagraph()
	title.SetStyle("Title")
	title.AddRun("Document Title")

	custom := doc.AddStyle("MyCustom", "paragraph")
	custom.SetBasedOn("Normal")
	custom.SetFont("Georgia")
	custom.SetFontSize(14)
	custom.SetBold(true)

	cp := doc.Body().AddParagraph()
	cp.SetStyle("MyCustom")
	cp.AddRun("Custom styled text")

	var buf bytes.Buffer
	if err := doc.Write(&buf); err != nil {
		t.Fatalf("Write: %v", err)
	}
	doc.Close()

	doc2, err := OpenReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()), nil)
	if err != nil {
		t.Fatalf("OpenReader: %v", err)
	}
	defer doc2.Close()

	for _, id := range []string{"Normal", "Heading1", "Heading2", "Title"} {
		s := doc2.StyleByID(id)
		if s == nil {
			t.Errorf("style %q missing after round-trip", id)
		}
	}

	cs := doc2.StyleByID("MyCustom")
	if cs == nil {
		t.Fatal("custom style missing after round-trip")
	}
	if cs.BasedOn() != "Normal" {
		t.Errorf("custom basedOn = %q", cs.BasedOn())
	}
	if cs.el.RPr == nil || cs.el.RPr.Bold == nil || !*cs.el.RPr.Bold {
		t.Error("custom bold not preserved")
	}

	paras := make([]struct{ text, style string }, 0)
	for p := range doc2.Body().Paragraphs() {
		paras = append(paras, struct{ text, style string }{p.Text(), p.Style()})
	}
	if len(paras) < 5 {
		t.Fatalf("expected >= 5 paragraphs, got %d", len(paras))
	}
	if paras[0].style != "Heading1" {
		t.Errorf("first para style = %q, want Heading1", paras[0].style)
	}
	if paras[3].style != "Title" {
		t.Errorf("fourth para style = %q, want Title", paras[3].style)
	}
	if paras[4].style != "MyCustom" {
		t.Errorf("fifth para style = %q, want MyCustom", paras[4].style)
	}
}

func TestOpenExisting_PreservesStyles(t *testing.T) {
	doc, _ := New(nil)
	doc.Body().AddParagraph().SetStyle("Heading1")
	doc.Body().AddParagraph().AddRun("text")

	var buf bytes.Buffer
	doc.Write(&buf)
	doc.Close()

	doc2, _ := OpenReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()), nil)
	defer doc2.Close()

	countBefore := 0
	for range doc2.Styles() {
		countBefore++
	}

	doc2.Body().AddParagraph().SetStyle("Heading1")

	countAfter := 0
	for range doc2.Styles() {
		countAfter++
	}

	if countAfter != countBefore {
		t.Errorf("style count changed: before=%d after=%d", countBefore, countAfter)
	}
}
