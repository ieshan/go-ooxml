package docx

import (
	"archive/zip"
	"bytes"
	"io"
	"strings"
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

func TestStyles_XMLNamespaceIntegrity(t *testing.T) {
	doc, _ := New(nil)
	doc.Body().AddParagraph().SetStyle("Heading1")
	doc.Body().AddParagraph().AddRun("text")

	var buf bytes.Buffer
	if err := doc.Write(&buf); err != nil {
		t.Fatalf("Write: %v", err)
	}
	doc.Close()

	// Extract word/styles.xml from the zip and inspect raw XML bytes.
	r, err := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	if err != nil {
		t.Fatalf("zip.NewReader: %v", err)
	}

	var stylesXML string
	for _, f := range r.File {
		if f.Name == "word/styles.xml" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			stylesXML = string(data)
			break
		}
	}
	if stylesXML == "" {
		t.Fatal("word/styles.xml not found in package")
	}

	// Must contain proper w: prefixed elements.
	for _, want := range []string{"w:docDefaults", "w:rPrDefault", "w:latentStyles", "w:lsdException", "w:qFormat"} {
		if !strings.Contains(stylesXML, want) {
			t.Errorf("styles.xml missing %q", want)
		}
	}

	// Must NOT contain namespace corruption signatures.
	for _, bad := range []string{`xmlns="w"`, `_xmlns`, `xmlns:_xmlns`} {
		if strings.Contains(stylesXML, bad) {
			t.Errorf("styles.xml contains corrupted namespace: %q", bad)
		}
	}

	// Must contain the proper OOXML namespace URI declaration.
	if !strings.Contains(stylesXML, `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"`) {
		t.Error("styles.xml missing proper xmlns:w declaration")
	}
}

func TestStyles_XMLNamespaceIntegrity_AfterRoundTrip(t *testing.T) {
	// Create a fresh document, save, reopen, modify, save again.
	// The re-saved document must still have valid namespaces.
	doc, _ := New(nil)
	doc.Body().AddParagraph().SetStyle("Heading1")
	doc.Body().AddParagraph().AddRun("text")

	var buf1 bytes.Buffer
	doc.Write(&buf1)
	doc.Close()

	// Reopen and modify (triggers UnmarshalXML on styles.xml).
	doc2, err := OpenReader(bytes.NewReader(buf1.Bytes()), int64(buf1.Len()), nil)
	if err != nil {
		t.Fatalf("OpenReader: %v", err)
	}
	doc2.Body().AddParagraph().SetStyle("Heading2")

	var buf2 bytes.Buffer
	doc2.Write(&buf2)
	doc2.Close()

	// Extract and inspect word/styles.xml from the re-saved document.
	r, _ := zip.NewReader(bytes.NewReader(buf2.Bytes()), int64(buf2.Len()))
	var stylesXML string
	for _, f := range r.File {
		if f.Name == "word/styles.xml" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			stylesXML = string(data)
			break
		}
	}
	if stylesXML == "" {
		t.Fatal("word/styles.xml not found after round-trip")
	}

	for _, want := range []string{"w:docDefaults", "w:rPrDefault", "w:latentStyles", "w:lsdException"} {
		if !strings.Contains(stylesXML, want) {
			t.Errorf("re-saved styles.xml missing %q", want)
		}
	}
	for _, bad := range []string{`xmlns="w"`, `_xmlns`, `xmlns:_xmlns`} {
		if strings.Contains(stylesXML, bad) {
			t.Errorf("re-saved styles.xml contains corrupted namespace: %q", bad)
		}
	}
}
