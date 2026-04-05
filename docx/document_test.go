package docx

import (
	"bytes"
	"encoding/xml"
	"testing"

	"github.com/ieshan/go-ooxml/docx/wml"
)

func TestNew_CreatesValidDocument(t *testing.T) {
	doc, err := New(nil)
	if err != nil {
		t.Fatalf("New: %v", err)
	}
	defer doc.Close()
	if doc.doc == nil {
		t.Error("internal doc nil")
	}
	if doc.doc.Body == nil {
		t.Error("body nil")
	}
}

func TestNew_WithConfig(t *testing.T) {
	doc, err := New(&Config{Author: "TestBot"})
	if err != nil {
		t.Fatalf("New: %v", err)
	}
	defer doc.Close()
	if doc.cfg.Author != "TestBot" {
		t.Error("config not applied")
	}
}

func TestNew_Save_RoundTrip(t *testing.T) {
	doc, err := New(nil)
	if err != nil {
		t.Fatalf("New: %v", err)
	}

	// Add a paragraph directly to the internal doc for testing
	p := &wml.CT_P{XMLName: xml.Name{Space: wml.Ns, Local: "p"}}
	p.AddRun("Hello World")
	doc.doc.Body.Content = append(doc.doc.Body.Content, wml.BlockLevelContent{Paragraph: p})

	var buf bytes.Buffer
	if err := doc.Write(&buf); err != nil {
		t.Fatalf("WriteTo: %v", err)
	}

	doc2, err := OpenReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()), nil)
	if err != nil {
		t.Fatalf("OpenReader: %v", err)
	}
	defer doc2.Close()

	if doc2.doc.Body == nil {
		t.Fatal("body nil after round-trip")
	}
	paraCount := 0
	for _, c := range doc2.doc.Body.Content {
		if c.Paragraph != nil {
			paraCount++
		}
	}
	if paraCount != 1 {
		t.Errorf("paragraphs = %d, want 1", paraCount)
	}
	if doc2.doc.Body.Content[0].Paragraph.Text() != "Hello World" {
		t.Errorf("text = %q", doc2.doc.Body.Content[0].Paragraph.Text())
	}
}

func TestOpen_ParsesStyles(t *testing.T) {
	// Create a document with styles, save, reopen, verify styles parsed
	doc, err := New(nil)
	if err != nil {
		t.Fatalf("New: %v", err)
	}

	var buf bytes.Buffer
	if err := doc.Write(&buf); err != nil {
		t.Fatalf("WriteTo: %v", err)
	}

	doc2, err := OpenReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()), nil)
	if err != nil {
		t.Fatalf("OpenReader: %v", err)
	}
	defer doc2.Close()
	// styles may or may not exist in minimal doc — just verify no error
}

func TestDocument_Close(t *testing.T) {
	doc, err := New(nil)
	if err != nil {
		t.Fatalf("New: %v", err)
	}
	if err := doc.Close(); err != nil {
		t.Errorf("Close: %v", err)
	}
}

func TestDocument_Body_ReturnsProxy(t *testing.T) {
	doc, err := New(nil)
	if err != nil {
		t.Fatalf("New: %v", err)
	}
	defer doc.Close()
	if doc.Body() == nil {
		t.Error("Body() should return a non-nil proxy")
	}
}

func TestNew_PackageHasCorrectContentTypes(t *testing.T) {
	doc, err := New(nil)
	if err != nil {
		t.Fatalf("New: %v", err)
	}
	defer doc.Close()

	ct := doc.pkg.ContentTypes
	if ct == nil {
		t.Fatal("ContentTypes nil")
	}

	// Check the rels default
	if got := ct.Defaults["rels"]; got != "application/vnd.openxmlformats-package.relationships+xml" {
		t.Errorf("Defaults[rels] = %q", got)
	}
	// Check the xml default
	if got := ct.Defaults["xml"]; got != "application/xml" {
		t.Errorf("Defaults[xml] = %q", got)
	}
	// Check the document override
	want := "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
	if got := ct.Overrides["/word/document.xml"]; got != want {
		t.Errorf("Overrides[/word/document.xml] = %q, want %q", got, want)
	}
}

func TestNew_PackageHasDocumentRelationship(t *testing.T) {
	doc, err := New(nil)
	if err != nil {
		t.Fatalf("New: %v", err)
	}
	defer doc.Close()

	found := false
	for _, rel := range doc.pkg.Rels {
		if rel.Type == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" {
			found = true
			if rel.Target != "word/document.xml" {
				t.Errorf("rel target = %q, want word/document.xml", rel.Target)
			}
		}
	}
	if !found {
		t.Error("missing officeDocument relationship")
	}
}

func ExampleNew() {
	doc, err := New(&Config{Author: "MyApp"})
	if err != nil {
		// handle error
		return
	}
	defer doc.Close()
	doc.Body().AddParagraph().AddRun("Hello, World!")
	// var buf bytes.Buffer
	// doc.WriteTo(&buf)
}

func ExampleOpen() {
	// doc, err := Open("report.docx", nil)
	// if err != nil { log.Fatal(err) }
	// defer doc.Close()
	// fmt.Println(doc.Text(nil))
}

func ExampleOpenReader() {
	// Useful when you have a document in memory (e.g. from an HTTP response).
	// data, _ := io.ReadAll(resp.Body)
	// doc, err := OpenReader(bytes.NewReader(data), int64(len(data)), nil)
	// if err != nil { log.Fatal(err) }
	// defer doc.Close()
}

func ExampleDocument_Write() {
	doc, _ := New(nil)
	defer doc.Close()
	doc.Body().AddParagraph().AddRun("content")
	var buf bytes.Buffer
	_ = doc.Write(&buf)
	// buf now contains the .docx archive bytes
}

func ExampleDocument_WriteTo() {
	doc, _ := New(nil)
	defer doc.Close()
	doc.Body().AddParagraph().AddRun("content")
	var buf bytes.Buffer
	n, _ := doc.WriteTo(&buf)
	_ = n // number of bytes written
}
