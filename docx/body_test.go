package docx

import (
	"bytes"
	"testing"
)

func TestBody_AddParagraph(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	p := doc.Body().AddParagraph()
	if p == nil {
		t.Fatal("nil paragraph")
	}
}

func TestBody_Paragraphs(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	doc.Body().AddParagraph().AddRun("one")
	doc.Body().AddParagraph().AddRun("two")
	count := 0
	for range doc.Body().Paragraphs() {
		count++
	}
	if count != 2 {
		t.Errorf("count = %d, want 2", count)
	}
}

func TestBody_Text(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	doc.Body().AddParagraph().AddRun("Hello")
	doc.Body().AddParagraph().AddRun("World")
	text := doc.Body().Text()
	if text != "Hello\nWorld" {
		t.Errorf("Text() = %q", text)
	}
}

func TestBody_RoundTrip(t *testing.T) {
	doc, _ := New(nil)
	doc.Body().AddParagraph().AddRun("test content")
	var buf bytes.Buffer
	if err := doc.Write(&buf); err != nil {
		t.Fatalf("WriteTo: %v", err)
	}
	doc2, err := OpenReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()), nil)
	if err != nil {
		t.Fatalf("OpenReader: %v", err)
	}
	defer doc2.Close()
	count := 0
	for p := range doc2.Body().Paragraphs() {
		if p.Text() == "test content" {
			count++
		}
	}
	if count != 1 {
		t.Errorf("count = %d", count)
	}
}

func TestBody_Tables_Empty(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	count := 0
	for range doc.Body().Tables() {
		count++
	}
	if count != 0 {
		t.Errorf("count = %d, want 0", count)
	}
}

func TestBody_AddTable(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	tbl := doc.Body().AddTable(3, 3)
	if tbl == nil {
		t.Error("AddTable should return a table")
	}
}

func TestBody_Markdown_Empty(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	if doc.Body().Markdown() != "" {
		t.Error("empty body Markdown() should return empty string")
	}
}

func TestBody_Clear(t *testing.T) {
	doc, err := New(nil)
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()

	h := doc.Body().AddParagraph()
	h.SetStyle("Heading1")
	h.AddRun("Title")
	doc.Body().AddParagraph().AddRun("Body text")
	doc.Body().AddTable(2, 2)

	doc.Body().Clear()

	count := 0
	for range doc.Body().Paragraphs() {
		count++
	}
	if count != 0 {
		t.Errorf("paragraphs after clear = %d, want 0", count)
	}

	tblCount := 0
	for range doc.Body().Tables() {
		tblCount++
	}
	if tblCount != 0 {
		t.Errorf("tables after clear = %d, want 0", tblCount)
	}
}

func TestBody_Clear_PreservesStyles(t *testing.T) {
	doc, err := New(nil)
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()

	doc.Body().AddParagraph().AddRun("content")

	beforeCount := 0
	for range doc.Styles() {
		beforeCount++
	}

	doc.Body().Clear()

	afterCount := 0
	for range doc.Styles() {
		afterCount++
	}
	if afterCount != beforeCount {
		t.Errorf("styles changed: before=%d, after=%d", beforeCount, afterCount)
	}
}

func TestBody_Clear_EmptyDoc(t *testing.T) {
	doc, err := New(nil)
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()

	doc.Body().Clear()

	count := 0
	for range doc.Body().Paragraphs() {
		count++
	}
	if count != 0 {
		t.Errorf("paragraphs = %d, want 0", count)
	}
}

func ExampleBody_AddParagraph() {
	doc, _ := New(nil)
	defer doc.Close()
	p := doc.Body().AddParagraph()
	p.AddRun("Hello World")
	// doc.WriteTo(w)
}

func ExampleBody_Clear() {
	doc, _ := New(nil)
	defer doc.Close()

	doc.Body().AddParagraph().AddRun("old content")
	doc.Body().AddTable(2, 2)

	// Remove all body content while preserving section properties.
	doc.Body().Clear()

	// Body is now empty — add fresh content.
	doc.Body().AddParagraph().AddRun("new content")
}
