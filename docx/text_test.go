package docx

import (
	"strings"
	"testing"
)

func TestDocument_Text_Basic(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	doc.Body().AddParagraph().AddRun("Hello")
	doc.Body().AddParagraph().AddRun("World")
	text := doc.Text(nil)
	if text != "Hello\nWorld" {
		t.Errorf("text = %q", text)
	}
}

func TestDocument_Text_WithTable(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	doc.Body().AddParagraph().AddRun("Before table")
	tbl := doc.Body().AddTable(1, 2)
	tbl.Cell(0, 0).AddParagraph().AddRun("A")
	tbl.Cell(0, 1).AddParagraph().AddRun("B")
	doc.Body().AddParagraph().AddRun("After table")

	text := doc.Text(nil)
	// Should contain all text
	if !strings.Contains(text, "Before table") {
		t.Error("missing before")
	}
	if !strings.Contains(text, "A") {
		t.Error("missing cell A")
	}
	if !strings.Contains(text, "B") {
		t.Error("missing cell B")
	}
	if !strings.Contains(text, "After table") {
		t.Error("missing after")
	}
}

func TestDocument_Text_WithComments(t *testing.T) {
	doc, _ := New(&Config{Author: "Bot"})
	defer doc.Close()
	r := doc.Body().AddParagraph().AddRun("text")
	doc.AddComment(r, r, "review this")

	// Without comments
	text := doc.Text(nil)
	if strings.Contains(text, "review") {
		t.Error("should not include comments by default")
	}

	// With comments
	text = doc.Text(&TextOptions{IncludeComments: true})
	if !strings.Contains(text, "review this") {
		t.Error("should include comments")
	}
}

func TestDocument_Text_Empty(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	text := doc.Text(nil)
	if text != "" {
		t.Errorf("empty doc text = %q", text)
	}
}

func TestDocument_Text_MultipleFormats(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	p := doc.Body().AddParagraph()
	p.AddRun("Hello ")
	r := p.AddRun("World")
	bold := true
	r.SetBold(&bold)
	// Bold formatting shouldn't affect plain text
	if doc.Text(nil) != "Hello World" {
		t.Errorf("text = %q", doc.Text(nil))
	}
}

func TestDocument_Text_TableCellsOrdered(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	tbl := doc.Body().AddTable(2, 2)
	tbl.Cell(0, 0).AddParagraph().AddRun("R0C0")
	tbl.Cell(0, 1).AddParagraph().AddRun("R0C1")
	tbl.Cell(1, 0).AddParagraph().AddRun("R1C0")
	tbl.Cell(1, 1).AddParagraph().AddRun("R1C1")

	text := doc.Text(nil)
	// Rows separated by newlines, cells within a row by tabs.
	if !strings.Contains(text, "R0C0\tR0C1") {
		t.Errorf("expected tab between row0 cells, got %q", text)
	}
	if !strings.Contains(text, "R1C0\tR1C1") {
		t.Errorf("expected tab between row1 cells, got %q", text)
	}
	// Row 0 should appear before row 1 with a newline.
	row0End := strings.Index(text, "R0C1")
	row1Start := strings.Index(text, "R1C0")
	if row0End == -1 || row1Start == -1 || row0End >= row1Start {
		t.Errorf("unexpected ordering or missing rows in %q", text)
	}
}

func TestDocument_Text_IncludeHeadersFooters_NoOp(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	doc.Body().AddParagraph().AddRun("body text")

	// Headers/footers are not yet implemented — options are no-ops.
	text := doc.Text(&TextOptions{IncludeHeaders: true, IncludeFooters: true})
	if text != "body text" {
		t.Errorf("text = %q, want %q", text, "body text")
	}
}

func ExampleDocument_Text() {
	doc, _ := New(nil)
	defer doc.Close()
	doc.Body().AddParagraph().AddRun("First paragraph")
	doc.Body().AddParagraph().AddRun("Second paragraph")
	_ = doc.Text(nil) // "First paragraph\nSecond paragraph"
}
