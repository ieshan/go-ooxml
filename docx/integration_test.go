package docx

import (
	"bytes"
	"strings"
	"testing"
)

func TestIntegration_SimpleDocument_RoundTrip(t *testing.T) {
	// Create a document with multiple paragraphs
	doc, _ := New(&Config{Author: "Test"})
	doc.Body().AddParagraph().AddRun("First paragraph")
	p2 := doc.Body().AddParagraph()
	p2.SetStyle("Heading1")
	p2.AddRun("Section Title")
	doc.Body().AddParagraph().AddRun("Body text here")

	// Save and re-open
	var buf bytes.Buffer
	if err := doc.Write(&buf); err != nil {
		t.Fatalf("WriteTo: %v", err)
	}
	doc2, err := OpenReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()), nil)
	if err != nil {
		t.Fatalf("OpenReader: %v", err)
	}
	defer doc2.Close()

	// Verify content
	texts := []string{}
	for p := range doc2.Body().Paragraphs() {
		texts = append(texts, p.Text())
	}
	if len(texts) != 3 {
		t.Errorf("paragraphs = %d", len(texts))
	}
	if texts[0] != "First paragraph" {
		t.Errorf("p0 = %q", texts[0])
	}
	if texts[1] != "Section Title" {
		t.Errorf("p1 = %q", texts[1])
	}
}

func TestIntegration_FormattedRuns_RoundTrip(t *testing.T) {
	doc, _ := New(nil)
	p := doc.Body().AddParagraph()
	r1 := p.AddRun("normal ")
	r2 := p.AddRun("bold ")
	b := true
	r2.SetBold(&b)
	r3 := p.AddRun("italic")
	r3.SetItalic(&b)

	var buf bytes.Buffer
	doc.Write(&buf)
	doc2, _ := OpenReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()), nil)
	defer doc2.Close()

	// Suppress "declared and not used" error for r1
	_ = r1

	for p := range doc2.Body().Paragraphs() {
		for r := range p.Runs() {
			switch r.Text() {
			case "normal ":
				if r.Bold() != nil {
					t.Error("normal should not be bold")
				}
			case "bold ":
				if r.Bold() == nil || !*r.Bold() {
					t.Error("bold should be bold")
				}
			case "italic":
				if r.Italic() == nil || !*r.Italic() {
					t.Error("italic should be italic")
				}
			}
		}
	}
}

func TestIntegration_Table_RoundTrip(t *testing.T) {
	doc, _ := New(nil)
	tbl := doc.Body().AddTable(2, 3)
	tbl.Cell(0, 0).AddParagraph().AddRun("R0C0")
	tbl.Cell(0, 1).AddParagraph().AddRun("R0C1")
	tbl.Cell(0, 2).AddParagraph().AddRun("R0C2")
	tbl.Cell(1, 0).AddParagraph().AddRun("R1C0")
	tbl.Cell(1, 1).AddParagraph().AddRun("R1C1")
	tbl.Cell(1, 2).AddParagraph().AddRun("R1C2")

	var buf bytes.Buffer
	doc.Write(&buf)
	doc2, _ := OpenReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()), nil)
	defer doc2.Close()

	count := 0
	for range doc2.Body().Tables() {
		count++
	}
	if count != 1 {
		t.Errorf("tables = %d", count)
	}

	text := doc2.Text(nil)
	for _, s := range []string{"R0C0", "R0C1", "R0C2", "R1C0", "R1C1", "R1C2"} {
		if !strings.Contains(text, s) {
			t.Errorf("missing %q in %q", s, text)
		}
	}
}

func TestIntegration_Comments_RoundTrip(t *testing.T) {
	doc, _ := New(&Config{Author: "Reviewer"})
	r := doc.Body().AddParagraph().AddRun("Review this section")
	doc.AddComment(r, r, "Needs revision")

	var buf bytes.Buffer
	doc.Write(&buf)
	doc2, _ := OpenReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()), nil)
	defer doc2.Close()

	count := 0
	for c := range doc2.Comments() {
		if c.Text() == "Needs revision" && c.Author() == "Reviewer" {
			count++
		}
	}
	if count != 1 {
		t.Errorf("comments = %d", count)
	}
}

func TestIntegration_Revisions_RoundTrip(t *testing.T) {
	doc, _ := New(&Config{Author: "Editor"})
	r := doc.Body().AddParagraph().AddRun("draft")
	doc.AddRevision(r, r, "final")

	var buf bytes.Buffer
	doc.Write(&buf)
	doc2, _ := OpenReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()), nil)
	defer doc2.Close()

	count := 0
	for range doc2.Revisions() {
		count++
	}
	if count < 1 {
		t.Errorf("revisions = %d", count)
	}
}

func TestIntegration_SearchAndModify_RoundTrip(t *testing.T) {
	doc, _ := New(nil)
	doc.Body().AddParagraph().AddRun("Hello World")
	doc.Body().AddParagraph().AddRun("Hello Again")

	// Replace all "Hello" with "Hi"
	doc.Search("Hello").ReplaceText("Hi")

	var buf bytes.Buffer
	doc.Write(&buf)
	doc2, _ := OpenReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()), nil)
	defer doc2.Close()

	text := doc2.Text(nil)
	if strings.Contains(text, "Hello") {
		t.Error("Hello should be replaced")
	}
	if !strings.Contains(text, "Hi") {
		t.Error("Hi should be present")
	}
}

func TestIntegration_ComplexDocument_RoundTrip(t *testing.T) {
	doc, _ := New(&Config{Author: "TestBot"})

	// Title
	title := doc.Body().AddParagraph()
	title.SetStyle("Heading1")
	title.AddRun("Quarterly Report")

	// Body with formatting
	p := doc.Body().AddParagraph()
	p.AddRun("Revenue increased by ")
	bold := doc.Body().AddParagraph().AddRun("42%")
	bold.SetBold(new(true))

	// Table
	tbl := doc.Body().AddTable(2, 2)
	tbl.Cell(0, 0).AddParagraph().AddRun("Q1")
	tbl.Cell(0, 1).AddParagraph().AddRun("$1M")
	tbl.Cell(1, 0).AddParagraph().AddRun("Q2")
	tbl.Cell(1, 1).AddParagraph().AddRun("$1.5M")

	// Comment
	r := doc.Body().AddParagraph().AddRun("TODO: verify numbers")
	doc.AddComment(r, r, "Double-check with finance")

	// Save, re-open, verify
	var buf bytes.Buffer
	doc.Write(&buf)
	doc2, err := OpenReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()), nil)
	if err != nil {
		t.Fatalf("OpenReader: %v", err)
	}
	defer doc2.Close()

	text := doc2.Text(&TextOptions{IncludeComments: true})
	for _, s := range []string{"Quarterly Report", "Revenue", "Q1", "$1M", "TODO", "Double-check"} {
		if !strings.Contains(text, s) {
			t.Errorf("missing %q", s)
		}
	}

	// Verify markdown extraction works
	md := doc2.Markdown(nil)
	if !strings.Contains(md, "# Quarterly Report") {
		t.Error("heading not in markdown")
	}
}
