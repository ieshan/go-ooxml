package docx

import (
	"bytes"
	"encoding/json"
	"strings"
	"testing"

	"github.com/ieshan/go-ooxml/common"
)

// ---------------------------------------------------------------------------
// Run export
// ---------------------------------------------------------------------------

func TestRunProseMirror_PlainText(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	r := doc.Body().AddParagraph().AddRun("hello")
	data, err := r.ProseMirror()
	if err != nil {
		t.Fatal(err)
	}
	var node PMNode
	json.Unmarshal(data, &node)
	if node.Type != "text" {
		t.Errorf("type = %q", node.Type)
	}
	if node.Text != "hello" {
		t.Errorf("text = %q", node.Text)
	}
	if len(node.Marks) != 0 {
		t.Errorf("marks = %d", len(node.Marks))
	}
}

func TestRunProseMirror_BoldItalic(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	r := doc.Body().AddParagraph().AddRun("styled")
	b := true
	r.SetBold(&b)
	r.SetItalic(&b)
	data, _ := r.ProseMirror()
	var node PMNode
	json.Unmarshal(data, &node)
	hasBold, hasItalic := false, false
	for _, m := range node.Marks {
		if m.Type == "bold" {
			hasBold = true
		}
		if m.Type == "italic" {
			hasItalic = true
		}
	}
	if !hasBold {
		t.Error("missing bold mark")
	}
	if !hasItalic {
		t.Error("missing italic mark")
	}
}

func TestRunProseMirror_TextStyle(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	r := doc.Body().AddParagraph().AddRun("styled")
	r.SetFontName("Arial")
	r.SetFontSize(18.0)
	r.SetFontColor(common.RGB(255, 0, 0))
	data, _ := r.ProseMirror()
	var node PMNode
	json.Unmarshal(data, &node)
	for _, m := range node.Marks {
		if m.Type == "textStyle" {
			if m.Attrs["fontFamily"] != "Arial" {
				t.Error("fontFamily")
			}
			if m.Attrs["fontSize"] != "18pt" {
				t.Errorf("fontSize = %v", m.Attrs["fontSize"])
			}
			if m.Attrs["color"] != "#FF0000" {
				t.Errorf("color = %v", m.Attrs["color"])
			}
			return
		}
	}
	t.Error("missing textStyle mark")
}

func TestRunProseMirror_Underline(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	r := doc.Body().AddParagraph().AddRun("underlined")
	u := true
	r.SetUnderline(&u)
	data, _ := r.ProseMirror()
	var node PMNode
	json.Unmarshal(data, &node)
	for _, m := range node.Marks {
		if m.Type == "underline" {
			return
		}
	}
	t.Error("missing underline mark")
}

func TestRunProseMirror_Strike(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	r := doc.Body().AddParagraph().AddRun("struck")
	s := true
	r.SetStrikethrough(&s)
	data, _ := r.ProseMirror()
	var node PMNode
	json.Unmarshal(data, &node)
	for _, m := range node.Marks {
		if m.Type == "strike" {
			return
		}
	}
	t.Error("missing strike mark")
}

func TestRunProseMirror_Code(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	r := doc.Body().AddParagraph().AddRun("code")
	r.SetFontName("Courier New")
	data, _ := r.ProseMirror()
	var node PMNode
	json.Unmarshal(data, &node)
	for _, m := range node.Marks {
		if m.Type == "code" {
			return
		}
	}
	t.Error("missing code mark")
}

// ---------------------------------------------------------------------------
// Paragraph export
// ---------------------------------------------------------------------------

func TestParagraphProseMirror_Plain(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	p := doc.Body().AddParagraph()
	p.AddRun("hello")
	data, _ := p.ProseMirror()
	var node PMNode
	json.Unmarshal(data, &node)
	if node.Type != "paragraph" {
		t.Errorf("type = %q", node.Type)
	}
	if len(node.Content) != 1 {
		t.Errorf("content = %d", len(node.Content))
	}
	if node.Content[0].Text != "hello" {
		t.Errorf("text = %q", node.Content[0].Text)
	}
}

func TestParagraphProseMirror_Heading(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	p := doc.Body().AddParagraph()
	p.SetStyle("Heading2")
	p.AddRun("Title")
	data, _ := p.ProseMirror()
	var node PMNode
	json.Unmarshal(data, &node)
	if node.Type != "heading" {
		t.Errorf("type = %q", node.Type)
	}
	if node.Attrs["level"] != float64(2) {
		t.Errorf("level = %v", node.Attrs["level"])
	}
}

func TestParagraphProseMirror_Alignment(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	p := doc.Body().AddParagraph()
	p.AddRun("centered")
	p.Format().SetAlignment("center")
	data, _ := p.ProseMirror()
	var node PMNode
	json.Unmarshal(data, &node)
	if node.Attrs["textAlign"] != "center" {
		t.Errorf("textAlign = %v", node.Attrs["textAlign"])
	}
}

func TestParagraphProseMirror_Hyperlink(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	p := doc.Body().AddParagraph()
	p.AddRun("Visit ")
	p.AddHyperlink("https://example.com", "Example")
	data, _ := p.ProseMirror()
	var node PMNode
	json.Unmarshal(data, &node)
	if len(node.Content) < 2 {
		t.Fatal("expected 2+ content nodes")
	}
	linkNode := node.Content[1]
	for _, m := range linkNode.Marks {
		if m.Type == "link" && m.Attrs["href"] == "https://example.com" {
			return
		}
	}
	t.Error("missing link mark")
}

func TestParagraphProseMirror_Blockquote(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	p := doc.Body().AddParagraph()
	p.SetStyle("Quote")
	p.AddRun("A quote")
	data, _ := p.ProseMirror()
	var node PMNode
	json.Unmarshal(data, &node)
	if node.Type != "blockquote" {
		t.Errorf("type = %q, want blockquote", node.Type)
	}
}

// ---------------------------------------------------------------------------
// Table export
// ---------------------------------------------------------------------------

func TestTableProseMirror_Basic(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	tbl := doc.Body().AddTable(2, 2)
	tbl.Cell(0, 0).AddParagraph().AddRun("H1")
	tbl.Cell(0, 1).AddParagraph().AddRun("H2")
	tbl.Cell(1, 0).AddParagraph().AddRun("A")
	tbl.Cell(1, 1).AddParagraph().AddRun("B")
	data, _ := tbl.ProseMirror()
	var node PMNode
	json.Unmarshal(data, &node)
	if node.Type != "table" {
		t.Errorf("type = %q", node.Type)
	}
	if len(node.Content) != 2 {
		t.Errorf("rows = %d", len(node.Content))
	}
	firstRow := node.Content[0]
	if firstRow.Content[0].Type != "tableHeader" {
		t.Errorf("first row cell type = %q", firstRow.Content[0].Type)
	}
	secondRow := node.Content[1]
	if secondRow.Content[0].Type != "tableCell" {
		t.Errorf("second row cell type = %q", secondRow.Content[0].Type)
	}
}

// ---------------------------------------------------------------------------
// Document export
// ---------------------------------------------------------------------------

func TestDocumentProseMirror_FullDocument(t *testing.T) {
	doc, _ := New(&Config{Author: "Alice"})
	defer doc.Close()
	doc.Properties().SetTitle("Test Doc")
	h := doc.Body().AddParagraph()
	h.SetStyle("Heading1")
	h.AddRun("Title")
	doc.Body().AddParagraph().AddRun("Body text")
	data, err := doc.ProseMirror(nil)
	if err != nil {
		t.Fatal(err)
	}
	var node PMNode
	json.Unmarshal(data, &node)
	if node.Type != "doc" {
		t.Errorf("type = %q", node.Type)
	}
	if node.Attrs["title"] != "Test Doc" {
		t.Errorf("title = %v", node.Attrs["title"])
	}
	if node.Attrs["author"] != "Alice" {
		t.Errorf("author = %v", node.Attrs["author"])
	}
	if len(node.Content) != 2 {
		t.Errorf("content = %d", len(node.Content))
	}
}

func TestDocumentProseMirror_WithHeaders(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	for section := range doc.Sections() {
		hdr := section.AddHeader(HdrFtrDefault)
		hdr.AddParagraph().AddRun("Header Text")
	}
	doc.Body().AddParagraph().AddRun("Body")
	data, _ := doc.ProseMirror(&ProseMirrorOptions{UseTipTapNames: true, IncludeHeaders: true})
	var node PMNode
	json.Unmarshal(data, &node)
	if node.Attrs["header"] == nil {
		t.Error("missing header in doc attrs")
	}
}

// ---------------------------------------------------------------------------
// Comment / Revision / SearchResult export
// ---------------------------------------------------------------------------

func TestCommentProseMirror(t *testing.T) {
	doc, _ := New(&Config{Author: "Bot"})
	defer doc.Close()
	r := doc.Body().AddParagraph().AddRun("text")
	c := doc.AddComment(r, r, "review note")
	data, _ := c.ProseMirror()
	var node PMNode
	json.Unmarshal(data, &node)
	if node.Type != "doc" {
		t.Errorf("type = %q", node.Type)
	}
	if len(node.Content) == 0 {
		t.Error("empty content")
	}
}

func TestRevisionProposedProseMirror(t *testing.T) {
	doc, _ := New(&Config{Author: "E"})
	defer doc.Close()
	r := doc.Body().AddParagraph().AddRun("old")
	rev := doc.AddRevision(r, r, "new")
	data, _ := rev.ProposedProseMirror()
	var node PMNode
	json.Unmarshal(data, &node)
	if node.Type != "doc" {
		t.Errorf("type = %q", node.Type)
	}
}

func TestSearchResultProseMirror(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	p := doc.Body().AddParagraph()
	r := p.AddRun("bold text")
	b := true
	r.SetBold(&b)
	results := doc.Find("bold text")
	if len(results) == 0 {
		t.Fatal("no results")
	}
	data, _ := results[0].ProseMirror()
	var node PMNode
	json.Unmarshal(data, &node)
	if node.Text != "bold text" {
		t.Errorf("text = %q", node.Text)
	}
	hasBold := false
	for _, m := range node.Marks {
		if m.Type == "bold" {
			hasBold = true
		}
	}
	if !hasBold {
		t.Error("missing bold")
	}
}

func TestBodyProseMirror(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	doc.Body().AddParagraph().AddRun("body text")
	data, err := doc.Body().ProseMirror()
	if err != nil {
		t.Fatal(err)
	}
	var node PMNode
	json.Unmarshal(data, &node)
	if node.Type != "doc" {
		t.Errorf("type = %q", node.Type)
	}
	if len(node.Content) != 1 {
		t.Errorf("content = %d", len(node.Content))
	}
}

func TestHeaderFooterProseMirror(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	for section := range doc.Sections() {
		hdr := section.AddHeader(HdrFtrDefault)
		hdr.AddParagraph().AddRun("Header Content")
		data, err := hdr.ProseMirror()
		if err != nil {
			t.Fatal(err)
		}
		var node PMNode
		json.Unmarshal(data, &node)
		if node.Type != "doc" {
			t.Errorf("type = %q", node.Type)
		}
		if len(node.Content) == 0 {
			t.Error("empty content")
		}
		if node.Content[0].Content[0].Text != "Header Content" {
			t.Errorf("text = %q", node.Content[0].Content[0].Text)
		}
	}
}

func TestRevisionOriginalProseMirror(t *testing.T) {
	doc, _ := New(&Config{Author: "E"})
	defer doc.Close()
	r := doc.Body().AddParagraph().AddRun("original text")
	doc.AddRevision(r, r, "replacement")

	// Find the delete revision.
	var delRev *Revision
	for rev := range doc.Revisions() {
		if rev.Type() == RevisionDelete {
			delRev = rev
			break
		}
	}
	if delRev == nil {
		t.Fatal("delete revision not found")
	}
	data, err := delRev.OriginalProseMirror()
	if err != nil {
		t.Fatal(err)
	}
	var node PMNode
	json.Unmarshal(data, &node)
	if node.Type != "doc" {
		t.Errorf("type = %q", node.Type)
	}
	if len(node.Content) == 0 {
		t.Error("empty content")
	}
}

// ---------------------------------------------------------------------------
// Round-trip
// ---------------------------------------------------------------------------

func TestProseMirror_RoundTrip_RichDocument(t *testing.T) {
	doc, _ := New(&Config{Author: "Alice"})
	defer doc.Close()
	doc.Properties().SetTitle("Round Trip Test")
	h := doc.Body().AddParagraph()
	h.SetStyle("Heading1")
	h.AddRun("Title")
	p := doc.Body().AddParagraph()
	r := p.AddRun("Bold text")
	b := true
	r.SetBold(&b)
	p.AddRun(" and ")
	r2 := p.AddRun("italic")
	it := true
	r2.SetItalic(&it)
	tbl := doc.Body().AddTable(1, 2)
	tbl.Cell(0, 0).AddParagraph().AddRun("A")
	tbl.Cell(0, 1).AddParagraph().AddRun("B")

	data, err := doc.ProseMirror(nil)
	if err != nil {
		t.Fatal(err)
	}

	doc2, err := FromProseMirror(data, nil)
	if err != nil {
		t.Fatal(err)
	}
	defer doc2.Close()

	text := doc2.Text(nil)
	if !strings.Contains(text, "Title") {
		t.Error("missing title")
	}
	if !strings.Contains(text, "Bold text") {
		t.Error("missing bold text")
	}
	if !strings.Contains(text, "italic") {
		t.Error("missing italic")
	}
	for p := range doc2.Body().Paragraphs() {
		for r := range p.Runs() {
			if r.Text() == "Bold text" && (r.Bold() == nil || !*r.Bold()) {
				t.Error("bold lost in round-trip")
			}
		}
	}
}

func TestProseMirror_RoundTrip_WriteThenRead(t *testing.T) {
	doc, _ := New(nil)
	doc.Body().AddParagraph().AddRun("persistent")
	pmData, _ := doc.ProseMirror(nil)
	doc.Close()

	doc2, _ := FromProseMirror(pmData, nil)
	var buf bytes.Buffer
	doc2.WriteTo(&buf)
	doc2.Close()

	doc3, _ := OpenReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()), nil)
	defer doc3.Close()
	if doc3.Text(nil) != "persistent" {
		t.Errorf("text = %q after full round-trip", doc3.Text(nil))
	}
}

func TestProseMirror_VanillaNames(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	r := doc.Body().AddParagraph().AddRun("test")
	b := true
	r.SetBold(&b)
	data, _ := doc.ProseMirror(&ProseMirrorOptions{UseTipTapNames: false})
	s := string(data)
	if !strings.Contains(s, `"strong"`) {
		t.Errorf("expected 'strong' in vanilla mode, got: %s", s)
	}
}

// ---------------------------------------------------------------------------
// Examples
// ---------------------------------------------------------------------------

func ExampleDocument_ProseMirror() {
	doc, _ := New(nil)
	defer doc.Close()
	doc.Body().AddParagraph().AddRun("Hello World")
	data, _ := doc.ProseMirror(nil)
	_ = data // ProseMirror JSON bytes
}

func ExampleFromProseMirror() {
	pmJSON := []byte(`{"type":"doc","content":[{"type":"paragraph","content":[{"type":"text","text":"Hello"}]}]}`)
	doc, _ := FromProseMirror(pmJSON, nil)
	defer doc.Close()
	_ = doc.Text(nil) // "Hello"
}
