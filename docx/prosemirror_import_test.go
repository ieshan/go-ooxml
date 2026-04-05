package docx

import (
	"strings"
	"testing"
)

func TestFromProseMirror_BasicParagraph(t *testing.T) {
	pmJSON := `{"type":"doc","content":[{"type":"paragraph","content":[{"type":"text","text":"hello"}]}]}`
	doc, err := FromProseMirror([]byte(pmJSON), nil)
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()
	if doc.Text(nil) != "hello" {
		t.Errorf("text = %q", doc.Text(nil))
	}
}

func TestFromProseMirror_BoldItalic(t *testing.T) {
	pmJSON := `{"type":"doc","content":[{"type":"paragraph","content":[
		{"type":"text","text":"bold","marks":[{"type":"bold"}]},
		{"type":"text","text":" and "},
		{"type":"text","text":"italic","marks":[{"type":"italic"}]}
	]}]}`
	doc, err := FromProseMirror([]byte(pmJSON), nil)
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()
	for p := range doc.Body().Paragraphs() {
		for r := range p.Runs() {
			if r.Text() == "bold" && (r.Bold() == nil || !*r.Bold()) {
				t.Error("bold not set")
			}
			if r.Text() == "italic" && (r.Italic() == nil || !*r.Italic()) {
				t.Error("italic not set")
			}
		}
	}
}

func TestFromProseMirror_Heading(t *testing.T) {
	pmJSON := `{"type":"doc","content":[{"type":"heading","attrs":{"level":2},"content":[{"type":"text","text":"Title"}]}]}`
	doc, _ := FromProseMirror([]byte(pmJSON), nil)
	defer doc.Close()
	for p := range doc.Body().Paragraphs() {
		if p.Style() != "Heading2" {
			t.Errorf("style = %q", p.Style())
		}
	}
}

func TestFromProseMirror_TextStyle(t *testing.T) {
	pmJSON := `{"type":"doc","content":[{"type":"paragraph","content":[
		{"type":"text","text":"styled","marks":[{"type":"textStyle","attrs":{"color":"#FF0000","fontSize":"18pt","fontFamily":"Arial"}}]}
	]}]}`
	doc, _ := FromProseMirror([]byte(pmJSON), nil)
	defer doc.Close()
	for p := range doc.Body().Paragraphs() {
		for r := range p.Runs() {
			if r.FontName() == nil || *r.FontName() != "Arial" {
				t.Error("fontFamily")
			}
			if r.FontSize() == nil || *r.FontSize() != 18.0 {
				t.Errorf("fontSize = %v", r.FontSize())
			}
			if r.FontColor() == nil || r.FontColor().Hex() != "FF0000" {
				t.Error("color")
			}
		}
	}
}

func TestFromProseMirror_Table(t *testing.T) {
	pmJSON := `{"type":"doc","content":[{"type":"table","content":[
		{"type":"tableRow","content":[
			{"type":"tableCell","attrs":{"colspan":1,"rowspan":1},"content":[{"type":"paragraph","content":[{"type":"text","text":"A"}]}]},
			{"type":"tableCell","attrs":{"colspan":1,"rowspan":1},"content":[{"type":"paragraph","content":[{"type":"text","text":"B"}]}]}
		]}
	]}]}`
	doc, _ := FromProseMirror([]byte(pmJSON), nil)
	defer doc.Close()
	count := 0
	for range doc.Body().Tables() {
		count++
	}
	if count != 1 {
		t.Errorf("tables = %d", count)
	}
}

func TestFromProseMirror_Link(t *testing.T) {
	pmJSON := `{"type":"doc","content":[{"type":"paragraph","content":[
		{"type":"text","text":"click","marks":[{"type":"link","attrs":{"href":"https://example.com"}}]}
	]}]}`
	doc, _ := FromProseMirror([]byte(pmJSON), nil)
	defer doc.Close()
	for p := range doc.Body().Paragraphs() {
		for h := range p.Hyperlinks() {
			if h.URL() != "https://example.com" {
				t.Errorf("url = %q", h.URL())
			}
			if h.Text() != "click" {
				t.Errorf("text = %q", h.Text())
			}
			return
		}
	}
	t.Error("hyperlink not found")
}

func TestFromProseMirror_VanillaNames(t *testing.T) {
	pmJSON := `{"type":"doc","content":[{"type":"paragraph","content":[
		{"type":"text","text":"test","marks":[{"type":"strong"},{"type":"em"}]}
	]}]}`
	doc, _ := FromProseMirror([]byte(pmJSON), nil)
	defer doc.Close()
	for p := range doc.Body().Paragraphs() {
		for r := range p.Runs() {
			if r.Bold() == nil || !*r.Bold() {
				t.Error("strong not mapped to bold")
			}
			if r.Italic() == nil || !*r.Italic() {
				t.Error("em not mapped to italic")
			}
		}
	}
}

func TestFromProseMirror_BulletList(t *testing.T) {
	pmJSON := `{"type":"doc","content":[{"type":"bulletList","content":[
		{"type":"listItem","content":[{"type":"paragraph","content":[{"type":"text","text":"Item 1"}]}]},
		{"type":"listItem","content":[{"type":"paragraph","content":[{"type":"text","text":"Item 2"}]}]}
	]}]}`
	doc, _ := FromProseMirror([]byte(pmJSON), nil)
	defer doc.Close()
	text := doc.Text(nil)
	if !strings.Contains(text, "Item 1") || !strings.Contains(text, "Item 2") {
		t.Errorf("text = %q", text)
	}
}

func TestFromProseMirror_OrderedList(t *testing.T) {
	pmJSON := `{"type":"doc","content":[{"type":"orderedList","attrs":{"start":1},"content":[
		{"type":"listItem","content":[{"type":"paragraph","content":[{"type":"text","text":"Step 1"}]}]}
	]}]}`
	doc, _ := FromProseMirror([]byte(pmJSON), nil)
	defer doc.Close()
	text := doc.Text(nil)
	if !strings.Contains(text, "Step 1") {
		t.Errorf("text = %q", text)
	}
}

func TestFromProseMirror_Blockquote(t *testing.T) {
	pmJSON := `{"type":"doc","content":[{"type":"blockquote","content":[
		{"type":"paragraph","content":[{"type":"text","text":"A quote"}]}
	]}]}`
	doc, _ := FromProseMirror([]byte(pmJSON), nil)
	defer doc.Close()
	for p := range doc.Body().Paragraphs() {
		if p.Style() == "Quote" {
			return
		}
	}
	t.Error("Quote style not found")
}

func TestFromProseMirror_CodeBlock(t *testing.T) {
	pmJSON := `{"type":"doc","content":[{"type":"codeBlock","content":[
		{"type":"text","text":"func main() {}"}
	]}]}`
	doc, _ := FromProseMirror([]byte(pmJSON), nil)
	defer doc.Close()
	for p := range doc.Body().Paragraphs() {
		for r := range p.Runs() {
			if r.Text() == "func main() {}" && r.FontName() != nil && *r.FontName() == "Courier New" {
				return
			}
		}
	}
	t.Error("code block not found")
}

func TestFromProseMirror_PageBreak(t *testing.T) {
	pmJSON := `{"type":"doc","content":[
		{"type":"paragraph","content":[{"type":"text","text":"before"}]},
		{"type":"pageBreak"},
		{"type":"paragraph","content":[{"type":"text","text":"after"}]}
	]}`
	doc, _ := FromProseMirror([]byte(pmJSON), nil)
	defer doc.Close()
	text := doc.Text(nil)
	if !strings.Contains(text, "before") || !strings.Contains(text, "after") {
		t.Errorf("text = %q", text)
	}
}

func TestFromProseMirror_HardBreak(t *testing.T) {
	pmJSON := `{"type":"doc","content":[{"type":"paragraph","content":[
		{"type":"text","text":"line1"},
		{"type":"hardBreak"},
		{"type":"text","text":"line2"}
	]}]}`
	doc, _ := FromProseMirror([]byte(pmJSON), nil)
	defer doc.Close()
	text := doc.Text(nil)
	if !strings.Contains(text, "line1") || !strings.Contains(text, "line2") {
		t.Errorf("text = %q", text)
	}
}

func TestFromProseMirror_Underline(t *testing.T) {
	pmJSON := `{"type":"doc","content":[{"type":"paragraph","content":[
		{"type":"text","text":"underlined","marks":[{"type":"underline"}]}
	]}]}`
	doc, _ := FromProseMirror([]byte(pmJSON), nil)
	defer doc.Close()
	for p := range doc.Body().Paragraphs() {
		for r := range p.Runs() {
			if r.Text() == "underlined" {
				if r.Underline() == nil || !*r.Underline() {
					t.Error("underline not set")
				}
				return
			}
		}
	}
	t.Error("underlined run not found")
}

func TestFromProseMirror_Strike(t *testing.T) {
	pmJSON := `{"type":"doc","content":[{"type":"paragraph","content":[
		{"type":"text","text":"struck","marks":[{"type":"strike"}]}
	]}]}`
	doc, _ := FromProseMirror([]byte(pmJSON), nil)
	defer doc.Close()
	for p := range doc.Body().Paragraphs() {
		for r := range p.Runs() {
			if r.Text() == "struck" {
				if r.Strikethrough() == nil || !*r.Strikethrough() {
					t.Error("strike not set")
				}
				return
			}
		}
	}
	t.Error("struck run not found")
}

func TestFromProseMirror_Code(t *testing.T) {
	pmJSON := `{"type":"doc","content":[{"type":"paragraph","content":[
		{"type":"text","text":"code","marks":[{"type":"code"}]}
	]}]}`
	doc, _ := FromProseMirror([]byte(pmJSON), nil)
	defer doc.Close()
	for p := range doc.Body().Paragraphs() {
		for r := range p.Runs() {
			if r.Text() == "code" && r.FontName() != nil && *r.FontName() == "Courier New" {
				return
			}
		}
	}
	t.Error("code run not found")
}

func TestFromProseMirror_Alignment(t *testing.T) {
	pmJSON := `{"type":"doc","content":[{"type":"paragraph","attrs":{"textAlign":"center"},"content":[
		{"type":"text","text":"centered"}
	]}]}`
	doc, _ := FromProseMirror([]byte(pmJSON), nil)
	defer doc.Close()
	// Collect paragraphs first to avoid lock re-entrancy (iterator holds RLock, Format needs Lock).
	var paras []*Paragraph
	for p := range doc.Body().Paragraphs() {
		paras = append(paras, p)
	}
	for _, p := range paras {
		if p.Format().Alignment() == "center" {
			return
		}
	}
	t.Error("alignment not set")
}

func TestFromProseMirror_DocProperties(t *testing.T) {
	pmJSON := `{"type":"doc","attrs":{"title":"My Doc","author":"Bob"},"content":[
		{"type":"paragraph","content":[{"type":"text","text":"text"}]}
	]}`
	doc, _ := FromProseMirror([]byte(pmJSON), nil)
	defer doc.Close()
	if doc.Properties().Title() != "My Doc" {
		t.Errorf("title = %q", doc.Properties().Title())
	}
	if doc.Properties().Author() != "Bob" {
		t.Errorf("author = %q", doc.Properties().Author())
	}
}

func TestFromProseMirror_VanillaListNames(t *testing.T) {
	pmJSON := `{"type":"doc","content":[{"type":"bullet_list","content":[
		{"type":"list_item","content":[{"type":"paragraph","content":[{"type":"text","text":"item"}]}]}
	]}]}`
	doc, _ := FromProseMirror([]byte(pmJSON), nil)
	defer doc.Close()
	if !strings.Contains(doc.Text(nil), "item") {
		t.Error("vanilla list names not handled")
	}
}

func TestImportProseMirror_AppendToExisting(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	doc.Body().AddParagraph().AddRun("existing")

	pmJSON := `{"type":"doc","content":[{"type":"paragraph","content":[{"type":"text","text":"appended"}]}]}`
	if err := doc.ImportProseMirror([]byte(pmJSON)); err != nil {
		t.Fatal(err)
	}
	text := doc.Text(nil)
	if !strings.Contains(text, "existing") {
		t.Error("lost existing")
	}
	if !strings.Contains(text, "appended") {
		t.Error("missing appended")
	}
}

func TestParagraph_SetProseMirror(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	p := doc.Body().AddParagraph()
	pmJSON := `{"type":"paragraph","content":[
		{"type":"text","text":"bold","marks":[{"type":"bold"}]},
		{"type":"text","text":" text"}
	]}`
	if err := p.SetProseMirror([]byte(pmJSON)); err != nil {
		t.Fatal(err)
	}
	if p.Text() != "bold text" {
		t.Errorf("text = %q", p.Text())
	}
	for r := range p.Runs() {
		if r.Text() == "bold" && (r.Bold() == nil || !*r.Bold()) {
			t.Error("bold not set")
		}
	}
}

func TestCell_SetProseMirror(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	tbl := doc.Body().AddTable(1, 1)
	c := tbl.Cell(0, 0)
	pmJSON := `{"type":"tableCell","content":[
		{"type":"paragraph","content":[{"type":"text","text":"cell content"}]}
	]}`
	if err := c.SetProseMirror([]byte(pmJSON)); err != nil {
		t.Fatal(err)
	}
	if c.Text() != "cell content" {
		t.Errorf("text = %q", c.Text())
	}
}

func ExampleDocument_ImportProseMirror() {
	doc, _ := New(nil)
	defer doc.Close()
	doc.Body().AddParagraph().AddRun("existing")
	pmJSON := []byte(`{"type":"doc","content":[{"type":"paragraph","content":[{"type":"text","text":"new"}]}]}`)
	_ = doc.ImportProseMirror(pmJSON)
}

func ExampleParagraph_SetProseMirror() {
	doc, _ := New(nil)
	defer doc.Close()
	p := doc.Body().AddParagraph()
	pmJSON := []byte(`{"type":"paragraph","content":[{"type":"text","text":"hello","marks":[{"type":"bold"}]}]}`)
	_ = p.SetProseMirror(pmJSON)
}
