package docx

import (
	"strings"
	"testing"
)

func TestRun_Markdown_Bold(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	r := doc.Body().AddParagraph().AddRun("bold")
	bold := true
	r.SetBold(&bold)
	if r.Markdown() != "**bold**" {
		t.Errorf("md = %q", r.Markdown())
	}
}

func TestRun_Markdown_Italic(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	r := doc.Body().AddParagraph().AddRun("italic")
	it := true
	r.SetItalic(&it)
	if r.Markdown() != "*italic*" {
		t.Errorf("md = %q", r.Markdown())
	}
}

func TestRun_Markdown_BoldItalic(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	r := doc.Body().AddParagraph().AddRun("both")
	b := true
	r.SetBold(&b)
	r.SetItalic(&b)
	if r.Markdown() != "***both***" {
		t.Errorf("md = %q", r.Markdown())
	}
}

func TestRun_Markdown_Strikethrough(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	r := doc.Body().AddParagraph().AddRun("struck")
	s := true
	r.SetStrikethrough(&s)
	if r.Markdown() != "~~struck~~" {
		t.Errorf("md = %q", r.Markdown())
	}
}

func TestRun_Markdown_Plain(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	r := doc.Body().AddParagraph().AddRun("plain")
	if r.Markdown() != "plain" {
		t.Errorf("md = %q", r.Markdown())
	}
}

func TestRun_Markdown_CodeFont(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	r := doc.Body().AddParagraph().AddRun("code")
	r.SetFontName("Courier New")
	if r.Markdown() != "`code`" {
		t.Errorf("md = %q", r.Markdown())
	}
}

func TestRun_Markdown_EmptyText(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	r := doc.Body().AddParagraph().AddRun("")
	b := true
	r.SetBold(&b)
	if r.Markdown() != "" {
		t.Errorf("md = %q, want empty", r.Markdown())
	}
}

func TestParagraph_Markdown_Heading(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	p := doc.Body().AddParagraph()
	p.SetStyle("Heading1")
	p.AddRun("Title")
	if p.Markdown() != "# Title" {
		t.Errorf("md = %q", p.Markdown())
	}
}

func TestParagraph_Markdown_Heading2(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	p := doc.Body().AddParagraph()
	p.SetStyle("Heading2")
	p.AddRun("Subtitle")
	if p.Markdown() != "## Subtitle" {
		t.Errorf("md = %q", p.Markdown())
	}
}

func TestParagraph_Markdown_Title(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	p := doc.Body().AddParagraph()
	p.SetStyle("Title")
	p.AddRun("Main")
	if p.Markdown() != "# Main" {
		t.Errorf("md = %q", p.Markdown())
	}
}

func TestParagraph_Markdown_Mixed(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	p := doc.Body().AddParagraph()
	p.AddRun("Hello ")
	r := p.AddRun("World")
	b := true
	r.SetBold(&b)
	if p.Markdown() != "Hello **World**" {
		t.Errorf("md = %q", p.Markdown())
	}
}

func TestParagraph_Markdown_Blockquote(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	p := doc.Body().AddParagraph()
	p.SetStyle("Quote")
	p.AddRun("quoted text")
	if p.Markdown() != "> quoted text" {
		t.Errorf("md = %q", p.Markdown())
	}
}

func TestTable_Markdown(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	tbl := doc.Body().AddTable(2, 2)
	tbl.Cell(0, 0).AddParagraph().AddRun("H1")
	tbl.Cell(0, 1).AddParagraph().AddRun("H2")
	tbl.Cell(1, 0).AddParagraph().AddRun("A")
	tbl.Cell(1, 1).AddParagraph().AddRun("B")
	md := tbl.Markdown()
	if !strings.Contains(md, "H1") || !strings.Contains(md, "H2") {
		t.Errorf("header: %q", md)
	}
	if !strings.Contains(md, "A") || !strings.Contains(md, "B") {
		t.Errorf("body: %q", md)
	}
	if !strings.Contains(md, "---") {
		t.Errorf("separator: %q", md)
	}
	// Verify pipe table structure
	lines := strings.Split(md, "\n")
	if len(lines) < 3 {
		t.Errorf("expected at least 3 lines, got %d: %q", len(lines), md)
	}
	for _, line := range lines {
		if !strings.HasPrefix(line, "|") || !strings.HasSuffix(line, "|") {
			t.Errorf("line not pipe-delimited: %q", line)
		}
	}
}

func TestTable_Markdown_SingleRow(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	tbl := doc.Body().AddTable(1, 2)
	tbl.Cell(0, 0).AddParagraph().AddRun("A")
	tbl.Cell(0, 1).AddParagraph().AddRun("B")
	md := tbl.Markdown()
	if !strings.Contains(md, "A") || !strings.Contains(md, "B") {
		t.Errorf("md = %q", md)
	}
	if !strings.Contains(md, "---") {
		t.Errorf("separator missing: %q", md)
	}
}

func TestCell_Markdown(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	tbl := doc.Body().AddTable(1, 1)
	c := tbl.Cell(0, 0)
	c.AddParagraph().AddRun("hello")
	if c.Markdown() != "hello" {
		t.Errorf("md = %q", c.Markdown())
	}
}

func TestBody_Markdown(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	p := doc.Body().AddParagraph()
	p.SetStyle("Heading1")
	p.AddRun("Title")
	doc.Body().AddParagraph().AddRun("Body text")
	md := doc.Body().Markdown()
	if !strings.Contains(md, "# Title") {
		t.Error("missing heading")
	}
	if !strings.Contains(md, "Body text") {
		t.Error("missing body")
	}
}

func TestBody_Markdown_WithTable(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	doc.Body().AddParagraph().AddRun("Before")
	tbl := doc.Body().AddTable(1, 2)
	tbl.Cell(0, 0).AddParagraph().AddRun("A")
	tbl.Cell(0, 1).AddParagraph().AddRun("B")
	doc.Body().AddParagraph().AddRun("After")
	md := doc.Body().Markdown()
	if !strings.Contains(md, "Before") {
		t.Error("missing before")
	}
	if !strings.Contains(md, "A") || !strings.Contains(md, "B") || !strings.Contains(md, "|") {
		t.Error("missing table")
	}
	if !strings.Contains(md, "After") {
		t.Error("missing after")
	}
}

func TestDocument_Markdown_Default(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	p := doc.Body().AddParagraph()
	p.SetStyle("Heading1")
	p.AddRun("Report")
	doc.Body().AddParagraph().AddRun("Content here")
	md := doc.Markdown(nil)
	if !strings.Contains(md, "# Report") {
		t.Error("missing heading")
	}
	if !strings.Contains(md, "Content here") {
		t.Error("missing content")
	}
}

func TestDocument_Markdown_WithComments(t *testing.T) {
	doc, _ := New(&Config{Author: "Bot"})
	defer doc.Close()
	r := doc.Body().AddParagraph().AddRun("text")
	doc.AddComment(r, r, "review this")

	md := doc.Markdown(nil)
	if strings.Contains(md, "review") {
		t.Error("should not include comments by default")
	}

	md = doc.Markdown(&MarkdownOptions{IncludeComments: true})
	if !strings.Contains(md, "review this") {
		t.Error("should include comments")
	}
}

func TestDocument_Markdown_HorizontalRule(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	doc.Body().AddParagraph().AddRun("text")
	opts := &MarkdownOptions{HorizontalRule: "***"}
	md := doc.Markdown(opts)
	// Just ensure it doesn't panic; HR is only used if there's a separator in the doc
	if md == "" {
		t.Error("should not be empty")
	}
}

func ExampleDocument_Markdown() {
	doc, _ := New(nil)
	defer doc.Close()
	p := doc.Body().AddParagraph()
	p.SetStyle("Heading1")
	p.AddRun("Report")
	doc.Body().AddParagraph().AddRun("Content here")
	_ = doc.Markdown(nil) // "# Report\n\nContent here"
}

// ---------------------------------------------------------------------------
// Markdown → DOCX import tests
// ---------------------------------------------------------------------------

func TestFromMarkdown_Heading(t *testing.T) {
	doc, err := FromMarkdown("# Hello World", nil)
	if err != nil {
		t.Fatalf("FromMarkdown: %v", err)
	}
	defer doc.Close()
	count := 0
	for p := range doc.Body().Paragraphs() {
		if p.Style() == "Heading1" && p.Text() == "Hello World" {
			count++
		}
	}
	if count != 1 {
		t.Errorf("heading count = %d", count)
	}
}

func TestFromMarkdown_MultiLevel(t *testing.T) {
	md := "# H1\n\n## H2\n\n### H3\n\nBody text"
	doc, err := FromMarkdown(md, nil)
	if err != nil {
		t.Fatalf("FromMarkdown: %v", err)
	}
	defer doc.Close()
	styles := []string{}
	for p := range doc.Body().Paragraphs() {
		styles = append(styles, p.Style())
	}
	if len(styles) != 4 {
		t.Errorf("count = %d, styles = %v", len(styles), styles)
	}
	if len(styles) > 0 && styles[0] != "Heading1" {
		t.Errorf("s[0] = %q", styles[0])
	}
	if len(styles) > 1 && styles[1] != "Heading2" {
		t.Errorf("s[1] = %q", styles[1])
	}
	if len(styles) > 2 && styles[2] != "Heading3" {
		t.Errorf("s[2] = %q", styles[2])
	}
}

func TestFromMarkdown_Bold(t *testing.T) {
	doc, err := FromMarkdown("This is **bold** text", nil)
	if err != nil {
		t.Fatalf("FromMarkdown: %v", err)
	}
	defer doc.Close()
	for p := range doc.Body().Paragraphs() {
		for r := range p.Runs() {
			if r.Text() == "bold" {
				if r.Bold() == nil || !*r.Bold() {
					t.Error("not bold")
				}
				return
			}
		}
	}
	t.Error("bold run not found")
}

func TestFromMarkdown_Italic(t *testing.T) {
	doc, err := FromMarkdown("This is *italic* text", nil)
	if err != nil {
		t.Fatalf("FromMarkdown: %v", err)
	}
	defer doc.Close()
	for p := range doc.Body().Paragraphs() {
		for r := range p.Runs() {
			if r.Text() == "italic" {
				if r.Italic() == nil || !*r.Italic() {
					t.Error("not italic")
				}
				return
			}
		}
	}
	t.Error("italic run not found")
}

func TestFromMarkdown_Code(t *testing.T) {
	doc, err := FromMarkdown("Use `fmt.Println` here", nil)
	if err != nil {
		t.Fatalf("FromMarkdown: %v", err)
	}
	defer doc.Close()
	for p := range doc.Body().Paragraphs() {
		for r := range p.Runs() {
			if r.Text() == "fmt.Println" {
				if r.FontName() == nil || *r.FontName() != "Courier New" {
					t.Error("not monospace")
				}
				return
			}
		}
	}
	t.Error("code run not found")
}

func TestFromMarkdown_Table(t *testing.T) {
	md := "| A | B |\n|---|---|\n| 1 | 2 |"
	doc, err := FromMarkdown(md, nil)
	if err != nil {
		t.Fatalf("FromMarkdown: %v", err)
	}
	defer doc.Close()
	count := 0
	for range doc.Body().Tables() {
		count++
	}
	if count != 1 {
		t.Errorf("tables = %d", count)
	}
}

func TestFromMarkdown_HorizontalRule(t *testing.T) {
	md := "Before\n\n---\n\nAfter"
	doc, err := FromMarkdown(md, nil)
	if err != nil {
		t.Fatalf("FromMarkdown: %v", err)
	}
	defer doc.Close()
	text := doc.Text(nil)
	if !strings.Contains(text, "Before") || !strings.Contains(text, "After") {
		t.Errorf("text = %q", text)
	}
}

func TestParagraph_SetMarkdown(t *testing.T) {
	doc, err := New(nil)
	if err != nil {
		t.Fatalf("New: %v", err)
	}
	defer doc.Close()
	p := doc.Body().AddParagraph()
	if err := p.SetMarkdown("This has **bold** and *italic*"); err != nil {
		t.Fatalf("SetMarkdown: %v", err)
	}
	text := p.Text()
	if text != "This has bold and italic" {
		t.Errorf("text = %q", text)
	}
	// Check formatting
	for r := range p.Runs() {
		if r.Text() == "bold" && (r.Bold() == nil || !*r.Bold()) {
			t.Error("not bold")
		}
		if r.Text() == "italic" && (r.Italic() == nil || !*r.Italic()) {
			t.Error("not italic")
		}
	}
}

func TestImportMarkdown(t *testing.T) {
	doc, err := New(nil)
	if err != nil {
		t.Fatalf("New: %v", err)
	}
	defer doc.Close()
	doc.Body().AddParagraph().AddRun("existing")
	if err := doc.ImportMarkdown("# New Section\n\nNew content"); err != nil {
		t.Fatalf("ImportMarkdown: %v", err)
	}
	text := doc.Text(nil)
	if !strings.Contains(text, "existing") {
		t.Error("lost existing")
	}
	if !strings.Contains(text, "New Section") {
		t.Error("missing imported")
	}
	if !strings.Contains(text, "New content") {
		t.Error("missing content")
	}
}

func TestFromMarkdown_BulletList(t *testing.T) {
	md := "- Item 1\n- Item 2\n- Item 3"
	doc, err := FromMarkdown(md, nil)
	if err != nil {
		t.Fatalf("FromMarkdown: %v", err)
	}
	defer doc.Close()
	text := doc.Text(nil)
	if !strings.Contains(text, "Item 1") {
		t.Error("missing item")
	}
}

func TestFromMarkdown_Strikethrough(t *testing.T) {
	doc, err := FromMarkdown("Text with ~~strike~~ here", nil)
	if err != nil {
		t.Fatalf("FromMarkdown: %v", err)
	}
	defer doc.Close()
	for p := range doc.Body().Paragraphs() {
		for r := range p.Runs() {
			if r.Text() == "strike" {
				if r.Strikethrough() == nil || !*r.Strikethrough() {
					t.Error("not strikethrough")
				}
				return
			}
		}
	}
	t.Error("strikethrough run not found")
}

func TestFromMarkdown_BoldItalic(t *testing.T) {
	doc, err := FromMarkdown("***bold italic***", nil)
	if err != nil {
		t.Fatalf("FromMarkdown: %v", err)
	}
	defer doc.Close()
	for p := range doc.Body().Paragraphs() {
		for r := range p.Runs() {
			if r.Bold() != nil && *r.Bold() && r.Italic() != nil && *r.Italic() {
				return
			}
		}
	}
	t.Error("bold+italic run not found")
}

func TestFromMarkdown_CodeBlock(t *testing.T) {
	md := "Before\n```\nfunc foo() {}\n```\nAfter"
	doc, err := FromMarkdown(md, nil)
	if err != nil {
		t.Fatalf("FromMarkdown: %v", err)
	}
	defer doc.Close()
	found := false
	for p := range doc.Body().Paragraphs() {
		for r := range p.Runs() {
			if r.Text() == "func foo() {}" && r.FontName() != nil && *r.FontName() == "Courier New" {
				found = true
			}
		}
	}
	if !found {
		t.Error("code block paragraph not found")
	}
}

func TestFromMarkdown_Link(t *testing.T) {
	doc, err := FromMarkdown("See [Go website](https://go.dev) for details", nil)
	if err != nil {
		t.Fatalf("FromMarkdown: %v", err)
	}
	defer doc.Close()
	text := doc.Text(nil)
	if !strings.Contains(text, "Go website") {
		t.Errorf("link text not found in %q", text)
	}
}

func TestFromMarkdown_NumberedList(t *testing.T) {
	md := "1. First\n2. Second\n3. Third"
	doc, err := FromMarkdown(md, nil)
	if err != nil {
		t.Fatalf("FromMarkdown: %v", err)
	}
	defer doc.Close()
	text := doc.Text(nil)
	if !strings.Contains(text, "First") {
		t.Error("missing first item")
	}
	if !strings.Contains(text, "Second") {
		t.Error("missing second item")
	}
}

func TestFromMarkdown_BlockQuote(t *testing.T) {
	doc, err := FromMarkdown("> This is a quote", nil)
	if err != nil {
		t.Fatalf("FromMarkdown: %v", err)
	}
	defer doc.Close()
	for p := range doc.Body().Paragraphs() {
		if p.Style() == "Quote" && strings.Contains(p.Text(), "This is a quote") {
			return
		}
	}
	t.Error("blockquote paragraph not found")
}

func TestCell_SetMarkdown(t *testing.T) {
	doc, err := New(nil)
	if err != nil {
		t.Fatalf("New: %v", err)
	}
	defer doc.Close()
	tbl := doc.Body().AddTable(1, 1)
	c := tbl.Cell(0, 0)
	if err := c.SetMarkdown("**bold** text"); err != nil {
		t.Fatalf("SetMarkdown: %v", err)
	}
	text := c.Text()
	if !strings.Contains(text, "bold") {
		t.Errorf("cell text = %q", text)
	}
}

func ExampleFromMarkdown() {
	md := "# Report\n\nThis is **important** content."
	doc, _ := FromMarkdown(md, nil)
	defer doc.Close()
	_ = doc.Markdown(nil) // Should round-trip reasonably
}

func ExampleDocument_ImportMarkdown() {
	doc, _ := New(nil)
	defer doc.Close()
	// Append markdown to an existing document.
	_ = doc.ImportMarkdown("## Appendix\n\nSee also the references.")
	// doc.WriteTo(w)
}

// ---------------------------------------------------------------------------
// I6: Hyperlink import — [text](url) creates a real hyperlink element
// ---------------------------------------------------------------------------

func TestFromMarkdown_Link_CreatesHyperlink(t *testing.T) {
	doc, err := FromMarkdown("See [Go website](https://go.dev) for details", nil)
	if err != nil {
		t.Fatalf("FromMarkdown: %v", err)
	}
	defer doc.Close()

	var found bool
	for p := range doc.Body().Paragraphs() {
		for h := range p.Hyperlinks() {
			if h.URL() == "https://go.dev" && h.Text() == "Go website" {
				found = true
			}
		}
	}
	if !found {
		t.Error("hyperlink element not created for [text](url) markdown")
	}
}

func TestFromMarkdown_Link_TextPreserved(t *testing.T) {
	doc, err := FromMarkdown("Visit [Example](https://example.com) now", nil)
	if err != nil {
		t.Fatalf("FromMarkdown: %v", err)
	}
	defer doc.Close()
	text := doc.Text(nil)
	if !strings.Contains(text, "Example") {
		t.Errorf("link text not found in document text: %q", text)
	}
}

// ---------------------------------------------------------------------------
// I7: Hyperlink extraction — paragraphToMarkdown emits [text](url)
// ---------------------------------------------------------------------------

func TestParagraph_Markdown_Hyperlink(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	p := doc.Body().AddParagraph()
	p.AddRun("See ")
	p.AddHyperlink("https://example.com", "Example")
	p.AddRun(" for info")

	md := p.Markdown()
	if !strings.Contains(md, "[Example](https://example.com)") {
		t.Errorf("hyperlink not rendered as markdown link: %q", md)
	}
	if !strings.Contains(md, "See ") {
		t.Errorf("preceding text missing: %q", md)
	}
	if !strings.Contains(md, " for info") {
		t.Errorf("trailing text missing: %q", md)
	}
}

func TestDocument_Markdown_HyperlinkRoundTrip(t *testing.T) {
	// Import markdown with a link, then export — should produce [text](url).
	doc, err := FromMarkdown("Click [here](https://roundtrip.example.com) to continue", nil)
	if err != nil {
		t.Fatalf("FromMarkdown: %v", err)
	}
	defer doc.Close()

	md := doc.Markdown(nil)
	if !strings.Contains(md, "[here](https://roundtrip.example.com)") {
		t.Errorf("round-trip markdown missing link syntax: %q", md)
	}
}

func ExampleParagraph_SetMarkdown() {
	doc, _ := New(nil)
	defer doc.Close()
	p := doc.Body().AddParagraph()
	_ = p.SetMarkdown("This has **bold** and *italic*")
	// p.Text() == "This has bold and italic"
}

func ExampleParagraph_InsertMarkdownAfter() {
	doc, _ := New(nil)
	defer doc.Close()
	p := doc.Body().AddParagraph()
	p.AddRun("existing")
	_, _ = p.InsertMarkdownAfter("## New Section\n\nNew content")
}

func ExampleCell_SetMarkdown() {
	doc, _ := New(nil)
	defer doc.Close()
	tbl := doc.Body().AddTable(1, 1)
	c := tbl.Cell(0, 0)
	_ = c.SetMarkdown("**bold** cell")
}
