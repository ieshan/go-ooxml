package docx

import (
	"regexp"
	"strings"
	"testing"
)

func TestDocument_Find(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	doc.Body().AddParagraph().AddRun("Hello World")
	doc.Body().AddParagraph().AddRun("Hello Again")

	results := doc.Find("Hello")
	if len(results) != 2 {
		t.Errorf("count = %d, want 2", len(results))
	}
}

func TestDocument_Find_AcrossRuns(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	p := doc.Body().AddParagraph()
	p.AddRun("Hel")
	p.AddRun("lo Wor")
	p.AddRun("ld")

	results := doc.Find("Hello World")
	if len(results) != 1 {
		t.Errorf("count = %d, want 1", len(results))
	}
	if len(results) > 0 && len(results[0].Runs) != 3 {
		t.Errorf("runs = %d, want 3", len(results[0].Runs))
	}
}

func TestDocument_Find_NotFound(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	doc.Body().AddParagraph().AddRun("Hello")
	if len(doc.Find("xyz")) != 0 {
		t.Error("should be empty")
	}
}

func TestDocument_FindRegex(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	doc.Body().AddParagraph().AddRun("Error code: 404")
	doc.Body().AddParagraph().AddRun("Error code: 500")

	results := doc.FindRegex(regexp.MustCompile(`\d{3}`))
	if len(results) != 2 {
		t.Errorf("count = %d", len(results))
	}
}

func TestSearch_First(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	doc.Body().AddParagraph().AddRun("Hello")
	doc.Body().AddParagraph().AddRun("Hello")

	results := doc.Search("Hello").First().Results()
	if len(results) != 1 {
		t.Errorf("count = %d, want 1", len(results))
	}
}

func TestSearch_InStyle(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	p1 := doc.Body().AddParagraph()
	p1.SetStyle("Heading1")
	p1.AddRun("Title")
	p2 := doc.Body().AddParagraph()
	p2.AddRun("Title") // same text, different style

	results := doc.Search("Title").InStyle("Heading1").Results()
	if len(results) != 1 {
		t.Errorf("count = %d, want 1", len(results))
	}
}

func TestSearch_SetBold(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	doc.Body().AddParagraph().AddRun("important")
	doc.Body().AddParagraph().AddRun("important")

	bold := true
	count := doc.Search("important").SetBold(&bold)
	if count != 2 {
		t.Errorf("count = %d, want 2", count)
	}
}

func TestSearch_ReplaceText(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	doc.Body().AddParagraph().AddRun("draft")
	doc.Body().AddParagraph().AddRun("draft")

	count := doc.Search("draft").ReplaceText("final")
	if count != 2 {
		t.Errorf("count = %d", count)
	}
	if doc.Text(nil) != "final\nfinal" {
		t.Errorf("text = %q", doc.Text(nil))
	}
}

func TestSearch_AddComment(t *testing.T) {
	doc, _ := New(&Config{Author: "Bot"})
	defer doc.Close()
	doc.Body().AddParagraph().AddRun("TODO fix this")

	comments := doc.Search("TODO").AddComment("Please resolve")
	if len(comments) != 1 {
		t.Errorf("comments = %d", len(comments))
	}
}

func TestSearchResult_Text(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	doc.Body().AddParagraph().AddRun("Hello World")
	results := doc.Find("World")
	if len(results) != 1 {
		t.Fatal("no results")
	}
	if results[0].Text() != "World" {
		t.Errorf("text = %q", results[0].Text())
	}
}

func TestSearch_CaseSensitive(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	doc.Body().AddParagraph().AddRun("Hello World")
	doc.Body().AddParagraph().AddRun("hello world")

	// Case-insensitive (default): should find both.
	results := doc.Search("hello").Results()
	if len(results) != 2 {
		t.Errorf("case-insensitive count = %d, want 2", len(results))
	}

	// Case-sensitive: should find only the lowercase one.
	results = doc.Search("hello").CaseSensitive(true).Results()
	if len(results) != 1 {
		t.Errorf("case-sensitive count = %d, want 1", len(results))
	}
}

func TestSearch_Nth(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	doc.Body().AddParagraph().AddRun("foo")
	doc.Body().AddParagraph().AddRun("foo")
	doc.Body().AddParagraph().AddRun("foo")

	results := doc.Search("foo").Nth(1).Results()
	if len(results) != 1 {
		t.Errorf("count = %d, want 1", len(results))
	}
}

func TestSearch_ReplaceText_AcrossRuns(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	p := doc.Body().AddParagraph()
	p.AddRun("foo")
	p.AddRun("bar")

	count := doc.Search("foobar").ReplaceText("baz")
	if count != 1 {
		t.Errorf("count = %d, want 1", count)
	}
	if doc.Text(nil) != "baz" {
		t.Errorf("text = %q, want %q", doc.Text(nil), "baz")
	}
}

func TestSearch_SetItalic(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	doc.Body().AddParagraph().AddRun("important")

	italic := true
	count := doc.Search("important").SetItalic(&italic)
	if count != 1 {
		t.Errorf("count = %d, want 1", count)
	}
}

func TestSearch_SetStyle(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	doc.Body().AddParagraph().AddRun("Title Text")
	doc.Body().AddParagraph().AddRun("Title Text")

	count := doc.Search("Title").SetStyle("Heading1")
	if count != 2 {
		t.Errorf("count = %d, want 2", count)
	}
}

func ExampleDocument_Search() {
	doc, _ := New(nil)
	defer doc.Close()
	doc.Body().AddParagraph().AddRun("Hello World")
	bold := true
	doc.Search("World").SetBold(&bold)
}

func ExampleDocument_Find() {
	doc, _ := New(nil)
	defer doc.Close()
	doc.Body().AddParagraph().AddRun("Contact support@example.com for help.")
	results := doc.Find("support")
	for _, r := range results {
		_ = r.Text() // "support"
	}
}

func ExampleDocument_FindRegex() {
	doc, _ := New(nil)
	defer doc.Close()
	doc.Body().AddParagraph().AddRun("Call us at 555-1234 or 555-5678.")
	pattern := regexp.MustCompile(`\d{3}-\d{4}`)
	results := doc.FindRegex(pattern)
	for _, r := range results {
		_ = r.Text() // "555-1234", "555-5678"
	}
}

func ExampleSearchResult_Text() {
	doc, _ := New(nil)
	defer doc.Close()
	doc.Body().AddParagraph().AddRun("Hello World")
	results := doc.Find("World")
	if len(results) > 0 {
		_ = results[0].Text() // "World"
	}
}

func ExampleSearchQuery_ReplaceTextFormatted() {
	doc, _ := New(nil)
	defer doc.Close()
	p := doc.Body().AddParagraph()
	r := p.AddRun("ACME Widget")
	b := true
	r.SetBold(&b)

	// Replace text while preserving bold formatting.
	doc.Search("ACME Widget").ReplaceTextFormatted("SuperWidget Pro")
	// "SuperWidget Pro" is now bold.
}

func ExampleSearchQuery_ReplaceMarkdown() {
	doc, _ := New(nil)
	defer doc.Close()
	p := doc.Body().AddParagraph()
	r := p.AddRun("old product")
	it := true
	r.SetItalic(&it)

	// Replace with markdown — inherited italic + markdown bold.
	doc.Search("old product").ReplaceMarkdown("**new** product")
	// "new" is bold+italic, " product" is italic.
}

func TestFind_InsideTableCell(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	// Add text only inside a table cell, not in a top-level paragraph.
	tbl := doc.Body().AddTable(1, 1)
	tbl.Cell(0, 0).AddParagraph().AddRun("hidden in table")

	results := doc.Find("hidden in table")
	if len(results) != 1 {
		t.Errorf("Find inside table cell: got %d results, want 1", len(results))
	}
}

func TestFindRegex_InsideTableCell(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	tbl := doc.Body().AddTable(1, 1)
	tbl.Cell(0, 0).AddParagraph().AddRun("item-42")

	pattern := regexp.MustCompile(`item-\d+`)
	results := doc.FindRegex(pattern)
	if len(results) != 1 {
		t.Errorf("FindRegex inside table cell: got %d results, want 1", len(results))
	}
}

// ---------------------------------------------------------------------------
// SearchResult.Markdown
// ---------------------------------------------------------------------------

// ---------------------------------------------------------------------------
// ReplaceMarkdown
// ---------------------------------------------------------------------------

func TestSearch_ReplaceMarkdown_BoldOnItalicBase(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	p := doc.Body().AddParagraph()
	r := p.AddRun("old text")
	it := true
	r.SetItalic(&it)

	count := doc.Search("old text").ReplaceMarkdown("**strong** point")
	if count != 1 {
		t.Errorf("count = %d, want 1", count)
	}

	for p := range doc.Body().Paragraphs() {
		for run := range p.Runs() {
			switch run.Text() {
			case "strong":
				if run.Bold() == nil || !*run.Bold() {
					t.Error("'strong' should be bold from markdown")
				}
				if run.Italic() == nil || !*run.Italic() {
					t.Error("'strong' should be italic from inheritance")
				}
			case " point":
				if run.Bold() != nil && *run.Bold() {
					t.Error("' point' should NOT be bold")
				}
				if run.Italic() == nil || !*run.Italic() {
					t.Error("' point' should be italic from inheritance")
				}
			}
		}
	}
}

func TestSearch_ReplaceMarkdown_PlainBase(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	doc.Body().AddParagraph().AddRun("draft text")

	count := doc.Search("draft text").ReplaceMarkdown("**bold** and *italic*")
	if count != 1 {
		t.Errorf("count = %d, want 1", count)
	}

	for p := range doc.Body().Paragraphs() {
		for run := range p.Runs() {
			switch run.Text() {
			case "bold":
				if run.Bold() == nil || !*run.Bold() {
					t.Error("'bold' should be bold")
				}
			case "italic":
				if run.Italic() == nil || !*run.Italic() {
					t.Error("'italic' should be italic")
				}
			case " and ":
				if run.Bold() != nil && *run.Bold() {
					t.Error("' and ' should not be bold")
				}
			}
		}
	}
}

func TestSearch_ReplaceMarkdown_InheritsFont(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	p := doc.Body().AddParagraph()
	r := p.AddRun("old")
	r.SetFontName("Arial")

	doc.Search("old").ReplaceMarkdown("**new**")

	for p := range doc.Body().Paragraphs() {
		for run := range p.Runs() {
			if run.Text() == "new" {
				if run.Bold() == nil || !*run.Bold() {
					t.Error("should be bold from markdown")
				}
				if run.FontName() == nil || *run.FontName() != "Arial" {
					t.Error("should inherit Arial font")
				}
				return
			}
		}
	}
	t.Error("replacement run not found")
}

func TestSearch_ReplaceMarkdown_PartialSingleRun(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	p := doc.Body().AddParagraph()
	r := p.AddRun("prefix TARGET suffix")
	b := true
	r.SetBold(&b)

	doc.Search("TARGET").ReplaceMarkdown("*new*")

	text := doc.Text(nil)
	if !strings.Contains(text, "prefix") || !strings.Contains(text, "new") || !strings.Contains(text, "suffix") {
		t.Errorf("text = %q, expected prefix/new/suffix", text)
	}

	for p := range doc.Body().Paragraphs() {
		for run := range p.Runs() {
			if run.Text() == "new" {
				if run.Italic() == nil || !*run.Italic() {
					t.Error("'new' should be italic from markdown")
				}
				if run.Bold() == nil || !*run.Bold() {
					t.Error("'new' should inherit bold")
				}
			}
		}
	}
}

// ---------------------------------------------------------------------------
// ReplaceTextFormatted
// ---------------------------------------------------------------------------

func TestSearch_ReplaceTextFormatted_PreservesBold(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	p := doc.Body().AddParagraph()
	r := p.AddRun("old text")
	b := true
	r.SetBold(&b)

	count := doc.Search("old text").ReplaceTextFormatted("new text")
	if count != 1 {
		t.Errorf("count = %d, want 1", count)
	}
	if doc.Text(nil) != "new text" {
		t.Errorf("text = %q", doc.Text(nil))
	}
	for p := range doc.Body().Paragraphs() {
		for run := range p.Runs() {
			if run.Text() == "new text" {
				if run.Bold() == nil || !*run.Bold() {
					t.Error("bold formatting was not preserved")
				}
				return
			}
		}
	}
	t.Error("replacement run not found")
}

func TestSearch_ReplaceTextFormatted_UnionAcrossRuns(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	p := doc.Body().AddParagraph()
	r1 := p.AddRun("bold ")
	b := true
	r1.SetBold(&b)
	r2 := p.AddRun("italic")
	it := true
	r2.SetItalic(&it)

	count := doc.Search("bold italic").ReplaceTextFormatted("merged")
	if count != 1 {
		t.Errorf("count = %d, want 1", count)
	}
	for p := range doc.Body().Paragraphs() {
		for run := range p.Runs() {
			if run.Text() == "merged" {
				if run.Bold() == nil || !*run.Bold() {
					t.Error("bold not preserved from union")
				}
				if run.Italic() == nil || !*run.Italic() {
					t.Error("italic not preserved from union")
				}
				return
			}
		}
	}
	t.Error("replacement run not found")
}

func TestSearch_ReplaceTextFormatted_SingleRunPartial(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	p := doc.Body().AddParagraph()
	r := p.AddRun("Hello World Goodbye")
	b := true
	r.SetBold(&b)

	doc.Search("World").ReplaceTextFormatted("Earth")
	text := doc.Text(nil)
	if text != "Hello Earth Goodbye" {
		t.Errorf("text = %q", text)
	}
	for p := range doc.Body().Paragraphs() {
		for run := range p.Runs() {
			if strings.Contains(run.Text(), "Earth") {
				if run.Bold() == nil || !*run.Bold() {
					t.Error("bold lost on partial replace")
				}
			}
		}
	}
}

func TestSearch_ReplaceTextFormatted_PlainRuns(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	doc.Body().AddParagraph().AddRun("draft")

	count := doc.Search("draft").ReplaceTextFormatted("final")
	if count != 1 {
		t.Errorf("count = %d", count)
	}
	if doc.Text(nil) != "final" {
		t.Errorf("text = %q", doc.Text(nil))
	}
}

// ---------------------------------------------------------------------------
// mergedRunProperties
// ---------------------------------------------------------------------------

func TestMergedRunProperties_UnionBoldItalic(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	p := doc.Body().AddParagraph()
	r1 := p.AddRun("bold ")
	b := true
	r1.SetBold(&b)
	r2 := p.AddRun("italic")
	it := true
	r2.SetItalic(&it)

	results := doc.Find("bold italic")
	if len(results) != 1 {
		t.Fatal("no results")
	}
	rpr := mergedRunProperties(results[0].Runs)
	if rpr == nil {
		t.Fatal("nil rpr")
	}
	if rpr.Bold == nil || !*rpr.Bold {
		t.Error("expected bold from union")
	}
	if rpr.Italic == nil || !*rpr.Italic {
		t.Error("expected italic from union")
	}
}

func TestMergedRunProperties_FontFromFirst(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	p := doc.Body().AddParagraph()
	r1 := p.AddRun("styled ")
	r1.SetFontName("Arial")
	r1.SetFontSize(14.0)
	p.AddRun("plain")

	results := doc.Find("styled plain")
	if len(results) != 1 {
		t.Fatal("no results")
	}
	rpr := mergedRunProperties(results[0].Runs)
	if rpr.FontName == nil || *rpr.FontName != "Arial" {
		t.Error("expected FontName from first run")
	}
	if rpr.FontSize == nil || *rpr.FontSize != "28" {
		t.Errorf("expected FontSize 28 (14pt * 2 half-pts), got %v", rpr.FontSize)
	}
}

func TestMergedRunProperties_NilForPlainRuns(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	doc.Body().AddParagraph().AddRun("plain text")

	results := doc.Find("plain text")
	if len(results) != 1 {
		t.Fatal("no results")
	}
	rpr := mergedRunProperties(results[0].Runs)
	if rpr != nil {
		t.Error("expected nil rpr for runs with no formatting")
	}
}

func TestSearchResult_Markdown_FormattedRuns(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	p := doc.Body().AddParagraph()
	r := p.AddRun("bold text")
	b := true
	r.SetBold(&b)

	results := doc.Find("bold text")
	if len(results) != 1 {
		t.Fatal("no results")
	}
	md := results[0].Markdown()
	if md != "**bold text**" {
		t.Errorf("Markdown() = %q, want %q", md, "**bold text**")
	}
}

func TestSearchResult_Markdown_MixedRuns(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	p := doc.Body().AddParagraph()
	p.AddRun("Hello ")
	r := p.AddRun("World")
	b := true
	r.SetBold(&b)

	results := doc.Find("Hello World")
	if len(results) != 1 {
		t.Fatal("no results")
	}
	md := results[0].Markdown()
	if md != "Hello **World**" {
		t.Errorf("Markdown() = %q, want %q", md, "Hello **World**")
	}
}
