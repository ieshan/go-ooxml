package docx

import (
	"bytes"
	"strings"
	"testing"
)

func TestDocument_AddComment(t *testing.T) {
	doc, _ := New(&Config{Author: "TestBot"})
	defer doc.Close()
	p := doc.Body().AddParagraph()
	r := p.AddRun("Review this text")

	c := doc.AddComment(r, r, "Please fix this")
	if c == nil {
		t.Fatal("nil comment")
	}
	if c.Author() != "TestBot" {
		t.Errorf("author = %q", c.Author())
	}
	if c.Text() != "Please fix this" {
		t.Errorf("text = %q", c.Text())
	}
}

func TestDocument_AddComment_WithOptions(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	r := doc.Body().AddParagraph().AddRun("text")

	c := doc.AddComment(r, r, "comment", WithAuthor("Jane"), WithInitials("J"))
	if c.Author() != "Jane" {
		t.Errorf("author = %q", c.Author())
	}
	if c.Initials() != "J" {
		t.Errorf("initials = %q", c.Initials())
	}
}

func TestDocument_Comments_Iterator(t *testing.T) {
	doc, _ := New(&Config{Author: "Bot"})
	defer doc.Close()
	r1 := doc.Body().AddParagraph().AddRun("first")
	r2 := doc.Body().AddParagraph().AddRun("second")
	doc.AddComment(r1, r1, "comment 1")
	doc.AddComment(r2, r2, "comment 2")

	count := 0
	for range doc.Comments() {
		count++
	}
	if count != 2 {
		t.Errorf("count = %d, want 2", count)
	}
}

func TestDocument_RemoveComment(t *testing.T) {
	doc, _ := New(&Config{Author: "Bot"})
	defer doc.Close()
	r := doc.Body().AddParagraph().AddRun("text")
	c := doc.AddComment(r, r, "to remove")

	if err := doc.RemoveComment(c); err != nil {
		t.Fatalf("RemoveComment: %v", err)
	}

	count := 0
	for range doc.Comments() {
		count++
	}
	if count != 0 {
		t.Errorf("count = %d after remove", count)
	}
}

func TestDocument_AddComment_RoundTrip(t *testing.T) {
	doc, _ := New(&Config{Author: "Bot"})
	r := doc.Body().AddParagraph().AddRun("important")
	doc.AddComment(r, r, "review this")

	var buf bytes.Buffer
	doc.Write(&buf)

	doc2, err := OpenReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()), nil)
	if err != nil {
		t.Fatalf("OpenReader: %v", err)
	}
	defer doc2.Close()

	count := 0
	for c := range doc2.Comments() {
		if c.Text() == "review this" {
			count++
		}
	}
	if count != 1 {
		t.Errorf("comments after round-trip = %d", count)
	}
}

func TestDocument_AddComment_ID(t *testing.T) {
	doc, _ := New(&Config{Author: "Bot"})
	defer doc.Close()
	r1 := doc.Body().AddParagraph().AddRun("first")
	r2 := doc.Body().AddParagraph().AddRun("second")
	c1 := doc.AddComment(r1, r1, "comment 1")
	c2 := doc.AddComment(r2, r2, "comment 2")

	if c1.ID() == c2.ID() {
		t.Errorf("duplicate IDs: %d", c1.ID())
	}
}

func TestComment_Markdown(t *testing.T) {
	doc, _ := New(&Config{Author: "Bot"})
	defer doc.Close()
	r := doc.Body().AddParagraph().AddRun("text")
	c := doc.AddComment(r, r, "stub text")

	// Markdown is a stub; should return Text() for now.
	if c.Markdown() != c.Text() {
		t.Errorf("Markdown() = %q, want %q", c.Markdown(), c.Text())
	}
}

func TestComment_Paragraphs(t *testing.T) {
	doc, _ := New(&Config{Author: "Bot"})
	defer doc.Close()
	r := doc.Body().AddParagraph().AddRun("text")
	c := doc.AddComment(r, r, "para text")

	count := 0
	for p := range c.Paragraphs() {
		if p.Text() != "para text" {
			t.Errorf("paragraph text = %q", p.Text())
		}
		count++
	}
	if count != 1 {
		t.Errorf("paragraph count = %d, want 1", count)
	}
}

func TestDocument_CommentsForText(t *testing.T) {
	doc, _ := New(&Config{Author: "Bot"})
	defer doc.Close()
	r1 := doc.Body().AddParagraph().AddRun("find me here")
	r2 := doc.Body().AddParagraph().AddRun("not this")
	doc.AddComment(r1, r1, "found it")
	doc.AddComment(r2, r2, "other")

	count := 0
	for c := range doc.CommentsForText("find me") {
		if c.Text() != "found it" {
			t.Errorf("unexpected comment text = %q", c.Text())
		}
		count++
	}
	if count != 1 {
		t.Errorf("CommentsForText count = %d, want 1", count)
	}
}

// ---------------------------------------------------------------------------
// I9: WithMarkdown() on AddComment creates styled runs
// ---------------------------------------------------------------------------

func TestAddComment_WithMarkdown_BoldRun(t *testing.T) {
	doc, _ := New(&Config{Author: "Bot"})
	defer doc.Close()
	r := doc.Body().AddParagraph().AddRun("text")
	c := doc.AddComment(r, r, "This is **important** feedback", WithMarkdown())
	if c == nil {
		t.Fatal("nil comment")
	}

	// Plain text should contain the words (without formatting markers).
	text := c.Text()
	if !strings.Contains(text, "important") {
		t.Errorf("comment text = %q, want to contain 'important'", text)
	}

	// The comment body should have a bold run.
	var foundBold bool
	for p := range c.Paragraphs() {
		for run := range p.Runs() {
			if run.Text() == "important" && run.Bold() != nil && *run.Bold() {
				foundBold = true
			}
		}
	}
	if !foundBold {
		t.Error("expected bold run 'important' in comment paragraph")
	}
}

func TestAddComment_WithMarkdown_PlainText(t *testing.T) {
	doc, _ := New(&Config{Author: "Bot"})
	defer doc.Close()
	r := doc.Body().AddParagraph().AddRun("text")
	c := doc.AddComment(r, r, "just plain text", WithMarkdown())
	if c.Text() != "just plain text" {
		t.Errorf("comment text = %q, want 'just plain text'", c.Text())
	}
}

func TestAddComment_WithMarkdown_Markdown(t *testing.T) {
	doc, _ := New(&Config{Author: "Bot"})
	defer doc.Close()
	r := doc.Body().AddParagraph().AddRun("text")
	c := doc.AddComment(r, r, "review **this** section", WithMarkdown())

	// Comment.Markdown() should preserve bold formatting.
	md := c.Markdown()
	if !strings.Contains(md, "**this**") {
		t.Errorf("comment Markdown = %q, expected **this**", md)
	}
}

func TestAddComment_WithoutMarkdown_PlainRun(t *testing.T) {
	doc, _ := New(&Config{Author: "Bot"})
	defer doc.Close()
	r := doc.Body().AddParagraph().AddRun("text")
	c := doc.AddComment(r, r, "no **markdown** here")
	// Without WithMarkdown, the asterisks should be literal.
	if c.Text() != "no **markdown** here" {
		t.Errorf("comment text = %q, want literal asterisks", c.Text())
	}
}

func ExampleDocument_AddComment() {
	doc, _ := New(&Config{Author: "Reviewer"})
	defer doc.Close()
	r := doc.Body().AddParagraph().AddRun("TODO: fix this")
	doc.AddComment(r, r, "Please resolve before release")
	// doc.WriteTo(w)
}

func ExampleDocument_Comments() {
	doc, _ := New(&Config{Author: "Reviewer"})
	defer doc.Close()
	r := doc.Body().AddParagraph().AddRun("see note")
	doc.AddComment(r, r, "expand this section")

	for c := range doc.Comments() {
		_ = c.Author() // "Reviewer"
		_ = c.Text()   // "expand this section"
	}
}

// ---------------------------------------------------------------------------
// OOXML Compliance Tests
// ---------------------------------------------------------------------------

func TestAddComment_AnnotationRef(t *testing.T) {
	// Per ECMA-376 17.13.4.1, the first paragraph of a comment MUST contain
	// an annotationRef element.
	doc, _ := New(&Config{Author: "Bot"})
	defer doc.Close()
	r := doc.Body().AddParagraph().AddRun("text")
	c := doc.AddComment(r, r, "note")

	// Walk the comment's first paragraph runs looking for annotationRef.
	found := false
	for p := range c.Paragraphs() {
		for run := range p.Runs() {
			for _, rc := range run.el.Content {
				if rc.AnnotationRef != nil {
					found = true
				}
			}
		}
		break // only check first paragraph
	}
	if !found {
		t.Error("comment body missing annotationRef in first paragraph")
	}
}

func TestAddComment_RefRunHasCommentReferenceStyle(t *testing.T) {
	// The run containing commentReference in the document body should
	// have rStyle = "CommentReference".
	doc, _ := New(&Config{Author: "Bot"})
	defer doc.Close()
	p := doc.Body().AddParagraph()
	r := p.AddRun("text")
	doc.AddComment(r, r, "note")

	// Walk paragraph inline content looking for the commentReference run.
	doc.mu.RLock()
	defer doc.mu.RUnlock()
	for _, ic := range p.el.Content {
		if ic.Run != nil {
			for _, rc := range ic.Run.Content {
				if rc.CommentReference != nil {
					if ic.Run.RPr == nil || ic.Run.RPr.RunStyle == nil || *ic.Run.RPr.RunStyle != "CommentReference" {
						t.Error("commentReference run missing CommentReference rStyle")
					}
					return
				}
			}
		}
	}
	t.Error("commentReference run not found in paragraph")
}

func TestDocument_CommentsForText_CrossParagraph(t *testing.T) {
	doc, _ := New(&Config{Author: "Bot"})
	defer doc.Close()
	r1 := doc.Body().AddParagraph().AddRun("start here")
	r2 := doc.Body().AddParagraph().AddRun("end here")
	doc.AddComment(r1, r2, "spans two paragraphs")

	count := 0
	for c := range doc.CommentsForText("start here") {
		if c.Text() == "spans two paragraphs" {
			count++
		}
	}
	if count != 1 {
		t.Errorf("cross-paragraph CommentsForText count = %d, want 1", count)
	}
}

func TestAddComment_DefaultAuthor(t *testing.T) {
	// Per ECMA-376, w:author is required and should not be empty.
	doc, _ := New(nil) // No author configured
	defer doc.Close()
	r := doc.Body().AddParagraph().AddRun("text")
	c := doc.AddComment(r, r, "note")
	if c.Author() == "" {
		t.Error("comment author should not be empty when no author configured")
	}
	if c.Author() != "Unknown" {
		t.Errorf("default author = %q, want %q", c.Author(), "Unknown")
	}
}

func ExampleComment_Text() {
	doc, _ := New(&Config{Author: "Reviewer"})
	defer doc.Close()
	r := doc.Body().AddParagraph().AddRun("text")
	c := doc.AddComment(r, r, "needs work")
	_ = c.Text() // "needs work"
}

func ExampleComment_Paragraphs() {
	doc, _ := New(&Config{Author: "Reviewer"})
	defer doc.Close()
	r := doc.Body().AddParagraph().AddRun("text")
	c := doc.AddComment(r, r, "see above")
	for p := range c.Paragraphs() {
		_ = p.Text() // "see above"
	}
}

func ExampleDocument_RemoveComment() {
	doc, _ := New(&Config{Author: "Reviewer"})
	defer doc.Close()
	r := doc.Body().AddParagraph().AddRun("text")
	c := doc.AddComment(r, r, "to remove")
	_ = doc.RemoveComment(c)
}

func ExampleDocument_CommentsForText() {
	doc, _ := New(&Config{Author: "Reviewer"})
	defer doc.Close()
	r := doc.Body().AddParagraph().AddRun("find me here")
	doc.AddComment(r, r, "found")
	for c := range doc.CommentsForText("find me") {
		_ = c.Text() // "found"
	}
}
