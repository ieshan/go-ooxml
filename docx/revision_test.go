package docx

import (
	"bytes"
	"strings"
	"testing"
	"time"
)

func TestDocument_AddRevision(t *testing.T) {
	doc, _ := New(&Config{Author: "Editor"})
	defer doc.Close()
	p := doc.Body().AddParagraph()
	r := p.AddRun("draft")

	rev := doc.AddRevision(r, r, "final")
	if rev == nil {
		t.Fatal("nil revision")
	}
	if rev.Author() != "Editor" {
		t.Errorf("author = %q", rev.Author())
	}
	if rev.ProposedText() != "final" {
		t.Errorf("proposed = %q", rev.ProposedText())
	}
	if rev.OriginalText() != "" {
		t.Errorf("original text for insert should be empty, got %q", rev.OriginalText())
	}
	if rev.Type() != RevisionInsert {
		t.Errorf("type = %v, want RevisionInsert", rev.Type())
	}
}

func TestDocument_Revisions_Iterator(t *testing.T) {
	doc, _ := New(&Config{Author: "E"})
	defer doc.Close()
	r1 := doc.Body().AddParagraph().AddRun("old1")
	r2 := doc.Body().AddParagraph().AddRun("old2")
	doc.AddRevision(r1, r1, "new1")
	doc.AddRevision(r2, r2, "new2")

	count := 0
	for range doc.Revisions() {
		count++
	}
	// Each AddRevision creates both a del and an ins — count depends on implementation
	if count < 2 {
		t.Errorf("count = %d, want at least 2", count)
	}
}

func TestDocument_AcceptRevision(t *testing.T) {
	doc, _ := New(&Config{Author: "E"})
	defer doc.Close()
	r := doc.Body().AddParagraph().AddRun("old")
	rev := doc.AddRevision(r, r, "new")

	if err := doc.AcceptRevision(rev); err != nil {
		t.Fatalf("Accept: %v", err)
	}

	// After accepting, "new" should be normal text, "old" should be gone
	text := doc.Body().Text()
	if !strings.Contains(text, "new") {
		t.Errorf("text should contain 'new': %q", text)
	}
	if strings.Contains(text, "old") {
		t.Errorf("text should not contain 'old': %q", text)
	}
}

func TestDocument_RejectRevision(t *testing.T) {
	doc, _ := New(&Config{Author: "E"})
	defer doc.Close()
	r := doc.Body().AddParagraph().AddRun("old")
	rev := doc.AddRevision(r, r, "new")

	if err := doc.RejectRevision(rev); err != nil {
		t.Fatalf("Reject: %v", err)
	}

	// After rejecting, "old" should be normal text, "new" should be gone
	text := doc.Body().Text()
	if !strings.Contains(text, "old") {
		t.Errorf("text should contain 'old': %q", text)
	}
	if strings.Contains(text, "new") {
		t.Errorf("text should not contain 'new': %q", text)
	}
}

func TestDocument_AddRevision_WithOptions(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	r := doc.Body().AddParagraph().AddRun("text")
	rev := doc.AddRevision(r, r, "changed", WithRevisionAuthor("Jane"))
	if rev.Author() != "Jane" {
		t.Errorf("author = %q", rev.Author())
	}
}

func TestDocument_AddRevision_WithDate(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	r := doc.Body().AddParagraph().AddRun("text")
	dt := time.Date(2025, 3, 15, 10, 30, 0, 0, time.UTC)
	rev := doc.AddRevision(r, r, "changed", WithRevisionDate(dt))
	if !rev.Date().Equal(dt) {
		t.Errorf("date = %v, want %v", rev.Date(), dt)
	}
}

func TestDocument_RemoveRevision(t *testing.T) {
	doc, _ := New(&Config{Author: "E"})
	defer doc.Close()
	r := doc.Body().AddParagraph().AddRun("old")
	rev := doc.AddRevision(r, r, "new")

	if err := doc.RemoveRevision(rev); err != nil {
		t.Fatalf("Remove: %v", err)
	}

	// After removing, revision tracking is gone; count should be 0
	count := 0
	for range doc.Revisions() {
		count++
	}
	if count != 0 {
		t.Errorf("count after remove = %d, want 0", count)
	}
}

func TestDocument_OverrideRevision(t *testing.T) {
	doc, _ := New(&Config{Author: "E"})
	defer doc.Close()
	r := doc.Body().AddParagraph().AddRun("old")
	rev := doc.AddRevision(r, r, "new")

	rev2 := doc.OverrideRevision(rev, "newer", WithRevisionAuthor("F"))
	if rev2 == nil {
		t.Fatal("nil revision from override")
	}
	if rev2.ProposedText() != "newer" {
		t.Errorf("proposed = %q, want 'newer'", rev2.ProposedText())
	}
	if rev2.Author() != "F" {
		t.Errorf("author = %q, want 'F'", rev2.Author())
	}
}

func TestDocument_RevisionsForText(t *testing.T) {
	doc, _ := New(&Config{Author: "E"})
	defer doc.Close()
	r1 := doc.Body().AddParagraph().AddRun("old1")
	r2 := doc.Body().AddParagraph().AddRun("old2")
	doc.AddRevision(r1, r1, "new1")
	doc.AddRevision(r2, r2, "new2")

	count := 0
	for range doc.RevisionsForText("new1") {
		count++
	}
	if count != 1 {
		t.Errorf("count for 'new1' = %d, want 1", count)
	}
}

func TestDocument_Revision_ID(t *testing.T) {
	doc, _ := New(&Config{Author: "E"})
	defer doc.Close()
	r := doc.Body().AddParagraph().AddRun("old")
	rev := doc.AddRevision(r, r, "new")
	if rev.ID() < 1 {
		t.Errorf("ID = %d, want >= 1", rev.ID())
	}
}

func TestRevision_OriginalMarkdown_Stub(t *testing.T) {
	doc, _ := New(&Config{Author: "E"})
	defer doc.Close()
	r := doc.Body().AddParagraph().AddRun("old")
	rev := doc.AddRevision(r, r, "new")
	// Insert revision: OriginalMarkdown is empty (it's a deletion side), ProposedMarkdown has text.
	if rev.OriginalMarkdown() != "" {
		t.Errorf("OriginalMarkdown for insert revision should be empty, got %q", rev.OriginalMarkdown())
	}
	if rev.ProposedMarkdown() != "new" {
		t.Errorf("ProposedMarkdown = %q, want %q", rev.ProposedMarkdown(), "new")
	}
}

// ---------------------------------------------------------------------------
// I8: OriginalMarkdown / ProposedMarkdown use runToMarkdown
// ---------------------------------------------------------------------------

func TestRevision_ProposedMarkdown_BoldText(t *testing.T) {
	doc, _ := New(&Config{Author: "E"})
	defer doc.Close()
	r := doc.Body().AddParagraph().AddRun("old")
	rev := doc.AddRevision(r, r, "**bold**", WithMarkdownRevision())
	// ProposedMarkdown should reflect the bold formatting via runToMarkdown.
	md := rev.ProposedMarkdown()
	if !strings.Contains(md, "bold") {
		t.Errorf("ProposedMarkdown = %q, expected to contain bold text", md)
	}
}

func TestRevision_ProposedMarkdown_PlainText(t *testing.T) {
	doc, _ := New(&Config{Author: "E"})
	defer doc.Close()
	r := doc.Body().AddParagraph().AddRun("old")
	rev := doc.AddRevision(r, r, "plain new text")
	if rev.ProposedMarkdown() != "plain new text" {
		t.Errorf("ProposedMarkdown = %q, want %q", rev.ProposedMarkdown(), "plain new text")
	}
}

func TestRevision_OriginalMarkdown_DeleteRevision(t *testing.T) {
	doc, _ := New(&Config{Author: "E"})
	defer doc.Close()
	r := doc.Body().AddParagraph().AddRun("original text")
	doc.AddRevision(r, r, "replacement")

	// Find the delete revision (paired with the insert).
	var delRev *Revision
	for rv := range doc.Revisions() {
		if rv.Type() == RevisionDelete {
			delRev = rv
			break
		}
	}
	if delRev == nil {
		t.Fatal("delete revision not found")
	}
	// OriginalMarkdown should return the deleted text.
	orig := delRev.OriginalMarkdown()
	if !strings.Contains(orig, "original text") {
		t.Errorf("OriginalMarkdown = %q, want 'original text'", orig)
	}
	// ProposedMarkdown for a delete revision is empty.
	if delRev.ProposedMarkdown() != "" {
		t.Errorf("ProposedMarkdown for delete should be empty, got %q", delRev.ProposedMarkdown())
	}
}

// ---------------------------------------------------------------------------
// I9: WithMarkdownRevision creates styled runs in the ins element
// ---------------------------------------------------------------------------

func TestAddRevision_WithMarkdownRevision_BoldRun(t *testing.T) {
	doc, _ := New(&Config{Author: "E"})
	defer doc.Close()
	r := doc.Body().AddParagraph().AddRun("old")
	rev := doc.AddRevision(r, r, "**bold** text", WithMarkdownRevision())
	if rev == nil {
		t.Fatal("nil revision")
	}

	// The proposed text (via Text()) should strip formatting and contain the words.
	proposed := rev.ProposedText()
	if !strings.Contains(proposed, "bold") || !strings.Contains(proposed, "text") {
		t.Errorf("ProposedText = %q, want to contain 'bold' and 'text'", proposed)
	}

	// ProposedMarkdown should contain bold formatting markers.
	md := rev.ProposedMarkdown()
	if !strings.Contains(md, "**bold**") {
		t.Errorf("ProposedMarkdown = %q, expected **bold**", md)
	}
}

func TestAddRevision_WithMarkdownRevision_PlainFallback(t *testing.T) {
	doc, _ := New(&Config{Author: "E"})
	defer doc.Close()
	r := doc.Body().AddParagraph().AddRun("old")
	rev := doc.AddRevision(r, r, "plain text", WithMarkdownRevision())
	if rev.ProposedText() != "plain text" {
		t.Errorf("ProposedText = %q, want 'plain text'", rev.ProposedText())
	}
}

func ExampleDocument_AddRevision() {
	doc, _ := New(&Config{Author: "Editor"})
	defer doc.Close()
	r := doc.Body().AddParagraph().AddRun("draft")
	doc.AddRevision(r, r, "final")
	// doc.WriteTo(w)
}

func ExampleDocument_Revisions() {
	doc, _ := New(&Config{Author: "Editor"})
	defer doc.Close()
	r := doc.Body().AddParagraph().AddRun("old text")
	doc.AddRevision(r, r, "new text")

	for rev := range doc.Revisions() {
		_ = rev.Author()       // "Editor"
		_ = rev.ProposedText() // "new text" (for insert revisions)
	}
}

func ExampleDocument_AcceptRevision() {
	doc, _ := New(&Config{Author: "Editor"})
	defer doc.Close()
	r := doc.Body().AddParagraph().AddRun("draft")
	rev := doc.AddRevision(r, r, "final")
	if err := doc.AcceptRevision(rev); err != nil {
		// handle error
		return
	}
	// The paragraph now contains "final" as permanent text.
}

// ---------------------------------------------------------------------------
// OOXML Compliance Tests
// ---------------------------------------------------------------------------

func TestRevision_UniqueIDs(t *testing.T) {
	// Per ECMA-376 17.13.1, each revision mark MUST have a unique w:id.
	doc, _ := New(&Config{Author: "E"})
	defer doc.Close()
	r := doc.Body().AddParagraph().AddRun("old")
	doc.AddRevision(r, r, "new")

	ids := make(map[int]bool)
	for rev := range doc.Revisions() {
		if ids[rev.ID()] {
			t.Errorf("duplicate revision ID: %d", rev.ID())
		}
		ids[rev.ID()] = true
	}
	if len(ids) < 2 {
		t.Errorf("expected at least 2 unique IDs, got %d", len(ids))
	}
}

func TestRevision_DelText_UsesDelTextElement(t *testing.T) {
	// Verify that deleted runs use delText (not text) per OOXML spec.
	doc, _ := New(&Config{Author: "E"})
	defer doc.Close()
	r := doc.Body().AddParagraph().AddRun("original")
	doc.AddRevision(r, r, "replacement")

	for rev := range doc.Revisions() {
		if rev.Type() == RevisionDelete {
			// The deleted text should be accessible via OriginalText.
			if rev.OriginalText() != "original" {
				t.Errorf("OriginalText = %q, want %q", rev.OriginalText(), "original")
			}
			return
		}
	}
	t.Error("delete revision not found")
}

func TestRevision_DefaultAuthor(t *testing.T) {
	// Per ECMA-376, w:author is required and should not be empty.
	doc, _ := New(nil) // No author configured
	defer doc.Close()
	r := doc.Body().AddParagraph().AddRun("text")
	rev := doc.AddRevision(r, r, "changed")
	if rev.Author() == "" {
		t.Error("revision author should not be empty when no author configured")
	}
	if rev.Author() != "Unknown" {
		t.Errorf("default author = %q, want %q", rev.Author(), "Unknown")
	}
}

func TestRevision_RoundTrip(t *testing.T) {
	doc, _ := New(&Config{Author: "Editor"})
	r := doc.Body().AddParagraph().AddRun("old text")
	doc.AddRevision(r, r, "new text")

	var buf bytes.Buffer
	doc.Write(&buf)

	doc2, err := OpenReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()), nil)
	if err != nil {
		t.Fatalf("OpenReader: %v", err)
	}
	defer doc2.Close()

	insCount := 0
	delCount := 0
	for rev := range doc2.Revisions() {
		switch rev.Type() {
		case RevisionInsert:
			insCount++
			if rev.ProposedText() != "new text" {
				t.Errorf("proposed = %q", rev.ProposedText())
			}
		case RevisionDelete:
			delCount++
			if rev.OriginalText() != "old text" {
				t.Errorf("original = %q", rev.OriginalText())
			}
		}
	}
	if insCount != 1 {
		t.Errorf("insert count = %d after round-trip", insCount)
	}
	if delCount != 1 {
		t.Errorf("delete count = %d after round-trip", delCount)
	}
}

func ExampleDocument_RejectRevision() {
	doc, _ := New(&Config{Author: "Editor"})
	defer doc.Close()
	r := doc.Body().AddParagraph().AddRun("draft")
	rev := doc.AddRevision(r, r, "final")
	_ = doc.RejectRevision(rev)
	// The paragraph now contains "draft" as permanent text.
}

func ExampleDocument_OverrideRevision() {
	doc, _ := New(&Config{Author: "Editor"})
	defer doc.Close()
	r := doc.Body().AddParagraph().AddRun("draft")
	rev := doc.AddRevision(r, r, "version 1")
	doc.OverrideRevision(rev, "version 2")
}

func ExampleDocument_RevisionsForText() {
	doc, _ := New(&Config{Author: "Editor"})
	defer doc.Close()
	r := doc.Body().AddParagraph().AddRun("old text")
	doc.AddRevision(r, r, "new text")
	for rev := range doc.RevisionsForText("new") {
		_ = rev.ProposedText() // "new text"
	}
}

func ExampleRevision_OriginalText() {
	doc, _ := New(&Config{Author: "Editor"})
	defer doc.Close()
	r := doc.Body().AddParagraph().AddRun("before")
	doc.AddRevision(r, r, "after")
	for rev := range doc.Revisions() {
		if rev.Type() == RevisionDelete {
			_ = rev.OriginalText() // "before"
		}
	}
}

func ExampleRevision_ProposedText() {
	doc, _ := New(&Config{Author: "Editor"})
	defer doc.Close()
	r := doc.Body().AddParagraph().AddRun("before")
	doc.AddRevision(r, r, "after")
	for rev := range doc.Revisions() {
		if rev.Type() == RevisionInsert {
			_ = rev.ProposedText() // "after"
		}
	}
}
