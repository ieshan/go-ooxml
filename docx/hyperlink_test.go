package docx

import (
	"bytes"
	"testing"
)

func TestParagraph_Hyperlinks_Empty(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	p := doc.Body().AddParagraph()
	p.AddRun("no links")
	count := 0
	for range p.Hyperlinks() {
		count++
	}
	if count != 0 {
		t.Error("should be empty")
	}
}

func TestParagraph_AddHyperlink(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	p := doc.Body().AddParagraph()
	h := p.AddHyperlink("https://example.com", "Click here")
	if h == nil {
		t.Fatal("expected non-nil Hyperlink")
	}
	if h.URL() != "https://example.com" {
		t.Errorf("URL = %q, want %q", h.URL(), "https://example.com")
	}
	if h.Text() != "Click here" {
		t.Errorf("Text = %q, want %q", h.Text(), "Click here")
	}
}

func TestParagraph_Hyperlinks_Iter(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	p := doc.Body().AddParagraph()
	p.AddHyperlink("https://a.com", "A")
	p.AddHyperlink("https://b.com", "B")

	count := 0
	for range p.Hyperlinks() {
		count++
	}
	if count != 2 {
		t.Errorf("count = %d, want 2", count)
	}
}

func TestHyperlink_Runs(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	p := doc.Body().AddParagraph()
	h := p.AddHyperlink("https://example.com", "Hello")

	count := 0
	for r := range h.Runs() {
		count++
		if r.Text() != "Hello" {
			t.Errorf("run text = %q, want %q", r.Text(), "Hello")
		}
	}
	if count != 1 {
		t.Errorf("run count = %d, want 1", count)
	}
}

func ExampleParagraph_AddHyperlink() {
	doc, _ := New(nil)
	defer doc.Close()
	p := doc.Body().AddParagraph()
	p.AddRun("Visit ")
	p.AddHyperlink("https://example.com", "our website")
	p.AddRun(" for more info.")
	// doc.WriteTo(w)
}

func TestHyperlink_RoundTrip(t *testing.T) {
	doc, _ := New(nil)
	p := doc.Body().AddParagraph()
	p.AddHyperlink("https://roundtrip.com", "RoundTrip")

	var buf bytes.Buffer
	if err := doc.Write(&buf); err != nil {
		t.Fatalf("WriteTo: %v", err)
	}

	doc2, err := OpenReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()), nil)
	if err != nil {
		t.Fatalf("OpenReader: %v", err)
	}
	defer doc2.Close()

	var found bool
	for p2 := range doc2.Body().Paragraphs() {
		for h := range p2.Hyperlinks() {
			if h.URL() == "https://roundtrip.com" && h.Text() == "RoundTrip" {
				found = true
			}
		}
	}
	if !found {
		t.Error("hyperlink not found after round-trip")
	}
}
