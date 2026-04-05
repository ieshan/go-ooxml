package docx

import "testing"

func TestParagraph_AddRun(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	p := doc.Body().AddParagraph()
	r := p.AddRun("hello")
	if r == nil {
		t.Fatal("nil run")
	}
	if r.Text() != "hello" {
		t.Errorf("text = %q", r.Text())
	}
}

func TestParagraph_Runs(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	p := doc.Body().AddParagraph()
	p.AddRun("one")
	p.AddRun("two")
	count := 0
	for range p.Runs() {
		count++
	}
	if count != 2 {
		t.Errorf("count = %d, want 2", count)
	}
}

func TestParagraph_Text(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	p := doc.Body().AddParagraph()
	p.AddRun("Hello ")
	p.AddRun("World")
	if p.Text() != "Hello World" {
		t.Errorf("text = %q", p.Text())
	}
}

func TestParagraph_Style(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	p := doc.Body().AddParagraph()
	if p.Style() != "" {
		t.Error("default style should be empty")
	}
	p.SetStyle("Heading1")
	if p.Style() != "Heading1" {
		t.Errorf("style = %q", p.Style())
	}
}

func TestParagraph_Markdown_Empty(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	p := doc.Body().AddParagraph()
	if p.Markdown() != "" {
		t.Error("empty paragraph Markdown() should return empty string")
	}
}

func TestParagraph_Format(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	p := doc.Body().AddParagraph()
	if p.Format() == nil {
		t.Error("Format() should return non-nil")
	}
}

func TestParagraphFormat_Alignment(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	p := doc.Body().AddParagraph()
	f := p.Format()
	if f == nil {
		t.Fatal("nil format")
	}
	f.SetAlignment("center")
	if f.Alignment() != "center" {
		t.Errorf("got %q", f.Alignment())
	}
}

func ExampleParagraph_SetStyle() {
	doc, _ := New(nil)
	defer doc.Close()
	p := doc.Body().AddParagraph()
	p.SetStyle("Heading1")
	p.AddRun("Introduction")
	// doc.WriteTo(w)
}

func ExampleParagraph_AddRun() {
	doc, _ := New(nil)
	defer doc.Close()
	p := doc.Body().AddParagraph()
	r := p.AddRun("Hello ")
	tr := true
	r.SetBold(&tr)
	p.AddRun("World")
}

func ExampleParagraph_Markdown() {
	doc, _ := New(nil)
	defer doc.Close()
	p := doc.Body().AddParagraph()
	p.SetStyle("Heading1")
	p.AddRun("Title")
	_ = p.Markdown() // "# Title"
}

func ExampleParagraph_Format() {
	doc, _ := New(nil)
	defer doc.Close()
	p := doc.Body().AddParagraph()
	p.AddRun("centered text")
	p.Format().SetAlignment("center")
}
