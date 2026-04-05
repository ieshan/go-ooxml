package docx

import (
	"testing"

	"github.com/ieshan/go-ooxml/common"
)

func TestRun_Bold(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	r := doc.Body().AddParagraph().AddRun("bold test")
	if r.Bold() != nil {
		t.Error("default should be nil (inherit)")
	}
	tr := true
	r.SetBold(&tr)
	if r.Bold() == nil || *r.Bold() != true {
		t.Error("not bold")
	}
	r.SetBold(nil)
	if r.Bold() != nil {
		t.Error("should be nil after reset")
	}
}

func TestRun_Italic(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	r := doc.Body().AddParagraph().AddRun("test")
	tr := true
	r.SetItalic(&tr)
	if r.Italic() == nil || *r.Italic() != true {
		t.Error("not italic")
	}
}

func TestRun_SetText(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	r := doc.Body().AddParagraph().AddRun("old")
	r.SetText("new")
	if r.Text() != "new" {
		t.Errorf("text = %q", r.Text())
	}
}

func TestRun_FontName(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	r := doc.Body().AddParagraph().AddRun("test")
	if r.FontName() != nil {
		t.Error("default should be nil")
	}
	r.SetFontName("Arial")
	if r.FontName() == nil || *r.FontName() != "Arial" {
		t.Error("font not set")
	}
}

func TestRun_FontSize(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	r := doc.Body().AddParagraph().AddRun("test")
	if r.FontSize() != nil {
		t.Error("default should be nil")
	}
	r.SetFontSize(12.0)
	if r.FontSize() == nil || *r.FontSize() != 12.0 {
		t.Error("font size not set")
	}
}

func TestRun_FontColor(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	r := doc.Body().AddParagraph().AddRun("test")
	if r.FontColor() != nil {
		t.Error("default should be nil")
	}
	r.SetFontColor(common.RGB(255, 0, 0))
	if r.FontColor() == nil || r.FontColor().R != 255 {
		t.Error("font color not set")
	}
}

func TestRun_Underline(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	r := doc.Body().AddParagraph().AddRun("test")
	if r.Underline() != nil {
		t.Error("default should be nil")
	}
	tr := true
	r.SetUnderline(&tr)
	if r.Underline() == nil || *r.Underline() != true {
		t.Error("not underlined")
	}
	fa := false
	r.SetUnderline(&fa)
	if r.Underline() == nil || *r.Underline() != false {
		t.Error("should be explicitly false")
	}
	r.SetUnderline(nil)
	if r.Underline() != nil {
		t.Error("should be nil after clear")
	}
}

func TestRun_Strikethrough(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	r := doc.Body().AddParagraph().AddRun("test")
	if r.Strikethrough() != nil {
		t.Error("default should be nil")
	}
	tr := true
	r.SetStrikethrough(&tr)
	if r.Strikethrough() == nil || *r.Strikethrough() != true {
		t.Error("not strikethrough")
	}
}

func TestRun_AddBreaks(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	r := doc.Body().AddParagraph().AddRun("before")
	r.AddTab()
	r.AddLineBreak()
	text := r.Text()
	if text != "before\t\n" {
		t.Errorf("text = %q", text)
	}
}

func TestRun_AddPageBreak(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	r := doc.Body().AddParagraph().AddRun("")
	r.AddPageBreak()
	// Page break renders as \n in text
	text := r.Text()
	if text != "\n" {
		t.Errorf("text = %q, want newline for page break", text)
	}
}

func ExampleRun_SetBold() {
	doc, _ := New(nil)
	defer doc.Close()
	r := doc.Body().AddParagraph().AddRun("Important")
	tr := true
	r.SetBold(&tr)
}

func ExampleRun_Markdown() {
	doc, _ := New(nil)
	defer doc.Close()
	r := doc.Body().AddParagraph().AddRun("bold text")
	b := true
	r.SetBold(&b)
	_ = r.Markdown() // "**bold text**"
}

func ExampleRun_SetFontName() {
	doc, _ := New(nil)
	defer doc.Close()
	r := doc.Body().AddParagraph().AddRun("code")
	r.SetFontName("Courier New")
}

func ExampleRun_SetFontSize() {
	doc, _ := New(nil)
	defer doc.Close()
	r := doc.Body().AddParagraph().AddRun("large")
	r.SetFontSize(24.0)
}
