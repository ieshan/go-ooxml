package docx

import (
	"testing"
)

func TestHeaderFooterType_Constants(t *testing.T) {
	// Verify the constants have distinct values.
	types := []HeaderFooterType{HdrFtrDefault, HdrFtrFirst, HdrFtrEven}
	seen := make(map[HeaderFooterType]bool)
	for _, typ := range types {
		if seen[typ] {
			t.Errorf("duplicate HeaderFooterType value %d", typ)
		}
		seen[typ] = true
	}
}

func TestHeaderFooter_TextMarkdown_Stubs(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()

	hf := &HeaderFooter{doc: doc, typ: HdrFtrDefault}
	if got := hf.Text(); got != "" {
		t.Errorf("Text() = %q, want empty string", got)
	}
	if got := hf.Markdown(); got != "" {
		t.Errorf("Markdown() = %q, want empty string", got)
	}
}
