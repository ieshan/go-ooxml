package docx

import (
	"encoding/xml"
	"iter"

	"github.com/ieshan/go-ooxml/docx/wml"
)

// Paragraph wraps a wml.CT_P element.
type Paragraph struct {
	doc *Document
	el  *wml.CT_P
}

// Runs returns an iterator over all runs in the paragraph.
func (p *Paragraph) Runs() iter.Seq[*Run] {
	return func(yield func(*Run) bool) {
		p.doc.mu.RLock()
		defer p.doc.mu.RUnlock()
		for _, c := range p.el.Content {
			if c.Run != nil {
				if !yield(&Run{doc: p.doc, el: c.Run}) {
					return
				}
			}
		}
	}
}

// Text concatenates the text of all runs in the paragraph, including text
// inside tracked changes (insertions, deletions) and hyperlinks.
func (p *Paragraph) Text() string {
	p.doc.mu.RLock()
	defer p.doc.mu.RUnlock()
	return p.el.Text()
}

// Markdown returns the paragraph formatted as Markdown, including heading
// prefixes, list markers, blockquote prefixes, and run-level formatting.
// Hyperlinks are rendered as [text](url) using the document's relationships.
func (p *Paragraph) Markdown() string {
	p.doc.mu.RLock()
	defer p.doc.mu.RUnlock()
	return paragraphToMarkdown(p.el, p.doc.hyperlinkURLMap())
}

// AddRun adds a new run with the given text to the paragraph and returns it.
func (p *Paragraph) AddRun(text string) *Run {
	p.doc.mu.Lock()
	defer p.doc.mu.Unlock()

	r := &wml.CT_R{XMLName: xml.Name{Space: wml.Ns, Local: "r"}}
	r.AddText(text)
	p.el.Content = append(p.el.Content, wml.InlineContent{Run: r})
	return &Run{doc: p.doc, el: r}
}

// Style returns the paragraph style name, or empty string if none is set.
func (p *Paragraph) Style() string {
	p.doc.mu.RLock()
	defer p.doc.mu.RUnlock()

	if p.el.PPr != nil && p.el.PPr.Style != nil {
		return *p.el.PPr.Style
	}
	return ""
}

// SetStyle sets the paragraph style name.
func (p *Paragraph) SetStyle(name string) {
	p.doc.mu.Lock()
	defer p.doc.mu.Unlock()

	if p.el.PPr == nil {
		p.el.PPr = &wml.CT_PPr{
			XMLName: xml.Name{Space: wml.Ns, Local: "pPr"},
		}
	}
	p.el.PPr.Style = &name
}

// Format returns the paragraph formatting properties, creating the pPr element
// if it does not already exist.
func (p *Paragraph) Format() *ParagraphFormat {
	p.doc.mu.Lock()
	defer p.doc.mu.Unlock()
	if p.el.PPr == nil {
		p.el.PPr = &wml.CT_PPr{XMLName: xml.Name{Space: wml.Ns, Local: "pPr"}}
	}
	return &ParagraphFormat{doc: p.doc, el: p.el.PPr}
}

// ParagraphFormat wraps paragraph formatting.
type ParagraphFormat struct {
	doc *Document
	el  *wml.CT_PPr
}

// Alignment returns the paragraph alignment (e.g. "center", "left", "right",
// "both"), or empty string if not set.
func (pf *ParagraphFormat) Alignment() string {
	pf.doc.mu.RLock()
	defer pf.doc.mu.RUnlock()
	if pf.el.Alignment == nil {
		return ""
	}
	return *pf.el.Alignment
}

// SetAlignment sets the paragraph alignment. Pass an empty string to clear.
func (pf *ParagraphFormat) SetAlignment(a string) {
	pf.doc.mu.Lock()
	defer pf.doc.mu.Unlock()
	if a == "" {
		pf.el.Alignment = nil
		return
	}
	pf.el.Alignment = &a
}
