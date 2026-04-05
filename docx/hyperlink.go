package docx

import (
	"encoding/xml"
	"iter"
	"strings"

	"github.com/ieshan/go-ooxml/docx/wml"
	"github.com/ieshan/go-ooxml/opc"
)

// Hyperlink wraps a wml.CT_Hyperlink element and its containing document.
type Hyperlink struct {
	doc *Document
	el  *wml.CT_Hyperlink
}

// URL resolves the hyperlink's target URL by looking up its relationship ID
// in the document part's relationship collection.
func (h *Hyperlink) URL() string {
	h.doc.mu.RLock()
	defer h.doc.mu.RUnlock()

	if h.el.ID == "" {
		return h.el.Anchor
	}

	docPart, ok := h.doc.pkg.Parts[h.doc.docPath]
	if !ok {
		return ""
	}

	for _, rel := range docPart.Rels {
		if rel.ID == h.el.ID && rel.Type == opc.RelHyperlink {
			return rel.Target
		}
	}
	return ""
}

// Text concatenates the text of all runs inside the hyperlink element.
func (h *Hyperlink) Text() string {
	h.doc.mu.RLock()
	defer h.doc.mu.RUnlock()

	var sb strings.Builder
	for _, r := range h.el.Runs {
		sb.WriteString(r.Text())
	}
	return sb.String()
}

// Runs returns an iterator over all runs inside the hyperlink.
func (h *Hyperlink) Runs() iter.Seq[*Run] {
	return func(yield func(*Run) bool) {
		h.doc.mu.RLock()
		defer h.doc.mu.RUnlock()

		for _, r := range h.el.Runs {
			if !yield(&Run{doc: h.doc, el: r}) {
				return
			}
		}
	}
}

// Hyperlinks returns an iterator over all hyperlinks in the paragraph.
func (p *Paragraph) Hyperlinks() iter.Seq[*Hyperlink] {
	return func(yield func(*Hyperlink) bool) {
		p.doc.mu.RLock()
		defer p.doc.mu.RUnlock()

		for _, c := range p.el.Content {
			if c.Hyperlink != nil {
				if !yield(&Hyperlink{doc: p.doc, el: c.Hyperlink}) {
					return
				}
			}
		}
	}
}

// AddHyperlink creates a new hyperlink with the given URL and display text,
// adds it to the paragraph, and returns the Hyperlink wrapper.
//
// The URL is registered as an external relationship on the document part.
func (p *Paragraph) AddHyperlink(url, text string) *Hyperlink {
	p.doc.mu.Lock()
	defer p.doc.mu.Unlock()

	// Register the relationship on the document part.
	relID := p.doc.addHyperlinkRelationship(url)

	// Build the CT_Hyperlink with a single run containing the text.
	r := &wml.CT_R{XMLName: xml.Name{Space: wml.Ns, Local: "r"}}
	r.AddText(text)

	el := &wml.CT_Hyperlink{
		XMLName: xml.Name{Space: wml.Ns, Local: "hyperlink"},
		ID:      relID,
		Runs:    []*wml.CT_R{r},
	}

	p.el.Content = append(p.el.Content, wml.InlineContent{Hyperlink: el})
	return &Hyperlink{doc: p.doc, el: el}
}

// hyperlinkURLMap builds a map from relationship ID → target URL for all
// hyperlink relationships on the document part.
// Caller must hold at least a read lock before calling this.
func (d *Document) hyperlinkURLMap() map[string]string {
	docPart, ok := d.pkg.Parts[d.docPath]
	if !ok {
		return nil
	}
	m := make(map[string]string, len(docPart.Rels))
	for _, rel := range docPart.Rels {
		if rel.Type == opc.RelHyperlink {
			m[rel.ID] = rel.Target
		}
	}
	return m
}

// addHyperlinkRelationship adds an external hyperlink relationship to the
// document part and returns the generated relationship ID.
// Caller must hold the write lock.
func (d *Document) addHyperlinkRelationship(url string) string {
	docPart, ok := d.pkg.Parts[d.docPath]
	if !ok {
		return ""
	}

	relID := opc.NextRelID(docPart.Rels)
	docPart.Rels = append(docPart.Rels, opc.Relationship{
		ID:       relID,
		Type:     opc.RelHyperlink,
		Target:   url,
		External: true,
	})
	return relID
}
