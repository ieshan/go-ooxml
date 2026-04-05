package docx

import (
	"encoding/xml"
	"fmt"
	"iter"

	"github.com/ieshan/go-ooxml/docx/wml"
	"github.com/ieshan/go-ooxml/opc"
)

// HeaderFooterType identifies which variant of header or footer is referenced.
type HeaderFooterType int

const (
	// HdrFtrDefault is the default (odd-page) header or footer.
	HdrFtrDefault HeaderFooterType = iota
	// HdrFtrFirst is the first-page header or footer.
	HdrFtrFirst
	// HdrFtrEven is the even-page header or footer.
	HdrFtrEven
)

// HeaderFooter represents a header or footer part in the document.
// It has the same content model as the document body: paragraphs and tables.
type HeaderFooter struct {
	doc  *Document
	part *hdrFtrPart
}

// Type returns the header/footer variant (default, first, or even).
func (hf *HeaderFooter) Type() HeaderFooterType {
	return hf.part.typ
}

// Paragraphs returns an iterator over all paragraphs in the header/footer.
func (hf *HeaderFooter) Paragraphs() iter.Seq[*Paragraph] {
	return func(yield func(*Paragraph) bool) {
		hf.doc.mu.RLock()
		defer hf.doc.mu.RUnlock()
		if hf.part.body == nil {
			return
		}
		for _, c := range hf.part.body.Content {
			if c.Paragraph != nil {
				if !yield(&Paragraph{doc: hf.doc, el: c.Paragraph}) {
					return
				}
			}
		}
	}
}

// Tables returns an iterator over all tables in the header/footer.
func (hf *HeaderFooter) Tables() iter.Seq[*Table] {
	return func(yield func(*Table) bool) {
		hf.doc.mu.RLock()
		defer hf.doc.mu.RUnlock()
		if hf.part.body == nil {
			return
		}
		for _, c := range hf.part.body.Content {
			if c.Table != nil {
				if !yield(&Table{doc: hf.doc, el: c.Table}) {
					return
				}
			}
		}
	}
}

// AddParagraph appends a new empty paragraph to the header/footer and returns it.
func (hf *HeaderFooter) AddParagraph() *Paragraph {
	hf.doc.mu.Lock()
	defer hf.doc.mu.Unlock()
	if hf.part.body == nil {
		hf.part.body = &wml.CT_Body{XMLName: xml.Name{Space: wml.Ns, Local: "body"}}
	}
	p := &wml.CT_P{XMLName: xml.Name{Space: wml.Ns, Local: "p"}}
	hf.part.body.Content = append(hf.part.body.Content, wml.BlockLevelContent{Paragraph: p})
	return &Paragraph{doc: hf.doc, el: p}
}

// Text returns the plain text of the header/footer content.
func (hf *HeaderFooter) Text() string {
	hf.doc.mu.RLock()
	defer hf.doc.mu.RUnlock()
	if hf.part.body == nil {
		return ""
	}
	return bodyText(hf.part.body)
}

// Markdown returns the header/footer content formatted as Markdown.
func (hf *HeaderFooter) Markdown() string {
	hf.doc.mu.RLock()
	defer hf.doc.mu.RUnlock()
	if hf.part.body == nil {
		return ""
	}
	return bodyToMarkdown(hf.part.body, hf.doc.hyperlinkURLMap())
}

// Headers returns an iterator over all headers in this section.
func (s *Section) Headers() iter.Seq[*HeaderFooter] {
	return func(yield func(*HeaderFooter) bool) {
		s.doc.mu.RLock()
		defer s.doc.mu.RUnlock()
		for i := range s.doc.hdrFtrs {
			hf := &s.doc.hdrFtrs[i]
			if hf.isHdr {
				if !yield(&HeaderFooter{doc: s.doc, part: hf}) {
					return
				}
			}
		}
	}
}

// Footers returns an iterator over all footers in this section.
func (s *Section) Footers() iter.Seq[*HeaderFooter] {
	return func(yield func(*HeaderFooter) bool) {
		s.doc.mu.RLock()
		defer s.doc.mu.RUnlock()
		for i := range s.doc.hdrFtrs {
			hf := &s.doc.hdrFtrs[i]
			if !hf.isHdr {
				if !yield(&HeaderFooter{doc: s.doc, part: hf}) {
					return
				}
			}
		}
	}
}

// AddHeader creates a new header of the given type for this section and returns it.
func (s *Section) AddHeader(typ HeaderFooterType) *HeaderFooter {
	return s.addHdrFtr(typ, true)
}

// AddFooter creates a new footer of the given type for this section and returns it.
func (s *Section) AddFooter(typ HeaderFooterType) *HeaderFooter {
	return s.addHdrFtr(typ, false)
}

// addHdrFtr creates a new header or footer part, registers the relationship
// and content type, and adds a reference to the section properties.
func (s *Section) addHdrFtr(typ HeaderFooterType, isHdr bool) *HeaderFooter {
	s.doc.mu.Lock()
	defer s.doc.mu.Unlock()

	// Determine part path.
	num := s.doc.nextHdrFtrNum(isHdr)
	prefix := "footer"
	relType := opc.RelFooter
	ct := ctFooter
	if isHdr {
		prefix = "header"
		relType = opc.RelHeader
		ct = ctHeader
	}
	partPath := resolveRelPath(s.doc.docPath, fmt.Sprintf("%s%d.xml", prefix, num))

	// Create the body.
	body := &wml.CT_Body{XMLName: xml.Name{Space: wml.Ns, Local: "body"}}

	// Register the relationship on the document part.
	docPart, ok := s.doc.pkg.Parts[s.doc.docPath]
	if !ok {
		return nil
	}
	relID := opc.NextRelID(docPart.Rels)
	docPart.Rels = append(docPart.Rels, opc.Relationship{
		ID:     relID,
		Type:   relType,
		Target: fmt.Sprintf("%s%d.xml", prefix, num),
	})

	// Register content type.
	s.doc.pkg.ContentTypes.Overrides["/"+partPath] = ct

	// Create the OPC part with empty content (will be filled on save).
	s.doc.pkg.AddPart(partPath, ct, []byte{})

	// Add reference to section properties.
	typStr := hdrFtrTypeString(typ)
	local := "footerReference"
	if isHdr {
		local = "headerReference"
	}
	ref := wml.CT_HdrFtrRef{
		XMLName: xml.Name{Space: wml.Ns, Local: local},
		Type:    typStr,
		ID:      relID,
	}
	if isHdr {
		s.el.HeaderRefs = append(s.el.HeaderRefs, ref)
	} else {
		s.el.FooterRefs = append(s.el.FooterRefs, ref)
	}

	// Store in document.
	hfPart := hdrFtrPart{
		path:  partPath,
		relID: relID,
		typ:   typ,
		isHdr: isHdr,
		body:  body,
	}
	s.doc.hdrFtrs = append(s.doc.hdrFtrs, hfPart)

	return &HeaderFooter{doc: s.doc, part: &s.doc.hdrFtrs[len(s.doc.hdrFtrs)-1]}
}

// hdrFtrTypeString returns the OOXML type string for a HeaderFooterType.
func hdrFtrTypeString(typ HeaderFooterType) string {
	switch typ {
	case HdrFtrFirst:
		return "first"
	case HdrFtrEven:
		return "even"
	default:
		return "default"
	}
}
