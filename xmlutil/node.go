package xmlutil

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
)

// RawXML holds a raw XML fragment that is preserved verbatim during
// marshal/unmarshal. Use it as a catch-all field in structs to capture
// unknown child elements without losing them during round-trip:
//
//	type Body struct {
//	    Paragraphs []*Paragraph `xml:"p"`
//	    Extra      []RawXML     `xml:",any"` // Unknown elements preserved
//	}
//
// RawXML implements [xml.Marshaler] and [xml.Unmarshaler].
type RawXML struct {
	Data []byte
}

// UnmarshalXML captures an XML element and its content as raw bytes.
// It manually reconstructs the element tag to preserve namespace prefixes,
// which Go's [xml.Encoder] does not do correctly (it generates broken
// declarations like xmlns="w" instead of xmlns:w="http://...").
func (r *RawXML) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	// Capture inner XML (raw bytes, preserves original prefixes).
	var inner struct {
		Content []byte `xml:",innerxml"`
	}
	if err := d.DecodeElement(&inner, &start); err != nil {
		return err
	}

	// Manually reconstruct the element from start + inner + end.
	// We avoid Go's xml.Encoder because it mangles namespace prefixes —
	// it doesn't preserve the original prefix (e.g., "w:") and generates
	// broken declarations like xmlns="w" instead of xmlns:w="http://...".
	var buf bytes.Buffer

	// Build a URI → prefix map from xmlns attributes on this element.
	nsMap := make(map[string]string) // URI → prefix
	for _, a := range start.Attr {
		if a.Name.Space == "xmlns" {
			nsMap[a.Value] = a.Name.Local // xmlns:w="http://..." → "http://..." → "w"
		}
	}

	// Determine the prefix for the element's own namespace.
	// Track a counter for generating unique fallback prefixes (ns0, ns1, ...).
	nsCounter := 0
	elemPrefix := nsMap[start.Name.Space]
	if elemPrefix == "" && start.Name.Space != "" {
		// The element inherits its namespace from a parent context.
		// We need to emit an xmlns:prefix declaration. Infer the prefix
		// from well-known OOXML conventions or generate a unique one.
		elemPrefix = inferPrefixFromContent(inner.Content, start.Name.Space)
		if elemPrefix == "" {
			elemPrefix = fmt.Sprintf("ns%d", nsCounter)
			nsCounter++
		}
		nsMap[start.Name.Space] = elemPrefix
	}

	// Write opening tag.
	buf.WriteByte('<')
	if elemPrefix != "" {
		buf.WriteString(elemPrefix)
		buf.WriteByte(':')
	}
	buf.WriteString(start.Name.Local)

	// Write xmlns declarations for any namespaces we resolved.
	// These are needed so inner content's prefix references are valid.
	for uri, prefix := range nsMap {
		buf.WriteString(" xmlns:")
		buf.WriteString(prefix)
		buf.WriteString(`="`)
		xml.EscapeText(&buf, []byte(uri))
		buf.WriteByte('"')
	}

	// Write non-namespace attributes, restoring prefixed names.
	for _, a := range start.Attr {
		if a.Name.Space == "xmlns" || (a.Name.Space == "" && a.Name.Local == "xmlns") {
			continue // Already handled above.
		}
		buf.WriteByte(' ')
		if a.Name.Space != "" {
			p := nsMap[a.Name.Space]
			if p == "" {
				// Unknown namespace on attribute — generate a unique prefix.
				p = inferPrefixFromContent(inner.Content, a.Name.Space)
				if p == "" {
					p = fmt.Sprintf("ns%d", nsCounter)
					nsCounter++
				}
				nsMap[a.Name.Space] = p
			}
			buf.WriteString(p)
			buf.WriteByte(':')
		}
		buf.WriteString(a.Name.Local)
		buf.WriteString(`="`)
		xml.EscapeText(&buf, []byte(a.Value))
		buf.WriteByte('"')
	}
	buf.WriteByte('>')

	// Write inner content (already raw XML with correct prefixes).
	buf.Write(inner.Content)

	// Write closing tag.
	buf.WriteString("</")
	if elemPrefix != "" {
		buf.WriteString(elemPrefix)
		buf.WriteByte(':')
	}
	buf.WriteString(start.Name.Local)
	buf.WriteByte('>')

	r.Data = buf.Bytes()
	return nil
}

// inferPrefixFromContent scans raw XML bytes for a namespace prefix that maps
// to the given URI. It looks for patterns like "prefix:" at the start of element
// names. This is a heuristic for reconstructing prefix information lost by
// Go's xml.Decoder.
func inferPrefixFromContent(content []byte, uri string) string {
	// Common OOXML prefix conventions.
	knownPrefixes := map[string]string{
		"http://schemas.openxmlformats.org/wordprocessingml/2006/main":           "w",
		"http://schemas.openxmlformats.org/officeDocument/2006/relationships":    "r",
		"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing": "wp",
		"http://schemas.openxmlformats.org/drawingml/2006/main":                  "a",
		"http://schemas.openxmlformats.org/markup-compatibility/2006":            "mc",
		"http://schemas.microsoft.com/office/word/2010/wordml":                   "w14",
	}
	if p, ok := knownPrefixes[uri]; ok {
		return p
	}
	return ""
}

// MarshalXML writes the stored raw XML bytes into the encoder by decoding
// the stored data token-by-token and re-encoding each token. This preserves
// the original XML structure while allowing the parent encoder to manage
// the overall output format.
func (r *RawXML) MarshalXML(e *xml.Encoder, _ xml.StartElement) error {
	if len(r.Data) == 0 {
		return nil
	}
	d := xml.NewDecoder(bytes.NewReader(r.Data))
	depth := 0
	for {
		tok, err := d.Token()
		if err != nil {
			if err == io.EOF && depth == 0 {
				return nil
			}
			return err
		}
		if err := e.EncodeToken(xml.CopyToken(tok)); err != nil {
			return err
		}
		switch tok.(type) {
		case xml.StartElement:
			depth++
		case xml.EndElement:
			depth--
			if depth == 0 {
				return nil
			}
		}
	}
}
