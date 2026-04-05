package xmlutil

import (
	"bytes"
	"encoding/xml"
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

func (r *RawXML) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	// Capture inner XML
	var inner struct {
		Content []byte `xml:",innerxml"`
	}
	if err := d.DecodeElement(&inner, &start); err != nil {
		return err
	}
	// Rebuild full element
	var buf bytes.Buffer
	enc := xml.NewEncoder(&buf)
	if err := enc.EncodeToken(start); err != nil {
		return err
	}
	if err := enc.Flush(); err != nil {
		return err
	}
	buf.Write(inner.Content)
	if err := enc.EncodeToken(start.End()); err != nil {
		return err
	}
	if err := enc.Flush(); err != nil {
		return err
	}
	r.Data = buf.Bytes()
	return nil
}

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
