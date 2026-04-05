package wml

import (
	"encoding/xml"

	"github.com/ieshan/go-ooxml/xmlutil"
)

// CT_Document represents <w:document> — the root element of document.xml
type CT_Document struct {
	XMLName xml.Name
	Body    *CT_Body
}

// CT_Body represents <w:body>
type CT_Body struct {
	XMLName xml.Name
	Content []BlockLevelContent
	SectPr  *xmlutil.RawXML // Section properties (preserved as raw for now)
}

// BlockLevelContent holds one block-level child. Exactly one field is non-nil.
type BlockLevelContent struct {
	Paragraph *CT_P
	Table     *CT_Tbl
	Raw       *xmlutil.RawXML
}

// Paragraphs returns all paragraphs in the body.
func (b *CT_Body) Paragraphs() []*CT_P {
	var result []*CT_P
	for _, c := range b.Content {
		if c.Paragraph != nil {
			result = append(result, c.Paragraph)
		}
	}
	return result
}

// UnmarshalXML implements xml.Unmarshaler for CT_Document.
func (doc *CT_Document) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	doc.XMLName = start.Name
	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "body":
				doc.Body = &CT_Body{}
				if err := d.DecodeElement(doc.Body, &t); err != nil {
					return err
				}
			default:
				if err := d.Skip(); err != nil {
					return err
				}
			}
		case xml.EndElement:
			return nil
		}
	}
}

// MarshalXML implements xml.Marshaler for CT_Document.
func (doc *CT_Document) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name = doc.XMLName
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	if doc.Body != nil {
		bodyStart := xml.StartElement{Name: xml.Name{Space: Ns, Local: "body"}}
		if err := e.EncodeElement(doc.Body, bodyStart); err != nil {
			return err
		}
	}
	return e.EncodeToken(start.End())
}

// UnmarshalXML implements xml.Unmarshaler for CT_Body.
func (b *CT_Body) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	b.XMLName = start.Name
	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "p":
				var p CT_P
				if err := d.DecodeElement(&p, &t); err != nil {
					return err
				}
				b.Content = append(b.Content, BlockLevelContent{Paragraph: &p})
			case "tbl":
				var tbl CT_Tbl
				if err := d.DecodeElement(&tbl, &t); err != nil {
					return err
				}
				b.Content = append(b.Content, BlockLevelContent{Table: &tbl})
			case "sectPr":
				var raw xmlutil.RawXML
				if err := raw.UnmarshalXML(d, t); err != nil {
					return err
				}
				b.SectPr = &raw
			default:
				var raw xmlutil.RawXML
				if err := raw.UnmarshalXML(d, t); err != nil {
					return err
				}
				b.Content = append(b.Content, BlockLevelContent{Raw: &raw})
			}
		case xml.EndElement:
			return nil
		}
	}
}

// MarshalXML implements xml.Marshaler for CT_Body.
func (b *CT_Body) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name = b.XMLName
	if err := e.EncodeToken(start); err != nil {
		return err
	}

	for _, c := range b.Content {
		switch {
		case c.Paragraph != nil:
			pStart := xml.StartElement{Name: c.Paragraph.XMLName}
			if err := e.EncodeElement(c.Paragraph, pStart); err != nil {
				return err
			}
		case c.Table != nil:
			tblStart := xml.StartElement{Name: c.Table.XMLName}
			if err := e.EncodeElement(c.Table, tblStart); err != nil {
				return err
			}
		case c.Raw != nil:
			if err := c.Raw.MarshalXML(e, xml.StartElement{}); err != nil {
				return err
			}
		}
	}

	if b.SectPr != nil {
		if err := b.SectPr.MarshalXML(e, xml.StartElement{}); err != nil {
			return err
		}
	}

	return e.EncodeToken(start.End())
}
