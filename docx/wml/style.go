package wml

import (
	"encoding/xml"

	"github.com/ieshan/go-ooxml/xmlutil"
)

// CT_Styles represents <w:styles> — root of styles.xml.
type CT_Styles struct {
	XMLName xml.Name
	Styles  []*CT_Style
	Extra   []xmlutil.RawXML // docDefaults, latentStyles, etc.
}

// CT_Style represents a single <w:style>.
type CT_Style struct {
	XMLName xml.Name
	Type    string  // paragraph, character, table, numbering
	StyleID string  // e.g., "Heading1"
	Default *bool   // w:default attr
	Name    string  // <w:name w:val="..."/>
	BasedOn *string // <w:basedOn w:val="..."/>
	Next    *string // <w:next w:val="..."/>
	PPr     *CT_PPr // paragraph properties
	RPr     *CT_RPr // run properties
	Extra   []xmlutil.RawXML
}

// UnmarshalXML implements xml.Unmarshaler for CT_Styles.
func (s *CT_Styles) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	s.XMLName = start.Name
	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "style":
				var style CT_Style
				if err := d.DecodeElement(&style, &t); err != nil {
					return err
				}
				s.Styles = append(s.Styles, &style)
			default:
				var raw xmlutil.RawXML
				if err := raw.UnmarshalXML(d, t); err != nil {
					return err
				}
				s.Extra = append(s.Extra, raw)
			}
		case xml.EndElement:
			return nil
		}
	}
}

// MarshalXML implements xml.Marshaler for CT_Styles.
func (s *CT_Styles) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name = s.XMLName
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	for i := range s.Extra {
		if err := s.Extra[i].MarshalXML(e, xml.StartElement{}); err != nil {
			return err
		}
	}
	for _, style := range s.Styles {
		styleStart := xml.StartElement{Name: style.XMLName}
		if err := e.EncodeElement(style, styleStart); err != nil {
			return err
		}
	}
	return e.EncodeToken(start.End())
}

// UnmarshalXML implements xml.Unmarshaler for CT_Style.
func (s *CT_Style) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	s.XMLName = start.Name
	// Parse attributes
	for _, a := range start.Attr {
		switch a.Name.Local {
		case "type":
			s.Type = a.Value
		case "styleId":
			s.StyleID = a.Value
		case "default":
			s.Default = new(a.Value == "1" || a.Value == "true")
		}
	}
	// Parse children
	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "name":
				s.Name = getAttrVal(t.Attr)
				if err := d.Skip(); err != nil {
					return err
				}
			case "basedOn":
				s.BasedOn = new(getAttrVal(t.Attr))
				if err := d.Skip(); err != nil {
					return err
				}
			case "next":
				s.Next = new(getAttrVal(t.Attr))
				if err := d.Skip(); err != nil {
					return err
				}
			case "pPr":
				s.PPr = &CT_PPr{}
				if err := d.DecodeElement(s.PPr, &t); err != nil {
					return err
				}
			case "rPr":
				s.RPr = &CT_RPr{}
				if err := d.DecodeElement(s.RPr, &t); err != nil {
					return err
				}
			default:
				var raw xmlutil.RawXML
				if err := raw.UnmarshalXML(d, t); err != nil {
					return err
				}
				s.Extra = append(s.Extra, raw)
			}
		case xml.EndElement:
			return nil
		}
	}
}

// MarshalXML implements xml.Marshaler for CT_Style.
func (s *CT_Style) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name = s.XMLName
	if s.Type != "" {
		start.Attr = append(start.Attr, xml.Attr{
			Name:  xml.Name{Space: Ns, Local: "type"},
			Value: s.Type,
		})
	}
	if s.StyleID != "" {
		start.Attr = append(start.Attr, xml.Attr{
			Name:  xml.Name{Space: Ns, Local: "styleId"},
			Value: s.StyleID,
		})
	}
	if s.Default != nil && *s.Default {
		start.Attr = append(start.Attr, xml.Attr{
			Name:  xml.Name{Space: Ns, Local: "default"},
			Value: "1",
		})
	}
	if err := e.EncodeToken(start); err != nil {
		return err
	}

	if s.Name != "" {
		if err := marshalValAttr(e, "name", &s.Name); err != nil {
			return err
		}
	}
	if err := marshalValAttr(e, "basedOn", s.BasedOn); err != nil {
		return err
	}
	if err := marshalValAttr(e, "next", s.Next); err != nil {
		return err
	}
	if s.PPr != nil {
		pprStart := xml.StartElement{Name: xml.Name{Space: Ns, Local: "pPr"}}
		if err := e.EncodeElement(s.PPr, pprStart); err != nil {
			return err
		}
	}
	if s.RPr != nil {
		rprStart := xml.StartElement{Name: xml.Name{Space: Ns, Local: "rPr"}}
		if err := e.EncodeElement(s.RPr, rprStart); err != nil {
			return err
		}
	}
	for i := range s.Extra {
		if err := s.Extra[i].MarshalXML(e, xml.StartElement{}); err != nil {
			return err
		}
	}

	return e.EncodeToken(start.End())
}
