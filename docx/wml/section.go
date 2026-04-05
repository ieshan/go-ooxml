package wml

import (
	"encoding/xml"
	"strconv"

	"github.com/ieshan/go-ooxml/xmlutil"
)

// CT_HdrFtrRef represents <w:headerReference> or <w:footerReference>.
type CT_HdrFtrRef struct {
	XMLName xml.Name
	Type    string // "default", "first", "even"
	ID      string // r:id relationship reference
}

// CT_SectPr represents <w:sectPr> — section properties.
type CT_SectPr struct {
	XMLName    xml.Name
	HeaderRefs []CT_HdrFtrRef // <w:headerReference>
	FooterRefs []CT_HdrFtrRef // <w:footerReference>
	PgSz       *CT_PgSz
	PgMar      *CT_PgMar
	Cols       *CT_Columns
	Type       *string          // <w:type w:val="continuous"/>
	Extra      []xmlutil.RawXML // other preserved elements
}

// CT_PgSz represents <w:pgSz>.
type CT_PgSz struct {
	W      string `xml:"w,attr,omitempty"`
	H      string `xml:"h,attr,omitempty"`
	Orient string `xml:"orient,attr,omitempty"`
}

// CT_PgMar represents <w:pgMar>.
type CT_PgMar struct {
	Top    string `xml:"top,attr,omitempty"`
	Right  string `xml:"right,attr,omitempty"`
	Bottom string `xml:"bottom,attr,omitempty"`
	Left   string `xml:"left,attr,omitempty"`
	Header string `xml:"header,attr,omitempty"`
	Footer string `xml:"footer,attr,omitempty"`
	Gutter string `xml:"gutter,attr,omitempty"`
}

// CT_Columns represents <w:cols>.
type CT_Columns struct {
	Num        *int    // w:num attr
	Space      *string // w:space attr
	EqualWidth *bool   // w:equalWidth attr
	Sep        *bool   // w:sep attr
	Cols       []*CT_Column
}

// CT_Column represents <w:col>.
type CT_Column struct {
	W     string `xml:"w,attr,omitempty"`
	Space string `xml:"space,attr,omitempty"`
}

// UnmarshalXML implements xml.Unmarshaler for CT_SectPr.
func (sp *CT_SectPr) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	sp.XMLName = start.Name
	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "pgSz":
				sp.PgSz = &CT_PgSz{}
				if err := d.DecodeElement(sp.PgSz, &t); err != nil {
					return err
				}
			case "pgMar":
				sp.PgMar = &CT_PgMar{}
				if err := d.DecodeElement(sp.PgMar, &t); err != nil {
					return err
				}
			case "cols":
				sp.Cols = &CT_Columns{}
				if err := d.DecodeElement(sp.Cols, &t); err != nil {
					return err
				}
			case "headerReference":
				ref := CT_HdrFtrRef{XMLName: t.Name}
				for _, a := range t.Attr {
					switch a.Name.Local {
					case "type":
						ref.Type = a.Value
					case "id":
						ref.ID = a.Value
					}
				}
				if err := d.Skip(); err != nil {
					return err
				}
				sp.HeaderRefs = append(sp.HeaderRefs, ref)
			case "footerReference":
				ref := CT_HdrFtrRef{XMLName: t.Name}
				for _, a := range t.Attr {
					switch a.Name.Local {
					case "type":
						ref.Type = a.Value
					case "id":
						ref.ID = a.Value
					}
				}
				if err := d.Skip(); err != nil {
					return err
				}
				sp.FooterRefs = append(sp.FooterRefs, ref)
			case "type":
				v := getAttrVal(t.Attr)
				sp.Type = &v
				if err := d.Skip(); err != nil {
					return err
				}
			default:
				var raw xmlutil.RawXML
				if err := raw.UnmarshalXML(d, t); err != nil {
					return err
				}
				sp.Extra = append(sp.Extra, raw)
			}
		case xml.EndElement:
			return nil
		}
	}
}

// MarshalXML implements xml.Marshaler for CT_SectPr.
// Per ECMA-376, section property children follow a specific order:
// headerReference*, footerReference*, type, pgSz, pgMar, ..., cols, ...
// Extra elements (which typically include header/footer refs) are emitted
// first to maintain schema-compliant ordering.
func (sp *CT_SectPr) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name = sp.XMLName
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	// Per ECMA-376: headerReference*, footerReference* come first.
	for _, ref := range sp.HeaderRefs {
		if err := marshalHdrFtrRef(e, ref); err != nil {
			return err
		}
	}
	for _, ref := range sp.FooterRefs {
		if err := marshalHdrFtrRef(e, ref); err != nil {
			return err
		}
	}
	// Emit other preserved elements.
	for i := range sp.Extra {
		if err := sp.Extra[i].MarshalXML(e, xml.StartElement{}); err != nil {
			return err
		}
	}
	if sp.Type != nil {
		if err := marshalValAttr(e, "type", sp.Type); err != nil {
			return err
		}
	}
	if sp.PgSz != nil {
		if err := marshalPgSz(e, sp.PgSz); err != nil {
			return err
		}
	}
	if sp.PgMar != nil {
		if err := marshalPgMar(e, sp.PgMar); err != nil {
			return err
		}
	}
	if sp.Cols != nil {
		colsStart := xml.StartElement{Name: xml.Name{Space: Ns, Local: "cols"}}
		if err := e.EncodeElement(sp.Cols, colsStart); err != nil {
			return err
		}
	}
	return e.EncodeToken(start.End())
}

// UnmarshalXML implements xml.Unmarshaler for CT_Columns.
func (c *CT_Columns) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	// Parse attributes
	for _, a := range start.Attr {
		switch a.Name.Local {
		case "num":
			n, _ := strconv.Atoi(a.Value)
			c.Num = &n
		case "space":
			v := a.Value
			c.Space = &v
		case "equalWidth":
			v := a.Value == "1" || a.Value == "true"
			c.EqualWidth = &v
		case "sep":
			v := a.Value == "1" || a.Value == "true"
			c.Sep = &v
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
			if t.Name.Local == "col" {
				col := &CT_Column{}
				if err := d.DecodeElement(col, &t); err != nil {
					return err
				}
				c.Cols = append(c.Cols, col)
			} else {
				if err := d.Skip(); err != nil {
					return err
				}
			}
		case xml.EndElement:
			return nil
		}
	}
}

// MarshalXML implements xml.Marshaler for CT_Columns.
func (c *CT_Columns) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	if c.Num != nil {
		start.Attr = append(start.Attr, xml.Attr{
			Name:  xml.Name{Space: Ns, Local: "num"},
			Value: strconv.Itoa(*c.Num),
		})
	}
	if c.Space != nil {
		start.Attr = append(start.Attr, xml.Attr{
			Name:  xml.Name{Space: Ns, Local: "space"},
			Value: *c.Space,
		})
	}
	if c.EqualWidth != nil {
		v := "0"
		if *c.EqualWidth {
			v = "1"
		}
		start.Attr = append(start.Attr, xml.Attr{
			Name:  xml.Name{Space: Ns, Local: "equalWidth"},
			Value: v,
		})
	}
	if c.Sep != nil {
		v := "0"
		if *c.Sep {
			v = "1"
		}
		start.Attr = append(start.Attr, xml.Attr{
			Name:  xml.Name{Space: Ns, Local: "sep"},
			Value: v,
		})
	}
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	for _, col := range c.Cols {
		if err := marshalColumn(e, col); err != nil {
			return err
		}
	}
	return e.EncodeToken(start.End())
}

// marshalPgSz emits <w:pgSz> with properly namespaced attributes.
// Go's encoding/xml generates broken namespace declarations for struct-tag-based
// attributes with namespaces, so we build attributes manually.
func marshalPgSz(e *xml.Encoder, p *CT_PgSz) error {
	s := xml.StartElement{Name: xml.Name{Space: Ns, Local: "pgSz"}}
	if p.W != "" {
		s.Attr = append(s.Attr, xml.Attr{Name: xml.Name{Space: Ns, Local: "w"}, Value: p.W})
	}
	if p.H != "" {
		s.Attr = append(s.Attr, xml.Attr{Name: xml.Name{Space: Ns, Local: "h"}, Value: p.H})
	}
	if p.Orient != "" {
		s.Attr = append(s.Attr, xml.Attr{Name: xml.Name{Space: Ns, Local: "orient"}, Value: p.Orient})
	}
	if err := e.EncodeToken(s); err != nil {
		return err
	}
	return e.EncodeToken(s.End())
}

// marshalPgMar emits <w:pgMar> with properly namespaced attributes.
func marshalPgMar(e *xml.Encoder, p *CT_PgMar) error {
	s := xml.StartElement{Name: xml.Name{Space: Ns, Local: "pgMar"}}
	for _, pair := range []struct{ local, val string }{
		{"top", p.Top}, {"right", p.Right}, {"bottom", p.Bottom}, {"left", p.Left},
		{"header", p.Header}, {"footer", p.Footer}, {"gutter", p.Gutter},
	} {
		if pair.val != "" {
			s.Attr = append(s.Attr, xml.Attr{Name: xml.Name{Space: Ns, Local: pair.local}, Value: pair.val})
		}
	}
	if err := e.EncodeToken(s); err != nil {
		return err
	}
	return e.EncodeToken(s.End())
}

// marshalHdrFtrRef emits a <w:headerReference> or <w:footerReference> element.
func marshalHdrFtrRef(e *xml.Encoder, ref CT_HdrFtrRef) error {
	s := xml.StartElement{Name: ref.XMLName}
	if ref.Type != "" {
		s.Attr = append(s.Attr, xml.Attr{Name: xml.Name{Space: Ns, Local: "type"}, Value: ref.Type})
	}
	if ref.ID != "" {
		s.Attr = append(s.Attr, xml.Attr{Name: xml.Name{Space: NsRelationships, Local: "id"}, Value: ref.ID})
	}
	if err := e.EncodeToken(s); err != nil {
		return err
	}
	return e.EncodeToken(s.End())
}

// marshalColumn emits <w:col> with properly namespaced attributes.
func marshalColumn(e *xml.Encoder, col *CT_Column) error {
	s := xml.StartElement{Name: xml.Name{Space: Ns, Local: "col"}}
	if col.W != "" {
		s.Attr = append(s.Attr, xml.Attr{Name: xml.Name{Space: Ns, Local: "w"}, Value: col.W})
	}
	if col.Space != "" {
		s.Attr = append(s.Attr, xml.Attr{Name: xml.Name{Space: Ns, Local: "space"}, Value: col.Space})
	}
	if err := e.EncodeToken(s); err != nil {
		return err
	}
	return e.EncodeToken(s.End())
}
