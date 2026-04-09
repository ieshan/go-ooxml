package wml

import (
	"encoding/xml"
	"strconv"
	"strings"

	"github.com/ieshan/go-ooxml/xmlutil"
)

// CT_R represents <w:r> — a run of text with uniform formatting.
type CT_R struct {
	XMLName xml.Name
	RPr     *CT_RPr
	Content []RunContent // Ordered children: text, tabs, breaks, etc.
}

// RunContent holds one child of a run. Exactly one field is non-nil.
type RunContent struct {
	Text             *CT_Text
	Tab              *CT_Tab
	Break            *CT_Break
	DelText          *CT_Text
	CommentReference *CT_CommentReference
	AnnotationRef    *CT_AnnotationRef
	Raw              *xmlutil.RawXML
}

// CT_Text represents <w:t>
type CT_Text struct {
	Space string `xml:"http://www.w3.org/XML/1998/namespace space,attr,omitempty"`
	Value string `xml:",chardata"`
}

// CT_Tab represents <w:tab/>
type CT_Tab struct{}

// CT_Break represents <w:br/> with optional type attr
type CT_Break struct {
	Type string `xml:"type,attr,omitempty"`
}

// CT_AnnotationRef represents <w:annotationRef/> — a back-reference from a
// comment body paragraph to the comment's anchor in the document.
type CT_AnnotationRef struct{}

// CT_RPr represents <w:rPr> — run properties (formatting).
type CT_RPr struct {
	Bold         *bool            // <w:b/> or <w:b w:val="..."/>
	Italic       *bool            // <w:i/>
	Underline    *string          // <w:u w:val="single"/>
	Strike       *bool            // <w:strike/>
	FontName     *string          // <w:rFonts w:ascii="..."/>
	FontHAnsi    *string          // <w:rFonts w:hAnsi="..."/>
	FontEastAsia *string          // <w:rFonts w:eastAsia="..."/>
	FontCS       *string          // <w:rFonts w:cs="..."/>
	FontSize     *string          // <w:sz w:val="24"/> (half-points)
	FontSizeCS   *string          // <w:szCs w:val="24"/> (half-points, complex script)
	Color        *string          // <w:color w:val="FF0000"/>
	ThemeColor   *string          // <w:color w:themeColor="..."/>
	ThemeShade   *string          // <w:color w:themeShade="..."/>
	ThemeTint    *string          // <w:color w:themeTint="..."/>
	RunStyle     *string          // <w:rStyle w:val="StyleName"/>
	Extra        []xmlutil.RawXML // Unknown elements preserved
}

// Text concatenates the visible text content of the run (w:t elements).
// Deleted text (w:delText) is excluded — use DelText() for that.
func (r *CT_R) Text() string {
	var sb strings.Builder
	for _, c := range r.Content {
		if c.Text != nil {
			sb.WriteString(c.Text.Value)
		} else if c.Tab != nil {
			sb.WriteByte('\t')
		} else if c.Break != nil {
			sb.WriteByte('\n')
		}
	}
	return sb.String()
}

// DelText concatenates the deleted text content of the run (w:delText elements).
func (r *CT_R) DelText() string {
	var sb strings.Builder
	for _, c := range r.Content {
		if c.DelText != nil {
			sb.WriteString(c.DelText.Value)
		}
	}
	return sb.String()
}

// AddText adds a CT_Text child to the run.
func (r *CT_R) AddText(text string) {
	space := ""
	if len(text) > 0 && (text[0] == ' ' || text[len(text)-1] == ' ') {
		space = "preserve"
	}
	r.Content = append(r.Content, RunContent{
		Text: &CT_Text{Space: space, Value: text},
	})
}

// UnmarshalXML implements xml.Unmarshaler for CT_R.
func (r *CT_R) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	r.XMLName = start.Name
	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "rPr":
				r.RPr = &CT_RPr{}
				if err := d.DecodeElement(r.RPr, &t); err != nil {
					return err
				}
			case "t":
				var ct CT_Text
				if err := d.DecodeElement(&ct, &t); err != nil {
					return err
				}
				r.Content = append(r.Content, RunContent{Text: &ct})
			case "delText":
				var ct CT_Text
				if err := d.DecodeElement(&ct, &t); err != nil {
					return err
				}
				r.Content = append(r.Content, RunContent{DelText: &ct})
			case "tab":
				if err := d.Skip(); err != nil {
					return err
				}
				r.Content = append(r.Content, RunContent{Tab: &CT_Tab{}})
			case "br":
				var br CT_Break
				if err := d.DecodeElement(&br, &t); err != nil {
					return err
				}
				r.Content = append(r.Content, RunContent{Break: &br})
			case "commentReference":
				var cr CT_CommentReference
				for _, a := range t.Attr {
					if a.Name.Local == "id" {
						cr.ID, _ = strconv.Atoi(a.Value)
					}
				}
				if err := d.Skip(); err != nil {
					return err
				}
				r.Content = append(r.Content, RunContent{CommentReference: &cr})
			case "annotationRef":
				if err := d.Skip(); err != nil {
					return err
				}
				r.Content = append(r.Content, RunContent{AnnotationRef: &CT_AnnotationRef{}})
			default:
				var raw xmlutil.RawXML
				if err := raw.UnmarshalXML(d, t); err != nil {
					return err
				}
				r.Content = append(r.Content, RunContent{Raw: &raw})
			}
		case xml.EndElement:
			return nil
		}
	}
}

// MarshalXML implements xml.Marshaler for CT_R.
func (r *CT_R) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name = r.XMLName
	if err := e.EncodeToken(start); err != nil {
		return err
	}

	if r.RPr != nil {
		rprStart := xml.StartElement{Name: xml.Name{Space: Ns, Local: "rPr"}}
		if err := e.EncodeElement(r.RPr, rprStart); err != nil {
			return err
		}
	}

	for _, c := range r.Content {
		switch {
		case c.Text != nil:
			tStart := xml.StartElement{Name: xml.Name{Space: Ns, Local: "t"}}
			if err := e.EncodeElement(c.Text, tStart); err != nil {
				return err
			}
		case c.DelText != nil:
			tStart := xml.StartElement{Name: xml.Name{Space: Ns, Local: "delText"}}
			if err := e.EncodeElement(c.DelText, tStart); err != nil {
				return err
			}
		case c.Tab != nil:
			tStart := xml.StartElement{Name: xml.Name{Space: Ns, Local: "tab"}}
			if err := e.EncodeToken(tStart); err != nil {
				return err
			}
			if err := e.EncodeToken(tStart.End()); err != nil {
				return err
			}
		case c.Break != nil:
			brStart := xml.StartElement{Name: xml.Name{Space: Ns, Local: "br"}}
			if c.Break.Type != "" {
				brStart.Attr = append(brStart.Attr, xml.Attr{Name: xml.Name{Space: Ns, Local: "type"}, Value: c.Break.Type})
			}
			if err := e.EncodeToken(brStart); err != nil {
				return err
			}
			if err := e.EncodeToken(brStart.End()); err != nil {
				return err
			}
		case c.CommentReference != nil:
			s := xml.StartElement{
				Name: xml.Name{Space: Ns, Local: "commentReference"},
				Attr: []xml.Attr{{Name: xml.Name{Space: Ns, Local: "id"}, Value: strconv.Itoa(c.CommentReference.ID)}},
			}
			if err := e.EncodeToken(s); err != nil {
				return err
			}
			if err := e.EncodeToken(s.End()); err != nil {
				return err
			}
		case c.AnnotationRef != nil:
			s := xml.StartElement{Name: xml.Name{Space: Ns, Local: "annotationRef"}}
			if err := e.EncodeToken(s); err != nil {
				return err
			}
			if err := e.EncodeToken(s.End()); err != nil {
				return err
			}
		case c.Raw != nil:
			if err := c.Raw.MarshalXML(e, xml.StartElement{}); err != nil {
				return err
			}
		}
	}

	return e.EncodeToken(start.End())
}

// parseBoolVal parses the w:val attribute for boolean toggle elements.
// absent or "" or "true" or "1" → true; "false" or "0" → false.
func parseBoolVal(attrs []xml.Attr) bool {
	for _, a := range attrs {
		if a.Name.Local == "val" {
			switch a.Value {
			case "false", "0":
				return false
			default:
				return true
			}
		}
	}
	return true // absent val means true
}

// getAttrVal returns the value of the w:val attribute, or empty string if absent.
func getAttrVal(attrs []xml.Attr) string {
	for _, a := range attrs {
		if a.Name.Local == "val" {
			return a.Value
		}
	}
	return ""
}

// getAttr returns the value of the named attribute, or empty string if absent.
func getAttr(attrs []xml.Attr, local string) string {
	for _, a := range attrs {
		if a.Name.Local == local {
			return a.Value
		}
	}
	return ""
}

// UnmarshalXML implements xml.Unmarshaler for CT_RPr.
func (rpr *CT_RPr) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "b":
				v := parseBoolVal(t.Attr)
				rpr.Bold = &v
				if err := d.Skip(); err != nil {
					return err
				}
			case "i":
				v := parseBoolVal(t.Attr)
				rpr.Italic = &v
				if err := d.Skip(); err != nil {
					return err
				}
			case "strike":
				v := parseBoolVal(t.Attr)
				rpr.Strike = &v
				if err := d.Skip(); err != nil {
					return err
				}
			case "u":
				v := getAttrVal(t.Attr)
				if v != "" {
					rpr.Underline = &v
				}
				if err := d.Skip(); err != nil {
					return err
				}
			case "rFonts":
				if v := getAttr(t.Attr, "ascii"); v != "" {
					rpr.FontName = &v
				}
				if v := getAttr(t.Attr, "hAnsi"); v != "" {
					rpr.FontHAnsi = &v
				}
				if v := getAttr(t.Attr, "eastAsia"); v != "" {
					rpr.FontEastAsia = &v
				}
				if v := getAttr(t.Attr, "cs"); v != "" {
					rpr.FontCS = &v
				}
				if err := d.Skip(); err != nil {
					return err
				}
			case "sz":
				v := getAttrVal(t.Attr)
				if v != "" {
					rpr.FontSize = &v
				}
				if err := d.Skip(); err != nil {
					return err
				}
			case "szCs":
				v := getAttrVal(t.Attr)
				if v != "" {
					rpr.FontSizeCS = &v
				}
				if err := d.Skip(); err != nil {
					return err
				}
			case "color":
				if v := getAttrVal(t.Attr); v != "" {
					rpr.Color = &v
				}
				if v := getAttr(t.Attr, "themeColor"); v != "" {
					rpr.ThemeColor = &v
				}
				if v := getAttr(t.Attr, "themeShade"); v != "" {
					rpr.ThemeShade = &v
				}
				if v := getAttr(t.Attr, "themeTint"); v != "" {
					rpr.ThemeTint = &v
				}
				if err := d.Skip(); err != nil {
					return err
				}
			case "rStyle":
				v := getAttrVal(t.Attr)
				if v != "" {
					rpr.RunStyle = &v
				}
				if err := d.Skip(); err != nil {
					return err
				}
			default:
				var raw xmlutil.RawXML
				if err := raw.UnmarshalXML(d, t); err != nil {
					return err
				}
				rpr.Extra = append(rpr.Extra, raw)
			}
		case xml.EndElement:
			return nil
		}
	}
}

// MarshalXML implements xml.Marshaler for CT_RPr.
// Element ordering follows the ECMA-376 XSD sequence for CT_RPr:
// rStyle, rFonts, b, i, strike, u, color, sz, Extra...
func (rpr *CT_RPr) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	if err := e.EncodeToken(start); err != nil {
		return err
	}

	// rStyle must be first per ECMA-376 17.3.2.27.
	if err := marshalValAttr(e, "rStyle", rpr.RunStyle); err != nil {
		return err
	}
	if err := marshalRFonts(e, rpr); err != nil {
		return err
	}
	if err := marshalBoolToggle(e, "b", rpr.Bold); err != nil {
		return err
	}
	if err := marshalBoolToggle(e, "i", rpr.Italic); err != nil {
		return err
	}
	if err := marshalBoolToggle(e, "strike", rpr.Strike); err != nil {
		return err
	}
	if err := marshalValAttr(e, "u", rpr.Underline); err != nil {
		return err
	}
	if err := marshalColor(e, rpr); err != nil {
		return err
	}
	if err := marshalValAttr(e, "sz", rpr.FontSize); err != nil {
		return err
	}
	if err := marshalValAttr(e, "szCs", rpr.FontSizeCS); err != nil {
		return err
	}

	for i := range rpr.Extra {
		if err := rpr.Extra[i].MarshalXML(e, xml.StartElement{}); err != nil {
			return err
		}
	}

	return e.EncodeToken(start.End())
}

// marshalRFonts emits <w:rFonts> with all four font attributes when at least
// one is set. When only w:ascii is set, w:hAnsi defaults to the same value
// to cover Western characters.
func marshalRFonts(e *xml.Encoder, rpr *CT_RPr) error {
	if rpr.FontName == nil && rpr.FontHAnsi == nil && rpr.FontEastAsia == nil && rpr.FontCS == nil {
		return nil
	}
	s := xml.StartElement{Name: xml.Name{Space: Ns, Local: "rFonts"}}
	if rpr.FontName != nil {
		s.Attr = append(s.Attr, xml.Attr{Name: xml.Name{Space: Ns, Local: "ascii"}, Value: *rpr.FontName})
	}
	hAnsi := rpr.FontHAnsi
	if hAnsi == nil && rpr.FontName != nil {
		hAnsi = rpr.FontName
	}
	if hAnsi != nil {
		s.Attr = append(s.Attr, xml.Attr{Name: xml.Name{Space: Ns, Local: "hAnsi"}, Value: *hAnsi})
	}
	if rpr.FontEastAsia != nil {
		s.Attr = append(s.Attr, xml.Attr{Name: xml.Name{Space: Ns, Local: "eastAsia"}, Value: *rpr.FontEastAsia})
	}
	if rpr.FontCS != nil {
		s.Attr = append(s.Attr, xml.Attr{Name: xml.Name{Space: Ns, Local: "cs"}, Value: *rpr.FontCS})
	}
	if err := e.EncodeToken(s); err != nil {
		return err
	}
	return e.EncodeToken(s.End())
}

// marshalColor emits <w:color> with val and optional theme attributes.
func marshalColor(e *xml.Encoder, rpr *CT_RPr) error {
	if rpr.Color == nil && rpr.ThemeColor == nil {
		return nil
	}
	s := xml.StartElement{Name: xml.Name{Space: Ns, Local: "color"}}
	if rpr.Color != nil {
		s.Attr = append(s.Attr, xml.Attr{Name: xml.Name{Space: Ns, Local: "val"}, Value: *rpr.Color})
	}
	if rpr.ThemeColor != nil {
		s.Attr = append(s.Attr, xml.Attr{Name: xml.Name{Space: Ns, Local: "themeColor"}, Value: *rpr.ThemeColor})
	}
	if rpr.ThemeShade != nil {
		s.Attr = append(s.Attr, xml.Attr{Name: xml.Name{Space: Ns, Local: "themeShade"}, Value: *rpr.ThemeShade})
	}
	if rpr.ThemeTint != nil {
		s.Attr = append(s.Attr, xml.Attr{Name: xml.Name{Space: Ns, Local: "themeTint"}, Value: *rpr.ThemeTint})
	}
	if err := e.EncodeToken(s); err != nil {
		return err
	}
	return e.EncodeToken(s.End())
}

// marshalBoolToggle marshals a *bool as an OOXML toggle element.
// nil → omit; true → <w:name/>; false → <w:name w:val="false"/>
func marshalBoolToggle(e *xml.Encoder, local string, val *bool) error {
	if val == nil {
		return nil
	}
	s := xml.StartElement{Name: xml.Name{Space: Ns, Local: local}}
	if !*val {
		s.Attr = []xml.Attr{{Name: xml.Name{Space: Ns, Local: "val"}, Value: "false"}}
	}
	if err := e.EncodeToken(s); err != nil {
		return err
	}
	return e.EncodeToken(s.End())
}

// marshalValAttr marshals a *string as <w:name w:val="..."/>.
func marshalValAttr(e *xml.Encoder, local string, val *string) error {
	if val == nil {
		return nil
	}
	s := xml.StartElement{
		Name: xml.Name{Space: Ns, Local: local},
		Attr: []xml.Attr{{Name: xml.Name{Space: Ns, Local: "val"}, Value: *val}},
	}
	if err := e.EncodeToken(s); err != nil {
		return err
	}
	return e.EncodeToken(s.End())
}

// marshalEmptyElement writes an empty element like <w:keepNext/>.
func marshalEmptyElement(e *xml.Encoder, local string) error {
	s := xml.StartElement{Name: xml.Name{Space: Ns, Local: local}}
	if err := e.EncodeToken(s); err != nil {
		return err
	}
	return e.EncodeToken(s.End())
}
