package wml

import (
	"encoding/xml"
	"strconv"
	"strings"

	"github.com/ieshan/go-ooxml/xmlutil"
)

const NsRelationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

// CT_P represents <w:p> — a paragraph.
type CT_P struct {
	XMLName xml.Name
	PPr     *CT_PPr
	Content []InlineContent // Ordered children
}

// InlineContent holds one inline child of a paragraph. Exactly one field is non-nil.
type InlineContent struct {
	Run               *CT_R
	Hyperlink         *CT_Hyperlink
	CommentRangeStart *CT_MarkupRange
	CommentRangeEnd   *CT_MarkupRange
	Ins               *CT_RunTrackChange
	Del               *CT_RunTrackChange
	Raw               *xmlutil.RawXML
}

// CT_Hyperlink represents <w:hyperlink>
type CT_Hyperlink struct {
	XMLName xml.Name
	ID      string // r:id attr for relationship reference
	Anchor  string // w:anchor attr for internal links
	Runs    []*CT_R
	Raw     []xmlutil.RawXML
}

// CT_MarkupRange represents <w:commentRangeStart> or <w:commentRangeEnd>
type CT_MarkupRange struct {
	XMLName xml.Name
	ID      int
}

// CT_RunTrackChange represents <w:ins> or <w:del> — revision wrappers
type CT_RunTrackChange struct {
	XMLName xml.Name
	ID      int
	Author  string
	Date    string
	Runs    []*CT_R
	Raw     []xmlutil.RawXML
}

// CT_PPr represents <w:pPr> — paragraph properties
type CT_PPr struct {
	XMLName    xml.Name
	Style      *string     // <w:pStyle w:val="..."/>
	KeepNext   *bool       // <w:keepNext/>
	KeepLines  *bool       // <w:keepLines/>
	Spacing    *CT_Spacing // <w:spacing .../>
	Ind        *CT_Ind     // <w:ind .../>
	Alignment  *string     // <w:jc w:val="center"/>
	OutlineLvl *int        // <w:outlineLvl w:val="0"/>
	NumPr      *CT_NumPr
	Extra      []xmlutil.RawXML
}

// CT_NumPr represents <w:numPr> — numbering properties
type CT_NumPr struct {
	ILvl  *int // <w:ilvl w:val="0"/>
	NumID *int // <w:numId w:val="1"/>
}

// CT_Spacing represents <w:spacing> — paragraph spacing properties.
type CT_Spacing struct {
	Before   *string // w:before (twips)
	After    *string // w:after (twips)
	Line     *string // w:line
	LineRule *string // w:lineRule ("auto", "exact", "atLeast")
}

// CT_Ind represents <w:ind> — paragraph indentation.
type CT_Ind struct {
	Left  *string // w:left (twips)
	Right *string // w:right (twips)
}

// Text concatenates visible text from all runs (including inside ins/hyperlink).
// Deleted text inside <w:del> is excluded; use DelText for that.
func (p *CT_P) Text() string {
	var sb strings.Builder
	for _, c := range p.Content {
		switch {
		case c.Run != nil:
			sb.WriteString(c.Run.Text())
		case c.Ins != nil:
			for _, r := range c.Ins.Runs {
				sb.WriteString(r.Text())
			}
		case c.Del != nil:
			// Deleted text is not visible; skip.
		case c.Hyperlink != nil:
			for _, r := range c.Hyperlink.Runs {
				sb.WriteString(r.Text())
			}
		}
	}
	return sb.String()
}

// AddRun adds a new run with text to the paragraph.
func (p *CT_P) AddRun(text string) *CT_R {
	r := &CT_R{XMLName: xml.Name{Space: Ns, Local: "r"}}
	r.AddText(text)
	p.Content = append(p.Content, InlineContent{Run: r})
	return r
}

// UnmarshalXML implements xml.Unmarshaler for CT_P.
func (p *CT_P) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	p.XMLName = start.Name
	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "pPr":
				p.PPr = &CT_PPr{}
				if err := d.DecodeElement(p.PPr, &t); err != nil {
					return err
				}
			case "r":
				var r CT_R
				if err := d.DecodeElement(&r, &t); err != nil {
					return err
				}
				p.Content = append(p.Content, InlineContent{Run: &r})
			case "hyperlink":
				var h CT_Hyperlink
				if err := d.DecodeElement(&h, &t); err != nil {
					return err
				}
				p.Content = append(p.Content, InlineContent{Hyperlink: &h})
			case "commentRangeStart":
				var m CT_MarkupRange
				m.XMLName = t.Name
				for _, a := range t.Attr {
					if a.Name.Local == "id" {
						m.ID, _ = strconv.Atoi(a.Value)
					}
				}
				if err := d.Skip(); err != nil {
					return err
				}
				p.Content = append(p.Content, InlineContent{CommentRangeStart: &m})
			case "commentRangeEnd":
				var m CT_MarkupRange
				m.XMLName = t.Name
				for _, a := range t.Attr {
					if a.Name.Local == "id" {
						m.ID, _ = strconv.Atoi(a.Value)
					}
				}
				if err := d.Skip(); err != nil {
					return err
				}
				p.Content = append(p.Content, InlineContent{CommentRangeEnd: &m})
			case "ins":
				var tc CT_RunTrackChange
				if err := d.DecodeElement(&tc, &t); err != nil {
					return err
				}
				p.Content = append(p.Content, InlineContent{Ins: &tc})
			case "del":
				var tc CT_RunTrackChange
				if err := d.DecodeElement(&tc, &t); err != nil {
					return err
				}
				p.Content = append(p.Content, InlineContent{Del: &tc})
			default:
				var raw xmlutil.RawXML
				if err := raw.UnmarshalXML(d, t); err != nil {
					return err
				}
				p.Content = append(p.Content, InlineContent{Raw: &raw})
			}
		case xml.EndElement:
			return nil
		}
	}
}

// MarshalXML implements xml.Marshaler for CT_P.
func (p *CT_P) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name = p.XMLName
	if err := e.EncodeToken(start); err != nil {
		return err
	}

	if p.PPr != nil {
		pprStart := xml.StartElement{Name: xml.Name{Space: Ns, Local: "pPr"}}
		if err := e.EncodeElement(p.PPr, pprStart); err != nil {
			return err
		}
	}

	for _, c := range p.Content {
		switch {
		case c.Run != nil:
			runStart := xml.StartElement{Name: c.Run.XMLName}
			if err := e.EncodeElement(c.Run, runStart); err != nil {
				return err
			}
		case c.Hyperlink != nil:
			hlStart := xml.StartElement{Name: c.Hyperlink.XMLName}
			if err := e.EncodeElement(c.Hyperlink, hlStart); err != nil {
				return err
			}
		case c.CommentRangeStart != nil:
			s := xml.StartElement{
				Name: xml.Name{Space: Ns, Local: "commentRangeStart"},
				Attr: []xml.Attr{{Name: xml.Name{Space: Ns, Local: "id"}, Value: strconv.Itoa(c.CommentRangeStart.ID)}},
			}
			if err := e.EncodeToken(s); err != nil {
				return err
			}
			if err := e.EncodeToken(s.End()); err != nil {
				return err
			}
		case c.CommentRangeEnd != nil:
			s := xml.StartElement{
				Name: xml.Name{Space: Ns, Local: "commentRangeEnd"},
				Attr: []xml.Attr{{Name: xml.Name{Space: Ns, Local: "id"}, Value: strconv.Itoa(c.CommentRangeEnd.ID)}},
			}
			if err := e.EncodeToken(s); err != nil {
				return err
			}
			if err := e.EncodeToken(s.End()); err != nil {
				return err
			}
		case c.Ins != nil:
			insStart := xml.StartElement{Name: xml.Name{Space: Ns, Local: "ins"}}
			if err := e.EncodeElement(c.Ins, insStart); err != nil {
				return err
			}
		case c.Del != nil:
			delStart := xml.StartElement{Name: xml.Name{Space: Ns, Local: "del"}}
			if err := e.EncodeElement(c.Del, delStart); err != nil {
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

// UnmarshalXML implements xml.Unmarshaler for CT_PPr.
func (ppr *CT_PPr) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	ppr.XMLName = start.Name
	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "pStyle":
				v := getAttrVal(t.Attr)
				if v != "" {
					ppr.Style = &v
				}
				if err := d.Skip(); err != nil {
					return err
				}
			case "keepNext":
				v := true
				ppr.KeepNext = &v
				if err := d.Skip(); err != nil {
					return err
				}
			case "keepLines":
				v := true
				ppr.KeepLines = &v
				if err := d.Skip(); err != nil {
					return err
				}
			case "spacing":
				ppr.Spacing = &CT_Spacing{}
				for _, a := range t.Attr {
					switch a.Name.Local {
					case "before":
						v := a.Value
						ppr.Spacing.Before = &v
					case "after":
						v := a.Value
						ppr.Spacing.After = &v
					case "line":
						v := a.Value
						ppr.Spacing.Line = &v
					case "lineRule":
						v := a.Value
						ppr.Spacing.LineRule = &v
					}
				}
				if err := d.Skip(); err != nil {
					return err
				}
			case "ind":
				ppr.Ind = &CT_Ind{}
				for _, a := range t.Attr {
					switch a.Name.Local {
					case "left":
						v := a.Value
						ppr.Ind.Left = &v
					case "right":
						v := a.Value
						ppr.Ind.Right = &v
					}
				}
				if err := d.Skip(); err != nil {
					return err
				}
			case "jc":
				v := getAttrVal(t.Attr)
				if v != "" {
					ppr.Alignment = &v
				}
				if err := d.Skip(); err != nil {
					return err
				}
			case "outlineLvl":
				v := getAttrVal(t.Attr)
				if v != "" {
					n, _ := strconv.Atoi(v)
					ppr.OutlineLvl = &n
				}
				if err := d.Skip(); err != nil {
					return err
				}
			case "numPr":
				ppr.NumPr = &CT_NumPr{}
				if err := d.DecodeElement(ppr.NumPr, &t); err != nil {
					return err
				}
			default:
				var raw xmlutil.RawXML
				if err := raw.UnmarshalXML(d, t); err != nil {
					return err
				}
				ppr.Extra = append(ppr.Extra, raw)
			}
		case xml.EndElement:
			return nil
		}
	}
}

// MarshalXML implements xml.Marshaler for CT_PPr.
func (ppr *CT_PPr) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	if err := e.EncodeToken(start); err != nil {
		return err
	}

	if err := marshalValAttr(e, "pStyle", ppr.Style); err != nil {
		return err
	}
	if ppr.KeepNext != nil && *ppr.KeepNext {
		if err := marshalEmptyElement(e, "keepNext"); err != nil {
			return err
		}
	}
	if ppr.KeepLines != nil && *ppr.KeepLines {
		if err := marshalEmptyElement(e, "keepLines"); err != nil {
			return err
		}
	}
	if ppr.Spacing != nil {
		sp := xml.StartElement{Name: xml.Name{Space: Ns, Local: "spacing"}}
		if ppr.Spacing.Before != nil {
			sp.Attr = append(sp.Attr, xml.Attr{Name: xml.Name{Space: Ns, Local: "before"}, Value: *ppr.Spacing.Before})
		}
		if ppr.Spacing.After != nil {
			sp.Attr = append(sp.Attr, xml.Attr{Name: xml.Name{Space: Ns, Local: "after"}, Value: *ppr.Spacing.After})
		}
		if ppr.Spacing.Line != nil {
			sp.Attr = append(sp.Attr, xml.Attr{Name: xml.Name{Space: Ns, Local: "line"}, Value: *ppr.Spacing.Line})
		}
		if ppr.Spacing.LineRule != nil {
			sp.Attr = append(sp.Attr, xml.Attr{Name: xml.Name{Space: Ns, Local: "lineRule"}, Value: *ppr.Spacing.LineRule})
		}
		if err := e.EncodeToken(sp); err != nil {
			return err
		}
		if err := e.EncodeToken(sp.End()); err != nil {
			return err
		}
	}
	if ppr.Ind != nil {
		ind := xml.StartElement{Name: xml.Name{Space: Ns, Local: "ind"}}
		if ppr.Ind.Left != nil {
			ind.Attr = append(ind.Attr, xml.Attr{Name: xml.Name{Space: Ns, Local: "left"}, Value: *ppr.Ind.Left})
		}
		if ppr.Ind.Right != nil {
			ind.Attr = append(ind.Attr, xml.Attr{Name: xml.Name{Space: Ns, Local: "right"}, Value: *ppr.Ind.Right})
		}
		if err := e.EncodeToken(ind); err != nil {
			return err
		}
		if err := e.EncodeToken(ind.End()); err != nil {
			return err
		}
	}
	if err := marshalValAttr(e, "jc", ppr.Alignment); err != nil {
		return err
	}
	if ppr.OutlineLvl != nil {
		v := strconv.Itoa(*ppr.OutlineLvl)
		if err := marshalValAttr(e, "outlineLvl", &v); err != nil {
			return err
		}
	}
	if ppr.NumPr != nil {
		numStart := xml.StartElement{Name: xml.Name{Space: Ns, Local: "numPr"}}
		if err := e.EncodeElement(ppr.NumPr, numStart); err != nil {
			return err
		}
	}
	for i := range ppr.Extra {
		if err := ppr.Extra[i].MarshalXML(e, xml.StartElement{}); err != nil {
			return err
		}
	}

	return e.EncodeToken(start.End())
}

// UnmarshalXML implements xml.Unmarshaler for CT_NumPr.
func (np *CT_NumPr) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "ilvl":
				v := getAttrVal(t.Attr)
				if v != "" {
					n, _ := strconv.Atoi(v)
					np.ILvl = &n
				}
				if err := d.Skip(); err != nil {
					return err
				}
			case "numId":
				v := getAttrVal(t.Attr)
				if v != "" {
					n, _ := strconv.Atoi(v)
					np.NumID = &n
				}
				if err := d.Skip(); err != nil {
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

// MarshalXML implements xml.Marshaler for CT_NumPr.
func (np *CT_NumPr) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	if np.ILvl != nil {
		s := xml.StartElement{
			Name: xml.Name{Space: Ns, Local: "ilvl"},
			Attr: []xml.Attr{{Name: xml.Name{Space: Ns, Local: "val"}, Value: strconv.Itoa(*np.ILvl)}},
		}
		if err := e.EncodeToken(s); err != nil {
			return err
		}
		if err := e.EncodeToken(s.End()); err != nil {
			return err
		}
	}
	if np.NumID != nil {
		s := xml.StartElement{
			Name: xml.Name{Space: Ns, Local: "numId"},
			Attr: []xml.Attr{{Name: xml.Name{Space: Ns, Local: "val"}, Value: strconv.Itoa(*np.NumID)}},
		}
		if err := e.EncodeToken(s); err != nil {
			return err
		}
		if err := e.EncodeToken(s.End()); err != nil {
			return err
		}
	}
	return e.EncodeToken(start.End())
}

// UnmarshalXML implements xml.Unmarshaler for CT_Hyperlink.
func (h *CT_Hyperlink) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	h.XMLName = start.Name
	for _, a := range start.Attr {
		switch {
		case a.Name.Local == "id":
			h.ID = a.Value
		case a.Name.Local == "anchor":
			h.Anchor = a.Value
		}
	}
	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "r":
				var r CT_R
				if err := d.DecodeElement(&r, &t); err != nil {
					return err
				}
				h.Runs = append(h.Runs, &r)
			default:
				var raw xmlutil.RawXML
				if err := raw.UnmarshalXML(d, t); err != nil {
					return err
				}
				h.Raw = append(h.Raw, raw)
			}
		case xml.EndElement:
			return nil
		}
	}
}

// MarshalXML implements xml.Marshaler for CT_Hyperlink.
func (h *CT_Hyperlink) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name = h.XMLName
	if h.ID != "" {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Space: NsRelationships, Local: "id"}, Value: h.ID})
	}
	if h.Anchor != "" {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Space: Ns, Local: "anchor"}, Value: h.Anchor})
	}
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	for _, r := range h.Runs {
		runStart := xml.StartElement{Name: r.XMLName}
		if err := e.EncodeElement(r, runStart); err != nil {
			return err
		}
	}
	for i := range h.Raw {
		if err := h.Raw[i].MarshalXML(e, xml.StartElement{}); err != nil {
			return err
		}
	}
	return e.EncodeToken(start.End())
}

// UnmarshalXML implements xml.Unmarshaler for CT_RunTrackChange.
func (tc *CT_RunTrackChange) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	tc.XMLName = start.Name
	for _, a := range start.Attr {
		switch a.Name.Local {
		case "id":
			tc.ID, _ = strconv.Atoi(a.Value)
		case "author":
			tc.Author = a.Value
		case "date":
			tc.Date = a.Value
		}
	}
	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "r":
				var r CT_R
				if err := d.DecodeElement(&r, &t); err != nil {
					return err
				}
				tc.Runs = append(tc.Runs, &r)
			default:
				var raw xmlutil.RawXML
				if err := raw.UnmarshalXML(d, t); err != nil {
					return err
				}
				tc.Raw = append(tc.Raw, raw)
			}
		case xml.EndElement:
			return nil
		}
	}
}

// MarshalXML implements xml.Marshaler for CT_RunTrackChange.
func (tc *CT_RunTrackChange) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name = tc.XMLName
	start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Space: Ns, Local: "id"}, Value: strconv.Itoa(tc.ID)})
	start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Space: Ns, Local: "author"}, Value: tc.Author})
	if tc.Date != "" {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Space: Ns, Local: "date"}, Value: tc.Date})
	}
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	for _, r := range tc.Runs {
		runStart := xml.StartElement{Name: r.XMLName}
		if err := e.EncodeElement(r, runStart); err != nil {
			return err
		}
	}
	for i := range tc.Raw {
		if err := tc.Raw[i].MarshalXML(e, xml.StartElement{}); err != nil {
			return err
		}
	}
	return e.EncodeToken(start.End())
}
