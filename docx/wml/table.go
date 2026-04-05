package wml

import (
	"encoding/xml"
	"strconv"

	"github.com/ieshan/go-ooxml/xmlutil"
)

// CT_Tbl represents <w:tbl> — a table.
type CT_Tbl struct {
	XMLName xml.Name
	TblPr   *CT_TblPr
	TblGrid *CT_TblGrid
	Rows    []*CT_Row
	Raw     []xmlutil.RawXML
}

// CT_TblPr represents <w:tblPr>
type CT_TblPr struct {
	Style *string      // <w:tblStyle w:val="..."/>
	Width *CT_TblWidth // <w:tblW w:w="..." w:type="..."/>
	Jc    *string      // <w:jc w:val="center"/>
	Extra []xmlutil.RawXML
}

// CT_TblWidth represents table/cell width.
type CT_TblWidth struct {
	W    string `xml:"w,attr"`
	Type string `xml:"type,attr"`
}

// CT_TblGrid represents <w:tblGrid>
type CT_TblGrid struct {
	Cols []*CT_TblGridCol
}

// CT_TblGridCol represents <w:gridCol>
type CT_TblGridCol struct {
	W string `xml:"w,attr"`
}

// CT_Row represents <w:tr>
type CT_Row struct {
	XMLName xml.Name
	TrPr    *xmlutil.RawXML // Preserved as raw
	Cells   []*CT_Tc
	Raw     []xmlutil.RawXML
}

// CT_Tc represents <w:tc> — a table cell (block container).
type CT_Tc struct {
	XMLName xml.Name
	TcPr    *CT_TcPr
	Content []BlockLevelContent // Paragraphs, nested tables
}

// CT_TcPr represents <w:tcPr>
type CT_TcPr struct {
	Width    *CT_TblWidth
	VMerge   *CT_VMerge
	GridSpan *int
	Shading  *CT_Shd
	Extra    []xmlutil.RawXML
}

// CT_VMerge represents <w:vMerge>
type CT_VMerge struct {
	Val string `xml:"val,attr,omitempty"`
}

// CT_Shd represents <w:shd>
type CT_Shd struct {
	Val   string `xml:"val,attr,omitempty"`
	Color string `xml:"color,attr,omitempty"`
	Fill  string `xml:"fill,attr,omitempty"`
}

// --- CT_Tbl ---

// UnmarshalXML implements xml.Unmarshaler for CT_Tbl.
func (tbl *CT_Tbl) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	tbl.XMLName = start.Name
	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "tblPr":
				tbl.TblPr = &CT_TblPr{}
				if err := d.DecodeElement(tbl.TblPr, &t); err != nil {
					return err
				}
			case "tblGrid":
				tbl.TblGrid = &CT_TblGrid{}
				if err := d.DecodeElement(tbl.TblGrid, &t); err != nil {
					return err
				}
			case "tr":
				var row CT_Row
				if err := d.DecodeElement(&row, &t); err != nil {
					return err
				}
				tbl.Rows = append(tbl.Rows, &row)
			default:
				var raw xmlutil.RawXML
				if err := raw.UnmarshalXML(d, t); err != nil {
					return err
				}
				tbl.Raw = append(tbl.Raw, raw)
			}
		case xml.EndElement:
			return nil
		}
	}
}

// MarshalXML implements xml.Marshaler for CT_Tbl.
func (tbl *CT_Tbl) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name = tbl.XMLName
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	if tbl.TblPr != nil {
		s := xml.StartElement{Name: xml.Name{Space: Ns, Local: "tblPr"}}
		if err := e.EncodeElement(tbl.TblPr, s); err != nil {
			return err
		}
	}
	if tbl.TblGrid != nil {
		s := xml.StartElement{Name: xml.Name{Space: Ns, Local: "tblGrid"}}
		if err := e.EncodeElement(tbl.TblGrid, s); err != nil {
			return err
		}
	}
	for _, row := range tbl.Rows {
		rowStart := xml.StartElement{Name: row.XMLName}
		if err := e.EncodeElement(row, rowStart); err != nil {
			return err
		}
	}
	for i := range tbl.Raw {
		if err := tbl.Raw[i].MarshalXML(e, xml.StartElement{}); err != nil {
			return err
		}
	}
	return e.EncodeToken(start.End())
}

// --- CT_TblPr ---

// UnmarshalXML implements xml.Unmarshaler for CT_TblPr.
func (pr *CT_TblPr) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "tblStyle":
				v := getAttrVal(t.Attr)
				if v != "" {
					pr.Style = &v
				}
				if err := d.Skip(); err != nil {
					return err
				}
			case "tblW":
				pr.Width = &CT_TblWidth{}
				for _, a := range t.Attr {
					switch a.Name.Local {
					case "w":
						pr.Width.W = a.Value
					case "type":
						pr.Width.Type = a.Value
					}
				}
				if err := d.Skip(); err != nil {
					return err
				}
			case "jc":
				v := getAttrVal(t.Attr)
				if v != "" {
					pr.Jc = &v
				}
				if err := d.Skip(); err != nil {
					return err
				}
			default:
				var raw xmlutil.RawXML
				if err := raw.UnmarshalXML(d, t); err != nil {
					return err
				}
				pr.Extra = append(pr.Extra, raw)
			}
		case xml.EndElement:
			return nil
		}
	}
}

// MarshalXML implements xml.Marshaler for CT_TblPr.
func (pr *CT_TblPr) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	if pr.Style != nil {
		s := xml.StartElement{
			Name: xml.Name{Space: Ns, Local: "tblStyle"},
			Attr: []xml.Attr{{Name: xml.Name{Space: Ns, Local: "val"}, Value: *pr.Style}},
		}
		if err := e.EncodeToken(s); err != nil {
			return err
		}
		if err := e.EncodeToken(s.End()); err != nil {
			return err
		}
	}
	if pr.Width != nil {
		s := xml.StartElement{
			Name: xml.Name{Space: Ns, Local: "tblW"},
			Attr: []xml.Attr{
				{Name: xml.Name{Space: Ns, Local: "w"}, Value: pr.Width.W},
				{Name: xml.Name{Space: Ns, Local: "type"}, Value: pr.Width.Type},
			},
		}
		if err := e.EncodeToken(s); err != nil {
			return err
		}
		if err := e.EncodeToken(s.End()); err != nil {
			return err
		}
	}
	if err := marshalValAttr(e, "jc", pr.Jc); err != nil {
		return err
	}
	for i := range pr.Extra {
		if err := pr.Extra[i].MarshalXML(e, xml.StartElement{}); err != nil {
			return err
		}
	}
	return e.EncodeToken(start.End())
}

// --- CT_TblGrid ---

// UnmarshalXML implements xml.Unmarshaler for CT_TblGrid.
func (g *CT_TblGrid) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if t.Name.Local == "gridCol" {
				col := &CT_TblGridCol{}
				for _, a := range t.Attr {
					if a.Name.Local == "w" {
						col.W = a.Value
					}
				}
				if err := d.Skip(); err != nil {
					return err
				}
				g.Cols = append(g.Cols, col)
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

// MarshalXML implements xml.Marshaler for CT_TblGrid.
func (g *CT_TblGrid) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	for _, col := range g.Cols {
		s := xml.StartElement{
			Name: xml.Name{Space: Ns, Local: "gridCol"},
			Attr: []xml.Attr{{Name: xml.Name{Space: Ns, Local: "w"}, Value: col.W}},
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

// --- CT_Row ---

// UnmarshalXML implements xml.Unmarshaler for CT_Row.
func (row *CT_Row) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	row.XMLName = start.Name
	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "trPr":
				var raw xmlutil.RawXML
				if err := raw.UnmarshalXML(d, t); err != nil {
					return err
				}
				row.TrPr = &raw
			case "tc":
				var tc CT_Tc
				if err := d.DecodeElement(&tc, &t); err != nil {
					return err
				}
				row.Cells = append(row.Cells, &tc)
			default:
				var raw xmlutil.RawXML
				if err := raw.UnmarshalXML(d, t); err != nil {
					return err
				}
				row.Raw = append(row.Raw, raw)
			}
		case xml.EndElement:
			return nil
		}
	}
}

// MarshalXML implements xml.Marshaler for CT_Row.
func (row *CT_Row) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name = row.XMLName
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	if row.TrPr != nil {
		if err := row.TrPr.MarshalXML(e, xml.StartElement{}); err != nil {
			return err
		}
	}
	for _, tc := range row.Cells {
		tcStart := xml.StartElement{Name: tc.XMLName}
		if err := e.EncodeElement(tc, tcStart); err != nil {
			return err
		}
	}
	for i := range row.Raw {
		if err := row.Raw[i].MarshalXML(e, xml.StartElement{}); err != nil {
			return err
		}
	}
	return e.EncodeToken(start.End())
}

// --- CT_Tc ---

// UnmarshalXML implements xml.Unmarshaler for CT_Tc.
func (tc *CT_Tc) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	tc.XMLName = start.Name
	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "tcPr":
				tc.TcPr = &CT_TcPr{}
				if err := d.DecodeElement(tc.TcPr, &t); err != nil {
					return err
				}
			case "p":
				var p CT_P
				if err := d.DecodeElement(&p, &t); err != nil {
					return err
				}
				tc.Content = append(tc.Content, BlockLevelContent{Paragraph: &p})
			case "tbl":
				var tbl CT_Tbl
				if err := d.DecodeElement(&tbl, &t); err != nil {
					return err
				}
				tc.Content = append(tc.Content, BlockLevelContent{Table: &tbl})
			default:
				var raw xmlutil.RawXML
				if err := raw.UnmarshalXML(d, t); err != nil {
					return err
				}
				tc.Content = append(tc.Content, BlockLevelContent{Raw: &raw})
			}
		case xml.EndElement:
			return nil
		}
	}
}

// MarshalXML implements xml.Marshaler for CT_Tc.
func (tc *CT_Tc) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name = tc.XMLName
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	if tc.TcPr != nil {
		s := xml.StartElement{Name: xml.Name{Space: Ns, Local: "tcPr"}}
		if err := e.EncodeElement(tc.TcPr, s); err != nil {
			return err
		}
	}
	for _, c := range tc.Content {
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
	return e.EncodeToken(start.End())
}

// --- CT_TcPr ---

// UnmarshalXML implements xml.Unmarshaler for CT_TcPr.
func (pr *CT_TcPr) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "tcW":
				pr.Width = &CT_TblWidth{}
				for _, a := range t.Attr {
					switch a.Name.Local {
					case "w":
						pr.Width.W = a.Value
					case "type":
						pr.Width.Type = a.Value
					}
				}
				if err := d.Skip(); err != nil {
					return err
				}
			case "vMerge":
				pr.VMerge = &CT_VMerge{}
				for _, a := range t.Attr {
					if a.Name.Local == "val" {
						pr.VMerge.Val = a.Value
					}
				}
				if err := d.Skip(); err != nil {
					return err
				}
			case "gridSpan":
				v := getAttrVal(t.Attr)
				if v != "" {
					n, _ := strconv.Atoi(v)
					pr.GridSpan = &n
				}
				if err := d.Skip(); err != nil {
					return err
				}
			case "shd":
				pr.Shading = &CT_Shd{}
				for _, a := range t.Attr {
					switch a.Name.Local {
					case "val":
						pr.Shading.Val = a.Value
					case "color":
						pr.Shading.Color = a.Value
					case "fill":
						pr.Shading.Fill = a.Value
					}
				}
				if err := d.Skip(); err != nil {
					return err
				}
			default:
				var raw xmlutil.RawXML
				if err := raw.UnmarshalXML(d, t); err != nil {
					return err
				}
				pr.Extra = append(pr.Extra, raw)
			}
		case xml.EndElement:
			return nil
		}
	}
}

// MarshalXML implements xml.Marshaler for CT_TcPr.
func (pr *CT_TcPr) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	if pr.Width != nil {
		s := xml.StartElement{
			Name: xml.Name{Space: Ns, Local: "tcW"},
			Attr: []xml.Attr{
				{Name: xml.Name{Space: Ns, Local: "w"}, Value: pr.Width.W},
				{Name: xml.Name{Space: Ns, Local: "type"}, Value: pr.Width.Type},
			},
		}
		if err := e.EncodeToken(s); err != nil {
			return err
		}
		if err := e.EncodeToken(s.End()); err != nil {
			return err
		}
	}
	if pr.GridSpan != nil {
		s := xml.StartElement{
			Name: xml.Name{Space: Ns, Local: "gridSpan"},
			Attr: []xml.Attr{{Name: xml.Name{Space: Ns, Local: "val"}, Value: strconv.Itoa(*pr.GridSpan)}},
		}
		if err := e.EncodeToken(s); err != nil {
			return err
		}
		if err := e.EncodeToken(s.End()); err != nil {
			return err
		}
	}
	if pr.VMerge != nil {
		s := xml.StartElement{Name: xml.Name{Space: Ns, Local: "vMerge"}}
		if pr.VMerge.Val != "" {
			s.Attr = []xml.Attr{{Name: xml.Name{Space: Ns, Local: "val"}, Value: pr.VMerge.Val}}
		}
		if err := e.EncodeToken(s); err != nil {
			return err
		}
		if err := e.EncodeToken(s.End()); err != nil {
			return err
		}
	}
	if pr.Shading != nil {
		s := xml.StartElement{Name: xml.Name{Space: Ns, Local: "shd"}}
		if pr.Shading.Val != "" {
			s.Attr = append(s.Attr, xml.Attr{Name: xml.Name{Space: Ns, Local: "val"}, Value: pr.Shading.Val})
		}
		if pr.Shading.Color != "" {
			s.Attr = append(s.Attr, xml.Attr{Name: xml.Name{Space: Ns, Local: "color"}, Value: pr.Shading.Color})
		}
		if pr.Shading.Fill != "" {
			s.Attr = append(s.Attr, xml.Attr{Name: xml.Name{Space: Ns, Local: "fill"}, Value: pr.Shading.Fill})
		}
		if err := e.EncodeToken(s); err != nil {
			return err
		}
		if err := e.EncodeToken(s.End()); err != nil {
			return err
		}
	}
	for i := range pr.Extra {
		if err := pr.Extra[i].MarshalXML(e, xml.StartElement{}); err != nil {
			return err
		}
	}
	return e.EncodeToken(start.End())
}
