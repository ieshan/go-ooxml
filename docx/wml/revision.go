package wml

import (
	"encoding/xml"
	"strconv"
)

// CT_RPrChange represents <w:rPrChange> — tracked run formatting change.
type CT_RPrChange struct {
	XMLName xml.Name
	ID      int
	Author  string
	Date    string
	RPr     *CT_RPr // Previous formatting
}

// CT_PPrChange represents <w:pPrChange> — tracked paragraph property change.
type CT_PPrChange struct {
	XMLName xml.Name
	ID      int
	Author  string
	Date    string
	PPr     *CT_PPr // Previous paragraph properties
}

// UnmarshalXML implements xml.Unmarshaler for CT_RPrChange.
func (r *CT_RPrChange) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	r.XMLName = start.Name
	for _, a := range start.Attr {
		switch a.Name.Local {
		case "id":
			r.ID, _ = strconv.Atoi(a.Value)
		case "author":
			r.Author = a.Value
		case "date":
			r.Date = a.Value
		}
	}
	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if t.Name.Local == "rPr" {
				r.RPr = &CT_RPr{}
				if err := d.DecodeElement(r.RPr, &t); err != nil {
					return err
				}
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

// MarshalXML implements xml.Marshaler for CT_RPrChange.
func (r *CT_RPrChange) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name = r.XMLName
	start.Attr = append(start.Attr,
		xml.Attr{Name: xml.Name{Space: Ns, Local: "id"}, Value: strconv.Itoa(r.ID)},
		xml.Attr{Name: xml.Name{Space: Ns, Local: "author"}, Value: r.Author},
	)
	if r.Date != "" {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Space: Ns, Local: "date"}, Value: r.Date})
	}
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	if r.RPr != nil {
		rprStart := xml.StartElement{Name: xml.Name{Space: Ns, Local: "rPr"}}
		if err := e.EncodeElement(r.RPr, rprStart); err != nil {
			return err
		}
	}
	return e.EncodeToken(start.End())
}

// UnmarshalXML implements xml.Unmarshaler for CT_PPrChange.
func (p *CT_PPrChange) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	p.XMLName = start.Name
	for _, a := range start.Attr {
		switch a.Name.Local {
		case "id":
			p.ID, _ = strconv.Atoi(a.Value)
		case "author":
			p.Author = a.Value
		case "date":
			p.Date = a.Value
		}
	}
	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if t.Name.Local == "pPr" {
				p.PPr = &CT_PPr{}
				if err := d.DecodeElement(p.PPr, &t); err != nil {
					return err
				}
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

// MarshalXML implements xml.Marshaler for CT_PPrChange.
func (p *CT_PPrChange) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name = p.XMLName
	start.Attr = append(start.Attr,
		xml.Attr{Name: xml.Name{Space: Ns, Local: "id"}, Value: strconv.Itoa(p.ID)},
		xml.Attr{Name: xml.Name{Space: Ns, Local: "author"}, Value: p.Author},
	)
	if p.Date != "" {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Space: Ns, Local: "date"}, Value: p.Date})
	}
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	if p.PPr != nil {
		pprStart := xml.StartElement{Name: xml.Name{Space: Ns, Local: "pPr"}}
		if err := e.EncodeElement(p.PPr, pprStart); err != nil {
			return err
		}
	}
	return e.EncodeToken(start.End())
}
