package wml

import (
	"encoding/xml"
	"strconv"

	"github.com/ieshan/go-ooxml/xmlutil"
)

// CT_Comments represents <w:comments> — root of comments.xml
type CT_Comments struct {
	XMLName  xml.Name
	Comments []*CT_Comment
}

// CT_Comment represents <w:comment>
type CT_Comment struct {
	XMLName  xml.Name
	ID       int
	Author   string
	Date     string
	Initials string
	Content  []BlockLevelContent // Comment body (paragraphs)
}

// CT_CommentReference represents <w:commentReference> inside a run
type CT_CommentReference struct {
	ID int `xml:"id,attr"`
}

// UnmarshalXML implements xml.Unmarshaler for CT_Comments.
func (c *CT_Comments) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	c.XMLName = start.Name
	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if t.Name.Local == "comment" {
				var comment CT_Comment
				if err := d.DecodeElement(&comment, &t); err != nil {
					return err
				}
				c.Comments = append(c.Comments, &comment)
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

// MarshalXML implements xml.Marshaler for CT_Comments.
func (c *CT_Comments) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name = c.XMLName
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	for _, comment := range c.Comments {
		commentStart := xml.StartElement{Name: comment.XMLName}
		if err := e.EncodeElement(comment, commentStart); err != nil {
			return err
		}
	}
	return e.EncodeToken(start.End())
}

// UnmarshalXML implements xml.Unmarshaler for CT_Comment.
func (c *CT_Comment) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	c.XMLName = start.Name
	for _, a := range start.Attr {
		switch a.Name.Local {
		case "id":
			c.ID, _ = strconv.Atoi(a.Value)
		case "author":
			c.Author = a.Value
		case "date":
			c.Date = a.Value
		case "initials":
			c.Initials = a.Value
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
			case "p":
				var p CT_P
				if err := d.DecodeElement(&p, &t); err != nil {
					return err
				}
				c.Content = append(c.Content, BlockLevelContent{Paragraph: &p})
			case "tbl":
				var tbl CT_Tbl
				if err := d.DecodeElement(&tbl, &t); err != nil {
					return err
				}
				c.Content = append(c.Content, BlockLevelContent{Table: &tbl})
			default:
				var raw xmlutil.RawXML
				if err := raw.UnmarshalXML(d, t); err != nil {
					return err
				}
				c.Content = append(c.Content, BlockLevelContent{Raw: &raw})
			}
		case xml.EndElement:
			return nil
		}
	}
}

// MarshalXML implements xml.Marshaler for CT_Comment.
func (c *CT_Comment) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name = c.XMLName
	start.Attr = append(start.Attr,
		xml.Attr{Name: xml.Name{Space: Ns, Local: "id"}, Value: strconv.Itoa(c.ID)},
		xml.Attr{Name: xml.Name{Space: Ns, Local: "author"}, Value: c.Author},
	)
	if c.Date != "" {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Space: Ns, Local: "date"}, Value: c.Date})
	}
	if c.Initials != "" {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Space: Ns, Local: "initials"}, Value: c.Initials})
	}
	if err := e.EncodeToken(start); err != nil {
		return err
	}

	for _, bc := range c.Content {
		switch {
		case bc.Paragraph != nil:
			pStart := xml.StartElement{Name: bc.Paragraph.XMLName}
			if err := e.EncodeElement(bc.Paragraph, pStart); err != nil {
				return err
			}
		case bc.Table != nil:
			tblStart := xml.StartElement{Name: bc.Table.XMLName}
			if err := e.EncodeElement(bc.Table, tblStart); err != nil {
				return err
			}
		case bc.Raw != nil:
			if err := bc.Raw.MarshalXML(e, xml.StartElement{}); err != nil {
				return err
			}
		}
	}

	return e.EncodeToken(start.End())
}
