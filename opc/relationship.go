package opc

import (
	"encoding/xml"
	"fmt"
	"strconv"
	"strings"

	"github.com/ieshan/go-ooxml/xmlutil"
)

// Standard OPC relationship type URIs.
const (
	RelOfficeDocument = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
	RelStyles         = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
	RelComments       = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
	RelHyperlink      = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
	RelHeader         = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
	RelFooter         = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"
	RelCoreProperties = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
)

// Relationship represents a single OPC relationship entry.
type Relationship struct {
	ID       string
	Type     string
	Target   string
	External bool
}

// xmlRelationship is the internal XML representation of a <Relationship> element.
type xmlRelationship struct {
	ID         string `xml:"Id,attr"`
	Type       string `xml:"Type,attr"`
	Target     string `xml:"Target,attr"`
	TargetMode string `xml:"TargetMode,attr,omitempty"`
}

// xmlRelationships is the internal XML representation of a <Relationships> document.
type xmlRelationships struct {
	XMLName xml.Name          `xml:"http://schemas.openxmlformats.org/package/2006/relationships Relationships"`
	Rels    []xmlRelationship `xml:"Relationship"`
}

// ParseRelationships parses the XML bytes of a .rels part and returns the
// slice of Relationship values it contains.
func ParseRelationships(data []byte) ([]Relationship, error) {
	var doc xmlRelationships
	if err := xmlutil.Unmarshal(data, &doc); err != nil {
		return nil, err
	}
	rels := make([]Relationship, len(doc.Rels))
	for i, r := range doc.Rels {
		rels[i] = Relationship{
			ID:       r.ID,
			Type:     r.Type,
			Target:   r.Target,
			External: r.TargetMode == "External",
		}
	}
	return rels, nil
}

// MarshalRelationships encodes a slice of Relationship values as XML bytes
// representing a .rels part, including the XML declaration header.
func MarshalRelationships(rels []Relationship) ([]byte, error) {
	doc := xmlRelationships{
		Rels: make([]xmlRelationship, len(rels)),
	}
	for i, r := range rels {
		xr := xmlRelationship{
			ID:     r.ID,
			Type:   r.Type,
			Target: r.Target,
		}
		if r.External {
			xr.TargetMode = "External"
		}
		doc.Rels[i] = xr
	}
	data, err := xml.Marshal(&doc)
	if err != nil {
		return nil, err
	}
	return xmlutil.AddXMLHeader(data), nil
}

// NextRelID scans existing relationship IDs for the maximum numeric suffix of
// the form "rId<N>" and returns the next ID (e.g. "rId4" when the max is 3).
// If no such IDs exist, it returns "rId1".
func NextRelID(rels []Relationship) string {
	maxVal := 0
	for _, r := range rels {
		if strings.HasPrefix(r.ID, "rId") {
			n, err := strconv.Atoi(r.ID[3:])
			if err == nil && n > maxVal {
				maxVal = n
			}
		}
	}
	return fmt.Sprintf("rId%d", maxVal+1)
}

// FindRelByType returns the first Relationship with the given type URI, or nil
// if none is found.
func FindRelByType(rels []Relationship, relType string) *Relationship {
	for i := range rels {
		if rels[i].Type == relType {
			return &rels[i]
		}
	}
	return nil
}
