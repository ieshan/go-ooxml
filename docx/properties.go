package docx

import (
	"encoding/xml"
	"fmt"
	"time"

	"github.com/ieshan/go-ooxml/opc"
	"github.com/ieshan/go-ooxml/xmlutil"
)

// Core properties content type and part path.
const (
	ctCoreProps   = "application/vnd.openxmlformats-package.core-properties+xml"
	partCoreProps = "docProps/core.xml"
)

// corePropertiesXML is the internal XML representation of docProps/core.xml.
type corePropertiesXML struct {
	XMLName        xml.Name `xml:"http://schemas.openxmlformats.org/package/2006/metadata/core-properties coreProperties"`
	Title          string   `xml:"http://purl.org/dc/elements/1.1/ title,omitempty"`
	Creator        string   `xml:"http://purl.org/dc/elements/1.1/ creator,omitempty"`
	Description    string   `xml:"http://purl.org/dc/elements/1.1/ description,omitempty"`
	Subject        string   `xml:"http://purl.org/dc/elements/1.1/ subject,omitempty"`
	Created        string   `xml:"http://purl.org/dc/terms/ created,omitempty"`
	Modified       string   `xml:"http://purl.org/dc/terms/ modified,omitempty"`
	LastModifiedBy string   `xml:"lastModifiedBy,omitempty"`
	Keywords       string   `xml:"keywords,omitempty"`
	Category       string   `xml:"category,omitempty"`
}

// Properties provides access to the document's core (Dublin Core) metadata.
type Properties struct {
	doc *Document
	cp  corePropertiesXML
}

// Properties returns the document's core properties, loading them from
// docProps/core.xml if present in the package.
func (d *Document) Properties() *Properties {
	d.mu.Lock()
	defer d.mu.Unlock()

	if d.props != nil {
		return d.props
	}

	p := &Properties{doc: d}

	// Try to load existing core properties.
	if part, ok := d.pkg.Parts[partCoreProps]; ok {
		data, err := part.Data()
		if err == nil {
			_ = xmlutil.Unmarshal(data, &p.cp)
		}
	}

	d.props = p
	return p
}

// Title returns the document title.
func (p *Properties) Title() string {
	p.doc.mu.RLock()
	defer p.doc.mu.RUnlock()
	return p.cp.Title
}

// SetTitle sets the document title.
func (p *Properties) SetTitle(v string) {
	p.doc.mu.Lock()
	defer p.doc.mu.Unlock()
	p.cp.Title = v
}

// Author returns the document creator/author.
func (p *Properties) Author() string {
	p.doc.mu.RLock()
	defer p.doc.mu.RUnlock()
	return p.cp.Creator
}

// SetAuthor sets the document creator/author.
func (p *Properties) SetAuthor(v string) {
	p.doc.mu.Lock()
	defer p.doc.mu.Unlock()
	p.cp.Creator = v
}

// Description returns the document description.
func (p *Properties) Description() string {
	p.doc.mu.RLock()
	defer p.doc.mu.RUnlock()
	return p.cp.Description
}

// SetDescription sets the document description.
func (p *Properties) SetDescription(v string) {
	p.doc.mu.Lock()
	defer p.doc.mu.Unlock()
	p.cp.Description = v
}

// Created returns the document creation time, or zero if not set.
func (p *Properties) Created() time.Time {
	p.doc.mu.RLock()
	defer p.doc.mu.RUnlock()
	return parseDCTime(p.cp.Created)
}

// Modified returns the document last-modified time, or zero if not set.
func (p *Properties) Modified() time.Time {
	p.doc.mu.RLock()
	defer p.doc.mu.RUnlock()
	return parseDCTime(p.cp.Modified)
}

// parseDCTime parses an ISO 8601 / RFC 3339 timestamp string used in Dublin
// Core metadata. Returns zero time on parse failure or empty input.
func parseDCTime(s string) time.Time {
	if s == "" {
		return time.Time{}
	}
	t, err := time.Parse(time.RFC3339, s)
	if err != nil {
		return time.Time{}
	}
	return t
}

// syncCoreProperties marshals the in-memory Properties back to the OPC package.
// It also ensures the necessary relationship and content type override exist.
// Caller must hold the write lock on the document.
func (d *Document) syncCoreProperties() error {
	if d.props == nil {
		return nil
	}

	data, err := marshalCoreProperties(&d.props.cp)
	if err != nil {
		return fmt.Errorf("docx: marshal core properties: %w", err)
	}

	upsertPartData(d.pkg, partCoreProps, ctCoreProps, data)
	d.ensureCorePropertiesRelationship()
	return nil
}

// marshalCoreProperties marshals corePropertiesXML to XML bytes with a header.
func marshalCoreProperties(cp *corePropertiesXML) ([]byte, error) {
	data, err := xml.Marshal(cp)
	if err != nil {
		return nil, err
	}
	return xmlutil.AddXMLHeader(data), nil
}

// ensureCorePropertiesRelationship ensures the package-level relationship to
// docProps/core.xml exists and the content type override is registered.
// Caller must hold the write lock.
func (d *Document) ensureCorePropertiesRelationship() {
	if opc.FindRelByType(d.pkg.Rels, opc.RelCoreProperties) != nil {
		return
	}
	d.pkg.Rels = append(d.pkg.Rels, opc.Relationship{
		ID:     opc.NextRelID(d.pkg.Rels),
		Type:   opc.RelCoreProperties,
		Target: partCoreProps,
	})
	d.pkg.ContentTypes.Overrides["/"+partCoreProps] = ctCoreProps
}
