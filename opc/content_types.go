package opc

import (
	"encoding/xml"
	"path"
	"slices"
	"strings"

	"github.com/ieshan/go-ooxml/xmlutil"
)

const contentTypesNS = "http://schemas.openxmlformats.org/package/2006/content-types"

// ContentTypes holds the parsed content of a [Content_Types].xml file.
// Defaults maps file extensions to content types.
// Overrides maps specific part paths to content types.
type ContentTypes struct {
	Defaults  map[string]string // extension → content type
	Overrides map[string]string // part path → content type
}

// xmlContentTypes is the internal XML representation of [Content_Types].xml.
type xmlContentTypes struct {
	XMLName   xml.Name      `xml:"Types"`
	Xmlns     string        `xml:"xmlns,attr"`
	Defaults  []xmlDefault  `xml:"Default"`
	Overrides []xmlOverride `xml:"Override"`
}

type xmlDefault struct {
	Extension   string `xml:"Extension,attr"`
	ContentType string `xml:"ContentType,attr"`
}

type xmlOverride struct {
	PartName    string `xml:"PartName,attr"`
	ContentType string `xml:"ContentType,attr"`
}

// ParseContentTypes parses [Content_Types].xml data.
func ParseContentTypes(data []byte) (*ContentTypes, error) {
	var raw xmlContentTypes
	if err := xmlutil.Unmarshal(data, &raw); err != nil {
		return nil, err
	}

	ct := &ContentTypes{
		Defaults:  make(map[string]string, len(raw.Defaults)),
		Overrides: make(map[string]string, len(raw.Overrides)),
	}
	for _, d := range raw.Defaults {
		ct.Defaults[d.Extension] = d.ContentType
	}
	for _, o := range raw.Overrides {
		ct.Overrides[o.PartName] = o.ContentType
	}
	return ct, nil
}

// ContentTypeFor returns the content type for a given part path.
// An override for the exact path takes precedence over the default for its extension.
func (ct *ContentTypes) ContentTypeFor(partPath string) string {
	if ct == nil {
		return ""
	}
	if mime, ok := ct.Overrides[partPath]; ok {
		return mime
	}
	ext := strings.TrimPrefix(path.Ext(partPath), ".")
	if ext == "" {
		return ""
	}
	return ct.Defaults[ext]
}

// Marshal serializes ContentTypes back to XML bytes. Keys are sorted for
// deterministic output.
func (ct *ContentTypes) Marshal() ([]byte, error) {
	raw := xmlContentTypes{
		Xmlns:     contentTypesNS,
		Defaults:  make([]xmlDefault, 0, len(ct.Defaults)),
		Overrides: make([]xmlOverride, 0, len(ct.Overrides)),
	}

	// Sort keys for deterministic output.
	exts := make([]string, 0, len(ct.Defaults))
	for ext := range ct.Defaults {
		exts = append(exts, ext)
	}
	slices.Sort(exts)
	for _, ext := range exts {
		raw.Defaults = append(raw.Defaults, xmlDefault{Extension: ext, ContentType: ct.Defaults[ext]})
	}

	paths := make([]string, 0, len(ct.Overrides))
	for p := range ct.Overrides {
		paths = append(paths, p)
	}
	slices.Sort(paths)
	for _, p := range paths {
		raw.Overrides = append(raw.Overrides, xmlOverride{PartName: p, ContentType: ct.Overrides[p]})
	}

	out, err := xml.Marshal(&raw)
	if err != nil {
		return nil, err
	}
	return xmlutil.AddXMLHeader(out), nil
}
