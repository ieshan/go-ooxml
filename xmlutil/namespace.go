// Package xmlutil provides XML utilities for OOXML processing, including
// namespace prefix preservation, raw XML node handling, and secure
// marshal/unmarshal with XXE prevention.
package xmlutil

import "iter"

// Standard OOXML namespace URIs.
const (
	NsWordprocessingML = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
	NsRelationships    = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
	NsDrawingWordML    = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
	NsDrawingML        = "http://schemas.openxmlformats.org/drawingml/2006/main"
	NsPicture          = "http://schemas.openxmlformats.org/drawingml/2006/picture"
	NsMarkupCompat     = "http://schemas.openxmlformats.org/markup-compatibility/2006"
	NsPackageRels      = "http://schemas.openxmlformats.org/package/2006/relationships"
	NsContentTypes     = "http://schemas.openxmlformats.org/package/2006/content-types"
	NsCoreProperties   = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
	NsDCElements       = "http://purl.org/dc/elements/1.1/"
	NsDCTerms          = "http://purl.org/dc/terms/"
	NsXSI              = "http://www.w3.org/2001/XMLSchema-instance"
	NsMSWordML14       = "http://schemas.microsoft.com/office/word/2010/wordml"
)

// NsEntry holds a namespace prefix and its corresponding URI.
type NsEntry struct {
	Prefix string
	URI    string
}

// Ns constructs an NsEntry from a prefix and URI string pair.
func Ns(prefix, uri string) NsEntry {
	return NsEntry{Prefix: prefix, URI: uri}
}

// NamespaceRegistry maps namespace prefixes to URIs and vice versa.
// It is immutable after construction and safe for concurrent use.
type NamespaceRegistry struct {
	prefixToURI map[string]string
	uriToPrefix map[string]string
}

// NewRegistry creates a NamespaceRegistry from the provided NsEntry values.
func NewRegistry(entries ...NsEntry) *NamespaceRegistry {
	r := &NamespaceRegistry{
		prefixToURI: make(map[string]string, len(entries)),
		uriToPrefix: make(map[string]string, len(entries)),
	}
	for _, e := range entries {
		r.prefixToURI[e.Prefix] = e.URI
		r.uriToPrefix[e.URI] = e.Prefix
	}
	return r
}

// URI returns the namespace URI for the given prefix, and whether it was found.
func (r *NamespaceRegistry) URI(prefix string) (string, bool) {
	uri, ok := r.prefixToURI[prefix]
	return uri, ok
}

// Prefix returns the namespace prefix for the given URI, and whether it was found.
func (r *NamespaceRegistry) Prefix(uri string) (string, bool) {
	prefix, ok := r.uriToPrefix[uri]
	return prefix, ok
}

// Prefixes returns an iterator over all (prefix, URI) pairs in the registry.
func (r *NamespaceRegistry) Prefixes() iter.Seq2[string, string] {
	return func(yield func(string, string) bool) {
		for prefix, uri := range r.prefixToURI {
			if !yield(prefix, uri) {
				return
			}
		}
	}
}

// OOXML is the standard registry of well-known OOXML namespace prefixes and URIs.
var OOXML = NewRegistry(
	Ns("w", NsWordprocessingML),
	Ns("r", NsRelationships),
	Ns("wp", NsDrawingWordML),
	Ns("a", NsDrawingML),
	Ns("pic", NsPicture),
	Ns("mc", NsMarkupCompat),
	Ns("pr", NsPackageRels),
	Ns("cp", NsCoreProperties),
	Ns("dc", NsDCElements),
	Ns("dcterms", NsDCTerms),
	Ns("xsi", NsXSI),
	Ns("w14", NsMSWordML14),
)
