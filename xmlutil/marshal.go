package xmlutil

import (
	"bytes"
	"encoding/xml"
	"errors"
	"io"
	"strings"
)

// Sentinel errors for security checks.
var (
	ErrXXE          = errors.New("ooxml: XML contains external entity reference")
	ErrNestingDepth = errors.New("ooxml: XML nesting depth exceeded")
	ErrTokenLimit   = errors.New("ooxml: XML token count exceeded")
)

const (
	defaultMaxNestingDepth = 256
	defaultMaxTokenCount   = int64(10_000_000)
)

// unmarshalConfig holds options for Unmarshal.
type unmarshalConfig struct {
	maxNestingDepth int
	maxTokenCount   int64
}

// UnmarshalOption configures the behavior of Unmarshal.
type UnmarshalOption func(*unmarshalConfig)

// WithMaxNestingDepth sets the maximum allowed XML nesting depth.
func WithMaxNestingDepth(n int) UnmarshalOption {
	return func(c *unmarshalConfig) {
		c.maxNestingDepth = n
	}
}

// WithMaxTokenCount sets the maximum allowed XML token count.
func WithMaxTokenCount(n int64) UnmarshalOption {
	return func(c *unmarshalConfig) {
		c.maxTokenCount = n
	}
}

// Unmarshal securely unmarshals XML data into v. It rejects XML containing
// DTD declarations (to prevent XXE attacks), enforces a nesting depth limit,
// and enforces a token count limit. After security checks pass, it uses
// standard xml.Unmarshal to decode.
func Unmarshal(data []byte, v any, opts ...UnmarshalOption) error {
	cfg := unmarshalConfig{
		maxNestingDepth: defaultMaxNestingDepth,
		maxTokenCount:   defaultMaxTokenCount,
	}
	for _, opt := range opts {
		opt(&cfg)
	}

	// DTD rejection: scan raw bytes for <!DOCTYPE.
	if bytes.Contains(data, []byte("<!DOCTYPE")) {
		return ErrXXE
	}

	// Token scan for nesting depth and token count.
	dec := xml.NewDecoder(bytes.NewReader(data))
	var depth int
	var tokenCount int64
	for {
		tok, err := dec.Token()
		if err != nil {
			if err == io.EOF {
				break
			}
			return err
		}
		tokenCount++
		if tokenCount > cfg.maxTokenCount {
			return ErrTokenLimit
		}
		switch tok.(type) {
		case xml.StartElement:
			depth++
			if depth > cfg.maxNestingDepth {
				return ErrNestingDepth
			}
		case xml.EndElement:
			depth--
		}
	}

	return xml.Unmarshal(data, v)
}

// Marshal marshals v into XML. If registry is non-nil, the output is
// post-processed to rewrite default namespace declarations into prefixed
// declarations using the registry mappings.
func Marshal(v any, registry *NamespaceRegistry) ([]byte, error) {
	data, err := xml.Marshal(v)
	if err != nil {
		return nil, err
	}
	if registry == nil {
		return data, nil
	}
	return rewriteNamespacePrefixes(data, registry)
}

// MarshalIndent is like Marshal but applies indentation to the output.
func MarshalIndent(v any, registry *NamespaceRegistry, indent string) ([]byte, error) {
	data, err := xml.MarshalIndent(v, "", indent)
	if err != nil {
		return nil, err
	}
	if registry == nil {
		return data, nil
	}
	return rewriteNamespacePrefixes(data, registry)
}

// xmlHeader is the standard XML declaration.
const xmlHeader = `<?xml version="1.0" encoding="UTF-8"?>`

// AddXMLHeader prepends the standard XML declaration to data if it is not
// already present.
func AddXMLHeader(data []byte) []byte {
	if bytes.HasPrefix(data, []byte("<?xml")) {
		return data
	}
	out := make([]byte, 0, len(xmlHeader)+len(data))
	out = append(out, xmlHeader...)
	out = append(out, data...)
	return out
}

// rewriteNamespacePrefixes rewrites XML output so that default namespace
// declarations (xmlns="uri") become prefixed declarations (xmlns:prefix="uri")
// and element/attribute names are rewritten to use the corresponding prefix.
func rewriteNamespacePrefixes(data []byte, registry *NamespaceRegistry) ([]byte, error) {
	dec := xml.NewDecoder(bytes.NewReader(data))
	var buf bytes.Buffer
	enc := xml.NewEncoder(&buf)

	// Track which namespace prefixes have been declared.
	declared := make(map[string]bool)

	for {
		tok, err := dec.Token()
		if err != nil {
			if err == io.EOF {
				break
			}
			return nil, err
		}

		switch t := tok.(type) {
		case xml.StartElement:
			t = rewriteStartElement(t, registry, declared)
			if err := enc.EncodeToken(t); err != nil {
				return nil, err
			}
		case xml.EndElement:
			t.Name = rewriteName(t.Name, registry)
			if err := enc.EncodeToken(t); err != nil {
				return nil, err
			}
		default:
			if err := enc.EncodeToken(xml.CopyToken(tok)); err != nil {
				return nil, err
			}
		}
	}

	if err := enc.Flush(); err != nil {
		return nil, err
	}
	return buf.Bytes(), nil
}

// rewriteName rewrites an xml.Name to use a prefix from the registry if the
// namespace URI is known. If the URI has no registered prefix, the name is
// returned unchanged.
func rewriteName(name xml.Name, registry *NamespaceRegistry) xml.Name {
	if name.Space == "" {
		return name
	}
	prefix, ok := registry.Prefix(name.Space)
	if !ok {
		return name
	}
	return xml.Name{
		Local: prefix + ":" + name.Local,
	}
}

// rewriteStartElement rewrites element and attribute names and manages
// namespace declarations.
func rewriteStartElement(el xml.StartElement, registry *NamespaceRegistry, declared map[string]bool) xml.StartElement {
	// Collect namespaces used in this element.
	usedNS := make(map[string]string) // uri -> prefix

	// Rewrite element name.
	if el.Name.Space != "" {
		if prefix, ok := registry.Prefix(el.Name.Space); ok {
			usedNS[el.Name.Space] = prefix
			el.Name = xml.Name{Local: prefix + ":" + el.Name.Local}
		}
	}

	// Filter and rewrite attributes.
	var attrs []xml.Attr
	for _, a := range el.Attr {
		// Skip default xmlns declarations for URIs we're rewriting to prefixed form.
		if a.Name.Local == "xmlns" && a.Name.Space == "" {
			if _, ok := registry.Prefix(a.Value); ok {
				// We'll add a prefixed declaration instead.
				continue
			}
		}
		// Skip synthetic _xmlns declarations generated by Go's xml.Encoder
		// when marshaling struct fields with namespace-qualified attribute tags.
		// These produce broken attributes like xmlns:_xmlns="xmlns" and
		// _xmlns:main="http://...".
		if strings.HasPrefix(a.Name.Local, "_xmlns") || strings.HasPrefix(a.Name.Space, "_xmlns") {
			continue
		}
		if a.Name.Space == "xmlns" {
			// This is a namespace declaration (xmlns:foo="...").
			// Skip it if the URI is known to the registry — we'll re-declare it properly.
			if _, ok := registry.Prefix(a.Value); ok {
				continue
			}
		}
		// Rewrite attribute names with namespaces.
		if a.Name.Space != "" && a.Name.Space != "xmlns" {
			if prefix, ok := registry.Prefix(a.Name.Space); ok {
				usedNS[a.Name.Space] = prefix
				a.Name = xml.Name{Local: prefix + ":" + a.Name.Local}
			}
		}
		attrs = append(attrs, a)
	}

	// Add xmlns:prefix declarations for namespaces used but not yet declared.
	for uri, prefix := range usedNS {
		if !declared[prefix] {
			declared[prefix] = true
			attrs = append(attrs, xml.Attr{
				Name:  xml.Name{Local: "xmlns:" + prefix},
				Value: uri,
			})
		}
	}

	el.Attr = attrs
	return el
}
