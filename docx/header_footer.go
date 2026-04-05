package docx

// HeaderFooterType identifies which variant of header or footer is referenced.
type HeaderFooterType int

const (
	// HdrFtrDefault is the default (odd-page) header or footer.
	HdrFtrDefault HeaderFooterType = iota
	// HdrFtrFirst is the first-page header or footer.
	HdrFtrFirst
	// HdrFtrEven is the even-page header or footer.
	HdrFtrEven
)

// HeaderFooter is a stub representing a header or footer part.
// Headers and footers are stored as separate OPC parts referenced from sectPr;
// full implementation will follow when OPC part references are wired up.
type HeaderFooter struct {
	doc *Document
	typ HeaderFooterType
	// Will hold reference to header/footer OPC part in future.
}

// Text returns an empty string (stub).
func (hf *HeaderFooter) Text() string { return "" }

// Markdown returns an empty string (stub).
func (hf *HeaderFooter) Markdown() string { return "" }
