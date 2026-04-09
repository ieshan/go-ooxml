package docx

import (
	"encoding/xml"
	"fmt"
	"io"
	"path"
	"strconv"
	"strings"
	"sync"
	"time"

	"github.com/ieshan/go-ooxml/docx/wml"
	"github.com/ieshan/go-ooxml/opc"
	"github.com/ieshan/go-ooxml/xmlutil"
)

// Content types for well-known DOCX parts.
const (
	ctDocument = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
	ctStyles   = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"
	ctComments = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
	ctHeader   = "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"
	ctFooter   = "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"
)

// Default part paths within a DOCX package.
const (
	partDocument = "word/document.xml"
	partStyles   = "word/styles.xml"
	partComments = "word/comments.xml"
)

// Config configures how a Document is opened or created.
type Config struct {
	Author   string    // Default author for comments/revisions
	Date     time.Time // Default date (zero = time.Now() per operation)
	Mode     OpenMode  // ModeInMemory (default) or ModeStreaming
	Security *opc.SecurityLimits
}

// OpenMode selects the memory strategy for the document.
type OpenMode int

const (
	// ModeInMemory loads the entire package into memory (default).
	ModeInMemory OpenMode = iota
	// ModeStreaming is reserved for future streaming support.
	ModeStreaming
)

// hdrFtrPart holds a parsed header or footer OPC part.
type hdrFtrPart struct {
	path  string           // OPC part path (e.g., "word/header1.xml")
	relID string           // relationship ID (e.g., "rId7")
	typ   HeaderFooterType // default, first, even
	isHdr bool             // true = header, false = footer
	body  *wml.CT_Body     // parsed content (same structure as document body)
}

// Document represents an Office Open XML word processing document (.docx).
// Thread-safe: concurrent reads are safe; write operations acquire an exclusive lock.
type Document struct {
	mu       sync.RWMutex
	pkg      *opc.Package
	doc      *wml.CT_Document // parsed document.xml
	styles   *wml.CT_Styles   // parsed styles.xml (may be nil)
	comments *wml.CT_Comments // parsed comments.xml (may be nil)
	props    *Properties      // lazily loaded core properties (may be nil)
	sectPr   *wml.CT_SectPr   // cached body section properties (may be nil)
	hdrFtrs  []hdrFtrPart     // parsed header/footer parts
	docPath  string           // path of main document part in the package
	cfg      Config
}

// New creates a blank document with a minimal valid structure.
func New(cfg *Config) (*Document, error) {
	c := resolveConfig(cfg)

	pkg := opc.NewPackage()

	// Set up content types.
	pkg.ContentTypes.Defaults["rels"] = "application/vnd.openxmlformats-package.relationships+xml"
	pkg.ContentTypes.Defaults["xml"] = "application/xml"
	pkg.ContentTypes.Overrides["/"+partDocument] = ctDocument

	// Add package-level relationship to the main document part.
	pkg.AddRelationship(opc.RelOfficeDocument, partDocument)

	// Create minimal CT_Document with an empty body.
	doc := &wml.CT_Document{
		XMLName: xml.Name{Space: wml.Ns, Local: "document"},
		Body: &wml.CT_Body{
			XMLName: xml.Name{Space: wml.Ns, Local: "body"},
		},
	}

	// Marshal and store the document part.
	docData, err := marshalDocument(doc)
	if err != nil {
		return nil, fmt.Errorf("docx: marshal new document: %w", err)
	}
	pkg.AddPart(partDocument, ctDocument, docData)

	d := &Document{
		pkg:     pkg,
		doc:     doc,
		docPath: partDocument,
		cfg:     c,
	}
	d.initDefaultStyles()
	return d, nil
}

// Open opens a DOCX file from the given path.
func Open(filePath string, cfg *Config) (*Document, error) {
	c := resolveConfig(cfg)

	var opts *opc.OpenOptions
	if c.Security != nil {
		opts = &opc.OpenOptions{Security: c.Security}
	}

	pkg, err := opc.OpenFile(filePath, opts)
	if err != nil {
		return nil, fmt.Errorf("docx: open %q: %w", filePath, err)
	}

	return loadDocument(pkg, c)
}

// OpenReader opens a DOCX from an io.ReaderAt.
func OpenReader(r io.ReaderAt, size int64, cfg *Config) (*Document, error) {
	c := resolveConfig(cfg)

	opcOpts := &opc.OpenOptions{}
	if c.Security != nil {
		opcOpts.Security = c.Security
	}

	var pkg *opc.Package
	var err error
	if c.Mode == ModeStreaming {
		pkg, err = opc.OpenStreaming(r, size, opcOpts)
	} else {
		pkg, err = opc.Open(r, size, opcOpts)
	}
	if err != nil {
		return nil, fmt.Errorf("docx: open reader: %w", err)
	}

	return loadDocument(pkg, c)
}

// WriteTo writes the document as a DOCX archive to w and returns the number
// of bytes written. It implements [io.WriterTo].
func (d *Document) WriteTo(w io.Writer) (int64, error) {
	d.mu.Lock()
	defer d.mu.Unlock()

	if err := d.syncParts(); err != nil {
		return 0, err
	}
	return d.pkg.WriteTo(w)
}

// Write writes the document as a DOCX archive to w.
func (d *Document) Write(w io.Writer) error {
	_, err := d.WriteTo(w)
	return err
}

// Close releases resources held by the document.
func (d *Document) Close() error {
	d.mu.Lock()
	defer d.mu.Unlock()

	d.pkg = nil
	d.doc = nil
	d.styles = nil
	d.comments = nil
	d.props = nil
	return nil
}

// Body returns a body proxy for manipulating document content.
func (d *Document) Body() *Body {
	return &Body{doc: d, el: d.doc.Body}
}

// loadDocument parses the main document, styles, and comments from an OPC package.
func loadDocument(pkg *opc.Package, cfg Config) (*Document, error) {
	// Find the main document part via package relationships.
	rel := opc.FindRelByType(pkg.Rels, opc.RelOfficeDocument)
	if rel == nil {
		return nil, fmt.Errorf("docx: %w: missing officeDocument relationship", ErrInvalidDoc)
	}

	docPath := rel.Target
	docPart, ok := pkg.Parts[docPath]
	if !ok {
		return nil, fmt.Errorf("docx: %w: part %q not found", ErrInvalidDoc, docPath)
	}

	// Parse document.xml.
	docData, err := docPart.Data()
	if err != nil {
		return nil, fmt.Errorf("docx: read document part: %w", err)
	}

	var doc wml.CT_Document
	if err := xmlutil.Unmarshal(docData, &doc); err != nil {
		return nil, fmt.Errorf("docx: parse document.xml: %w", err)
	}

	d := &Document{
		pkg:     pkg,
		doc:     &doc,
		docPath: docPath,
		cfg:     cfg,
	}

	// Optionally parse styles.xml if referenced by the document part rels.
	stylesPath := findPartRelTarget(docPart, opc.RelStyles, docPath)
	if stylesPath != "" {
		if stylesPart, ok := pkg.Parts[stylesPath]; ok {
			sData, err := stylesPart.Data()
			if err == nil {
				var styles wml.CT_Styles
				if err := xmlutil.Unmarshal(sData, &styles); err == nil {
					d.styles = &styles
				}
			}
		}
	}

	// Optionally parse comments.xml if referenced.
	commentsPath := findPartRelTarget(docPart, opc.RelComments, docPath)
	if commentsPath != "" {
		if commentsPart, ok := pkg.Parts[commentsPath]; ok {
			cData, err := commentsPart.Data()
			if err == nil {
				var comments wml.CT_Comments
				if err := xmlutil.Unmarshal(cData, &comments); err == nil {
					d.comments = &comments
				}
			}
		}
	}

	// Optionally parse header/footer parts referenced by the document part.
	for _, rel := range docPart.Rels {
		isHdr := rel.Type == opc.RelHeader
		isFtr := rel.Type == opc.RelFooter
		if !isHdr && !isFtr {
			continue
		}
		partPath := resolveRelPath(docPath, rel.Target)
		part, ok := pkg.Parts[partPath]
		if !ok {
			continue
		}
		data, err := part.Data()
		if err != nil {
			continue
		}
		// Headers/footers have a body-like structure (hdr or ftr root with paragraphs/tables).
		// We parse them as CT_Body since the content model is identical.
		var body wml.CT_Body
		if xmlutil.Unmarshal(data, &body) != nil {
			continue
		}
		// Determine the type (default/first/even) from the sectPr references.
		hfType := hdrFtrTypeFromRelID(d, rel.ID, isHdr)
		d.hdrFtrs = append(d.hdrFtrs, hdrFtrPart{
			path:  partPath,
			relID: rel.ID,
			typ:   hfType,
			isHdr: isHdr,
			body:  &body,
		})
	}

	// Optionally load core properties from docProps/core.xml.
	// The relationship is package-level (not document-part-level).
	coreRel := opc.FindRelByType(pkg.Rels, opc.RelCoreProperties)
	if coreRel != nil {
		corePath := coreRel.Target
		if corePart, ok := pkg.Parts[corePath]; ok {
			cpData, err := corePart.Data()
			if err == nil {
				p := &Properties{doc: d}
				if xmlutil.Unmarshal(cpData, &p.cp) == nil {
					d.props = p
				}
			}
		}
	}

	return d, nil
}

// syncParts marshals the in-memory document structures back into the OPC package parts.
func (d *Document) syncParts() error {
	// Write cached section properties back to the body's raw XML.
	d.syncBodySectPr()

	// Marshal document.xml.
	docData, err := marshalDocument(d.doc)
	if err != nil {
		return fmt.Errorf("docx: marshal document.xml: %w", err)
	}
	upsertPartData(d.pkg, d.docPath, ctDocument, docData)

	// Marshal styles.xml if present.
	if d.styles != nil {
		sData, err := marshalStyles(d.styles)
		if err != nil {
			return fmt.Errorf("docx: marshal styles.xml: %w", err)
		}
		stylesPath := resolveRelPath(d.docPath, "styles.xml")
		upsertPartData(d.pkg, stylesPath, ctStyles, sData)

		// Ensure part-level relationship exists.
		if docPart, ok := d.pkg.Parts[d.docPath]; ok {
			if opc.FindRelByType(docPart.Rels, opc.RelStyles) == nil {
				docPart.Rels = append(docPart.Rels, opc.Relationship{
					ID:     opc.NextRelID(docPart.Rels),
					Type:   opc.RelStyles,
					Target: "styles.xml",
				})
			}
		}

		// Ensure content type override.
		d.pkg.ContentTypes.Overrides["/"+stylesPath] = ctStyles
	}

	// Marshal comments.xml if present.
	if d.comments != nil && len(d.comments.Comments) > 0 {
		cData, err := marshalComments(d.comments)
		if err != nil {
			return fmt.Errorf("docx: marshal comments.xml: %w", err)
		}
		commentsPath := resolveRelPath(d.docPath, "comments.xml")
		upsertPartData(d.pkg, commentsPath, ctComments, cData)
	}

	// Marshal header/footer parts.
	if err := d.syncHdrFtrs(); err != nil {
		return err
	}

	// Marshal docProps/core.xml if core properties have been accessed.
	if err := d.syncCoreProperties(); err != nil {
		return err
	}

	return nil
}

// upsertPartData updates an existing part's data in-place (preserving its relationships)
// or creates a new part if it doesn't exist yet.
func upsertPartData(pkg *opc.Package, partPath, contentType string, data []byte) {
	if existing, ok := pkg.Parts[partPath]; ok {
		existing.SetData(data)
	} else {
		pkg.AddPart(partPath, contentType, data)
	}
}

// marshalDocument marshals a CT_Document to XML bytes with the XML header.
func marshalDocument(doc *wml.CT_Document) ([]byte, error) {
	data, err := xmlutil.Marshal(doc, xmlutil.OOXML)
	if err != nil {
		return nil, err
	}
	return xmlutil.AddXMLHeader(data), nil
}

// marshalStyles marshals CT_Styles to XML bytes with the XML header.
func marshalStyles(styles *wml.CT_Styles) ([]byte, error) {
	data, err := xmlutil.Marshal(styles, xmlutil.OOXML)
	if err != nil {
		return nil, err
	}
	return xmlutil.AddXMLHeader(data), nil
}

// marshalComments marshals CT_Comments to XML bytes with the XML header.
func marshalComments(comments *wml.CT_Comments) ([]byte, error) {
	data, err := xmlutil.Marshal(comments, xmlutil.OOXML)
	if err != nil {
		return nil, err
	}
	return xmlutil.AddXMLHeader(data), nil
}

// findPartRelTarget finds the target path for a relationship type on a part.
// It resolves relative targets against the part's directory.
func findPartRelTarget(part *opc.Part, relType, partPath string) string {
	rel := opc.FindRelByType(part.Rels, relType)
	if rel == nil {
		return ""
	}
	return resolveRelPath(partPath, rel.Target)
}

// resolveRelPath resolves a relative target path against the directory of
// the source part. If target is already absolute-looking, it is returned as-is.
func resolveRelPath(sourcePath, target string) string {
	if path.IsAbs(target) || len(target) == 0 {
		return target
	}
	dir := path.Dir(sourcePath)
	if dir == "." {
		return target
	}
	return path.Join(dir, target)
}

// resolveConfig returns a Config from the provided pointer, using defaults for nil.
func resolveConfig(cfg *Config) Config {
	if cfg == nil {
		return Config{}
	}
	return *cfg
}

// collectAllParagraphs returns all CT_P elements in the document body,
// including those nested inside table cells.
// Caller must hold at least a read lock.
func (d *Document) collectAllParagraphs() []*wml.CT_P {
	if d.doc.Body == nil {
		return nil
	}
	return collectParagraphsFromBlocks(d.doc.Body.Content)
}

// collectParagraphsFromBlocks recursively collects all CT_P elements
// from a slice of BlockLevelContent, including nested table cells.
func collectParagraphsFromBlocks(blocks []wml.BlockLevelContent) []*wml.CT_P {
	var result []*wml.CT_P
	for _, bc := range blocks {
		if bc.Paragraph != nil {
			result = append(result, bc.Paragraph)
		}
		if bc.Table != nil {
			for _, row := range bc.Table.Rows {
				for _, cell := range row.Cells {
					result = append(result, collectParagraphsFromBlocks(cell.Content)...)
				}
			}
		}
	}
	return result
}

// hdrFtrTypeFromRelID determines the HeaderFooterType (default/first/even) for a
// header/footer relationship by matching its ID against the sectPr references.
// Caller must have already set d.sectPr or d.doc.Body.SectPr.
func hdrFtrTypeFromRelID(d *Document, relID string, isHdr bool) HeaderFooterType {
	// Try to match against the parsed sectPr if available.
	sp := d.sectPr
	if sp == nil && d.doc.Body != nil && d.doc.Body.SectPr != nil && len(d.doc.Body.SectPr.Data) > 0 {
		var parsed wml.CT_SectPr
		if xmlutil.Unmarshal(d.doc.Body.SectPr.Data, &parsed) == nil {
			sp = &parsed
		}
	}
	if sp == nil {
		return HdrFtrDefault
	}
	refs := sp.FooterRefs
	if isHdr {
		refs = sp.HeaderRefs
	}
	for _, ref := range refs {
		if ref.ID == relID {
			switch ref.Type {
			case "first":
				return HdrFtrFirst
			case "even":
				return HdrFtrEven
			default:
				return HdrFtrDefault
			}
		}
	}
	return HdrFtrDefault
}

// syncHdrFtrs marshals all header/footer parts back to the OPC package.
// Caller must hold the write lock.
func (d *Document) syncHdrFtrs() error {
	for _, hf := range d.hdrFtrs {
		if hf.body == nil {
			continue
		}
		data, err := marshalHdrFtrBody(hf.body, hf.isHdr)
		if err != nil {
			return fmt.Errorf("docx: marshal %s: %w", hf.path, err)
		}
		ct := ctFooter
		if hf.isHdr {
			ct = ctHeader
		}
		upsertPartData(d.pkg, hf.path, ct, data)
	}
	return nil
}

// marshalHdrFtrBody marshals a header/footer body as XML with the appropriate
// root element (hdr or ftr).
func marshalHdrFtrBody(body *wml.CT_Body, isHdr bool) ([]byte, error) {
	// Header/footer XML uses <w:hdr> or <w:ftr> as root, which has the same
	// content model as <w:body>. We temporarily rename the XMLName for marshaling.
	local := "ftr"
	if isHdr {
		local = "hdr"
	}
	saved := body.XMLName
	body.XMLName = xml.Name{Space: wml.Ns, Local: local}
	data, err := xmlutil.Marshal(body, xmlutil.OOXML)
	body.XMLName = saved
	if err != nil {
		return nil, err
	}
	return xmlutil.AddXMLHeader(data), nil
}

// nextHdrFtrNum returns the next available header/footer part number by
// scanning existing part paths to avoid collisions with pre-existing files.
// Caller must hold the write lock.
func (d *Document) nextHdrFtrNum(isHdr bool) int {
	prefix := "footer"
	if isHdr {
		prefix = "header"
	}
	maxVal := 0
	for _, hf := range d.hdrFtrs {
		if hf.isHdr != isHdr {
			continue
		}
		// Extract number from path like "word/header2.xml".
		base := hf.path
		if idx := strings.LastIndex(base, "/"); idx >= 0 {
			base = base[idx+1:]
		}
		base = strings.TrimPrefix(base, prefix)
		base = strings.TrimSuffix(base, ".xml")
		if n, err := strconv.Atoi(base); err == nil && n > maxVal {
			maxVal = n
		}
	}
	return maxVal + 1
}
