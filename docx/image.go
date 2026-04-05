package docx

import (
	"iter"

	"github.com/ieshan/go-ooxml/common"
)

// Image represents an image embedded in the document as an OPC part.
//
// Note: full image support in OOXML requires parsing DrawingML elements
// inside runs. This type provides the foundation for future implementation.
type Image struct {
	doc         *Document
	contentType string
	partPath    string
	width       common.Length
	height      common.Length
}

// ContentType returns the MIME content type of the image (e.g. "image/png").
func (img *Image) ContentType() string {
	return img.contentType
}

// Data reads and returns the raw image bytes from the OPC package part.
func (img *Image) Data() ([]byte, error) {
	img.doc.mu.RLock()
	defer img.doc.mu.RUnlock()

	part, ok := img.doc.pkg.Parts[img.partPath]
	if !ok {
		return nil, nil
	}
	return part.Data()
}

// Images returns an iterator over images embedded in the run.
//
// Images in OOXML are represented as DrawingML elements inside runs, which
// requires significant additional parsing. This is a stub that always returns
// an empty iterator; the type infrastructure is provided for future use.
func (r *Run) Images() iter.Seq[*Image] {
	return func(yield func(*Image) bool) {
		// Stub: DrawingML parsing is not yet implemented.
	}
}
