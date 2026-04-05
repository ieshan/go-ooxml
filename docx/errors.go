// Package docx provides a high-level API for reading and writing Office Open
// XML word-processing documents (.docx).
//
// # Creating a document
//
//	doc, err := docx.New(&docx.Config{Author: "MyApp"})
//	if err != nil {
//	    log.Fatal(err)
//	}
//	defer doc.Close()
//	doc.Body().AddParagraph().AddRun("Hello, World!")
//	doc.WriteTo(w)
//
// # Opening an existing document
//
//	doc, err := docx.Open("report.docx", nil)
//	if err != nil {
//	    log.Fatal(err)
//	}
//	defer doc.Close()
//	fmt.Println(doc.Text(nil))
//
// # Thread safety
//
// Document is safe for concurrent reads. Concurrent writes or mixed
// read/write access are serialised internally via a sync.RWMutex.
package docx

import "errors"

var (
	// ErrInvalidDoc is returned when a document's structure is missing
	// required parts or relationships (e.g., no officeDocument relation).
	ErrInvalidDoc = errors.New("ooxml: document structure is invalid")

	// ErrNotFound is returned when a requested element (comment, revision,
	// row, etc.) is not present in the document.
	ErrNotFound = errors.New("ooxml: element not found")
)
