// Package main demonstrates creating a basic .docx document with headings,
// formatted text, and a hyperlink.
package main

import (
	"fmt"
	"os"

	"github.com/ieshan/go-ooxml/docx"
)

func main() {
	os.MkdirAll("output", 0o755)

	// Create a new document.
	doc, err := docx.New(&docx.Config{Author: "go-ooxml"})
	if err != nil {
		panic(err)
	}
	defer doc.Close()

	// Add a heading.
	h := doc.Body().AddParagraph()
	h.SetStyle("Heading1")
	h.AddRun("Quarterly Report")

	// Add a paragraph with mixed formatting.
	p := doc.Body().AddParagraph()
	p.AddRun("Revenue grew ")
	r := p.AddRun("15%")
	r.SetBold(new(true))
	p.AddRun(" year-over-year, driven by ")
	r2 := p.AddRun("new product launches")
	r2.SetItalic(new(true))
	p.AddRun(".")

	// Add a subheading.
	h2 := doc.Body().AddParagraph()
	h2.SetStyle("Heading2")
	h2.AddRun("Resources")

	// Add a paragraph with a hyperlink.
	p2 := doc.Body().AddParagraph()
	p2.AddRun("For details, see the ")
	p2.AddHyperlink("https://example.com/report", "full report")
	p2.AddRun(".")

	// Write the document to a file.
	f, err := os.Create("output/basic.docx")
	if err != nil {
		panic(err)
	}
	if _, err := doc.WriteTo(f); err != nil {
		f.Close()
		panic(err)
	}
	f.Close()
	fmt.Println("Created output/basic.docx")
}
