// Package main demonstrates opening an existing .docx, extracting text,
// modifying content, and saving to a new file.
package main

import (
	"bytes"
	"fmt"
	"os"

	"github.com/ieshan/go-ooxml/docx"
)

func main() {
	os.MkdirAll("output", 0o755)

	// Step 1: Create a document to work with.
	doc, err := docx.New(&docx.Config{Author: "go-ooxml"})
	if err != nil {
		panic(err)
	}
	h := doc.Body().AddParagraph()
	h.SetStyle("Heading1")
	h.AddRun("Original Document")
	doc.Body().AddParagraph().AddRun("This is the original content.")
	f, err := os.Create("output/original.docx")
	if err != nil {
		panic(err)
	}
	if _, err := doc.WriteTo(f); err != nil {
		f.Close()
		panic(err)
	}
	f.Close()
	doc.Close()
	fmt.Println("Created output/original.docx")

	// Step 2: Reopen the document.
	doc2, err := docx.Open("output/original.docx", nil)
	if err != nil {
		panic(err)
	}
	defer doc2.Close()

	// Step 3: Extract text.
	fmt.Println("\n--- Plain Text ---")
	fmt.Println(doc2.Text(nil))

	// Step 4: Extract as Markdown.
	fmt.Println("--- Markdown ---")
	fmt.Println(doc2.Markdown(nil))

	// Step 5: Add more content.
	doc2.Body().AddParagraph().AddRun("This paragraph was added later.")

	// Step 6: Write to a new file.
	f2, err := os.Create("output/modified.docx")
	if err != nil {
		panic(err)
	}
	if _, err := doc2.WriteTo(f2); err != nil {
		f2.Close()
		panic(err)
	}
	f2.Close()
	fmt.Println("\nCreated output/modified.docx")

	// Step 7: Demonstrate Write to a buffer.
	var buf bytes.Buffer
	doc2.Write(&buf)
	fmt.Printf("Document size in memory: %d bytes\n", buf.Len())
}
