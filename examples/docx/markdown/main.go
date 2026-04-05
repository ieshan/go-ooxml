// Package main demonstrates Markdown conversion: create from Markdown,
// export to Markdown, and import Markdown into an existing document.
package main

import (
	"fmt"
	"os"

	"github.com/ieshan/go-ooxml/docx"
)

func main() {
	os.MkdirAll("output", 0o755)

	// --- Create a document from Markdown ---
	md := `# Project Status

This is a **critical** update with *important* details.

## Highlights

- Revenue up ~~10%~~ 15%
- Customer count: ` + "`" + `1,234` + "`" + `
- See [details](https://example.com)

## Data

| Metric   | Q1   | Q2   |
|----------|------|------|
| Revenue  | $10M | $12M |
| Users    | 500  | 750  |

> This report is confidential.

---

End of report.`

	doc, err := docx.FromMarkdown(md, &docx.Config{Author: "go-ooxml"})
	if err != nil {
		panic(err)
	}
	defer doc.Close()

	f, err := os.Create("output/from-markdown.docx")
	if err != nil {
		panic(err)
	}
	if _, err := doc.WriteTo(f); err != nil {
		f.Close()
		panic(err)
	}
	f.Close()
	fmt.Println("Created output/from-markdown.docx")

	// --- Export back to Markdown ---
	fmt.Println("\n--- Exported Markdown ---")
	exported := doc.Markdown(nil)
	fmt.Println(exported)

	// --- Export with comments as footnotes ---
	r := doc.Body().AddParagraph().AddRun("Added after import")
	doc.AddComment(r, r, "Review this addition", docx.WithAuthor("Reviewer"))

	mdWithComments := doc.Markdown(&docx.MarkdownOptions{IncludeComments: true})
	fmt.Println("\n--- With Comments ---")
	fmt.Println(mdWithComments)

	// --- Import Markdown into existing document ---
	doc.ImportMarkdown("\n## Appendix\n\nAdditional notes here.")
	f2, err := os.Create("output/markdown-appended.docx")
	if err != nil {
		panic(err)
	}
	if _, err := doc.WriteTo(f2); err != nil {
		f2.Close()
		panic(err)
	}
	f2.Close()
	fmt.Println("\nCreated output/markdown-appended.docx")
}
