# Getting Started

## Install

```bash
go get github.com/ieshan/go-ooxml
```

Requires Go 1.26 or later. No external dependencies.

## Core Concepts

A `.docx` file is a ZIP archive containing XML parts. This library gives you a layered Go API:

```
Document
  └── Body
        ├── Paragraph
        │     ├── Run (text with formatting)
        │     └── Hyperlink
        └── Table
              └── Row
                    └── Cell
                          └── Paragraph (nested)
```

- **Document** — the top-level handle. Open or create one, manipulate it, save it.
- **Body** — the container for all block-level content (paragraphs and tables).
- **Paragraph** — a block of text. Has a style (e.g., "Heading1") and contains runs.
- **Run** — a span of text with uniform formatting (bold, italic, font, color).
- **Table** — rows and cells. Each cell contains paragraphs.

## Create Your First Document

```go
package main

import (
	"fmt"
	"os"

	"github.com/ieshan/go-ooxml/docx"
)

func main() {
	// Create a new document.
	doc, err := docx.New(&docx.Config{Author: "Alice"})
	if err != nil {
		panic(err)
	}
	defer doc.Close()

	// Add a heading.
	heading := doc.Body().AddParagraph()
	heading.SetStyle("Heading1")
	heading.AddRun("Hello, World!")

	// Add a body paragraph with mixed formatting.
	p := doc.Body().AddParagraph()
	p.AddRun("This is ")
	bold := true
	r := p.AddRun("bold")
	r.SetBold(&bold)
	p.AddRun(" and this is ")
	italic := true
	r2 := p.AddRun("italic")
	r2.SetItalic(&italic)
	p.AddRun(".")

	// Write to disk.
	f, err := os.Create("hello.docx")
	if err != nil {
		panic(err)
	}
	if _, err := doc.WriteTo(f); err != nil {
		f.Close()
		panic(err)
	}
	if err := f.Close(); err != nil {
		panic(err)
	}
	fmt.Println("Created hello.docx")
}
```

## Open an Existing Document

```go
doc, err := docx.Open("hello.docx", nil)
if err != nil {
	panic(err)
}
defer doc.Close()

// Extract plain text.
fmt.Println(doc.Text(nil))

// Extract as Markdown.
fmt.Println(doc.Markdown(nil))

// Modify and write to a new file.
doc.Body().AddParagraph().AddRun("Added later.")
f, _ := os.Create("hello-modified.docx")
doc.WriteTo(f)
f.Close()
```

## Open from an io.Reader

If you have the bytes in memory (e.g., from an HTTP upload):

```go
doc, err := docx.OpenReader(bytes.NewReader(data), int64(len(data)), nil)
if err != nil {
	panic(err)
}
defer doc.Close()
```

## Write to an io.Writer

`WriteTo` implements `io.WriterTo` and returns the number of bytes written:

```go
var buf bytes.Buffer
n, err := doc.WriteTo(&buf)
```

Or use `Write` if you don't need the byte count:

```go
var buf bytes.Buffer
err := doc.Write(&buf)
```

## Document Properties

```go
props := doc.Properties()
props.SetTitle("Quarterly Report")
props.SetAuthor("Alice")
fmt.Println(props.Title())  // "Quarterly Report"
```

## Next Steps

- [Comments](docx/comments.md) — add review comments to text ranges
- [Revisions](docx/revisions.md) — track changes with accept/reject workflow
- [Markdown](docx/markdown.md) — convert between Markdown and .docx
- [Tables](docx/tables.md) — create and manipulate tables
- [Formatting](docx/formatting.md) — fonts, styles, sections, page layout
- [Search](docx/search.md) — find, regex, and batch operations
- [Security](security.md) — protections for untrusted documents
- [Architecture](architecture.md) — how the library is structured
