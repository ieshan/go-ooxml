# go-ooxml

A Go library for reading, writing, and manipulating Microsoft Office documents (OOXML).

Currently supports `.docx` (Word) format. Designed for future expansion to `.pptx` and `.xlsx`.

## Features

- **Read and write** `.docx` files with round-trip fidelity
- **Manipulate content** — paragraphs, runs, tables, hyperlinks, images, sections
- **Comments** — add, remove, and query comments anchored to text ranges
- **Tracked changes** — add, accept, reject, and override revisions
- **Markdown** — bidirectional conversion between `.docx` and Markdown
- **Text extraction** — plain text and Markdown from any level (document, paragraph, cell)
- **Search** — full-text and regex search with a fluent API for batch operations
- **Security** — ZIP bomb, XXE, path traversal, and XML bomb protections
- **Thread-safe** — concurrent reads safe; writes acquire exclusive lock
- **Zero dependencies** — stdlib only

## Install

```bash
go get github.com/ieshan/go-ooxml
```

Requires Go 1.26+.

## Quick Start

### Create a document

```go
package main

import (
	"os"

	"github.com/ieshan/go-ooxml/docx"
)

func main() {
	doc, _ := docx.New(&docx.Config{Author: "Alice"})
	defer doc.Close()

	h := doc.Body().AddParagraph()
	h.SetStyle("Heading1")
	h.AddRun("Quarterly Report")

	p := doc.Body().AddParagraph()
	p.AddRun("Revenue grew ")
	bold := true
	r := p.AddRun("15%")
	r.SetBold(&bold)
	p.AddRun(" year-over-year.")

	f, _ := os.Create("report.docx")
	doc.WriteTo(f)
	f.Close()
}
```

### Open and extract text

```go
doc, _ := docx.Open("report.docx", nil)
defer doc.Close()

fmt.Println(doc.Text(nil))
fmt.Println(doc.Markdown(nil))
```

### Add a comment

```go
doc, _ := docx.New(&docx.Config{Author: "Reviewer"})
defer doc.Close()

r := doc.Body().AddParagraph().AddRun("TODO: verify numbers")
doc.AddComment(r, r, "Please double-check these figures",
	docx.WithAuthor("Jane"),
	docx.WithInitials("J"),
)
f, _ := os.Create("reviewed.docx")
doc.WriteTo(f)
f.Close()
```

### Markdown to DOCX

```go
md := "# Report\n\nThis is **important** content with a [link](https://example.com)."
doc, _ := docx.FromMarkdown(md, nil)
defer doc.Close()
f, _ := os.Create("from-markdown.docx")
doc.WriteTo(f)
f.Close()
```

## Security

Documents from untrusted sources are handled safely by default. The library protects against ZIP bombs (compression ratio and size limits), XML external entity (XXE) attacks, path traversal in ZIP entries, and duplicate ZIP entries. Limits are configurable via `opc.SecurityLimits`.

See [docs/security.md](docs/security.md) for details.

## Documentation

| Topic | Description |
|-------|-------------|
| [Getting Started](docs/getting-started.md) | Install, core concepts, first document |
| [Architecture](docs/architecture.md) | Package layers, round-trip fidelity, thread safety |
| [Security](docs/security.md) | ZIP bomb, XXE, path traversal protections |
| [Comments](docs/docx/comments.md) | Add, query, and remove comments |
| [Revisions](docs/docx/revisions.md) | Track changes, accept/reject workflow |
| [Markdown](docs/docx/markdown.md) | Bidirectional Markdown conversion |
| [Tables](docs/docx/tables.md) | Create and manipulate tables |
| [Formatting](docs/docx/formatting.md) | Fonts, styles, sections, page layout |
| [Search](docs/docx/search.md) | Find, regex, and fluent batch operations |

## Examples

Runnable examples are in the [`examples/`](examples/) directory. Each writes `.docx` files to a local `output/` folder.

```bash
cd examples/docx/basic && go run main.go
```

See [examples/README.md](examples/README.md) for the full list.

## License

TBD
