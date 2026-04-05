# Examples

Runnable examples demonstrating go-ooxml features. Each example creates `.docx` files in a local `output/` directory.

## Running

```bash
cd examples/docx/<example-name>
go run main.go
```

Then open the generated `.docx` files in Microsoft Word, LibreOffice, or Google Docs.

## Workflow Examples

| Example | Description | Run |
|---------|-------------|-----|
| [basic](docx/basic/) | Create a document with headings, formatted text, and a hyperlink | `cd examples/docx/basic && go run main.go` |
| [read-and-modify](docx/read-and-modify/) | Open a document, extract text, add content, save | `cd examples/docx/read-and-modify && go run main.go` |
| [review-workflow](docx/review-workflow/) | Comments + tracked changes: add, iterate, accept, reject | `cd examples/docx/review-workflow && go run main.go` |

## Feature Examples

| Example | Description | Run |
|---------|-------------|-----|
| [comments](docx/comments/) | Add, query, and remove comments | `cd examples/docx/comments && go run main.go` |
| [revisions](docx/revisions/) | Track changes, accept/reject workflow | `cd examples/docx/revisions && go run main.go` |
| [markdown](docx/markdown/) | Import and export Markdown | `cd examples/docx/markdown && go run main.go` |
| [tables](docx/tables/) | Create and populate tables | `cd examples/docx/tables && go run main.go` |
| [formatting](docx/formatting/) | Bold, italic, fonts, styles, alignment | `cd examples/docx/formatting && go run main.go` |
| [search](docx/search/) | Find, regex, replace, batch operations | `cd examples/docx/search && go run main.go` |
| [sections](docx/sections/) | Page layout, margins, columns, orientation | `cd examples/docx/sections && go run main.go` |
