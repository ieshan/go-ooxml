# Markdown Conversion

Bidirectional conversion between `.docx` documents and Markdown.

## Export to Markdown

```go
doc, _ := docx.Open("report.docx", nil)
defer doc.Close()

md := doc.Markdown(nil)
fmt.Println(md)
```

### Options

```go
md := doc.Markdown(&docx.MarkdownOptions{
	IncludeComments: true,   // append comments as footnotes
	HorizontalRule:  "***",  // custom HR string (default: "---")
})
```

### Per-Element Markdown

Every content element supports `.Markdown()`:

```go
doc.Body().Markdown()        // full body
paragraph.Markdown()         // single paragraph
run.Markdown()               // single run
table.Markdown()             // pipe table
cell.Markdown()              // cell content
comment.Markdown()           // comment body
rev.ProposedMarkdown()       // revision proposed text
rev.OriginalMarkdown()       // revision deleted text
```

## Create Document from Markdown

```go
doc, _ := docx.FromMarkdown("# Hello\n\nThis is **bold**.", nil)
defer doc.Close()
f, _ := os.Create("from-markdown.docx")
doc.WriteTo(f)
f.Close()
```

## Import Markdown into Existing Document

```go
doc, _ := docx.Open("existing.docx", nil)
defer doc.Close()

doc.ImportMarkdown("## New Section\n\nAppended content.")
f, _ := os.Create("updated.docx")
doc.WriteTo(f)
f.Close()
```

## Set Paragraph Content from Markdown

```go
p := doc.Body().AddParagraph()
p.SetMarkdown("This has **bold** and *italic* text")
```

## Set Cell Content from Markdown

```go
tbl := doc.Body().AddTable(1, 1)
tbl.Cell(0, 0).SetMarkdown("**Header** content")
```

## Insert Markdown After a Paragraph

```go
p := doc.Body().AddParagraph()
p.AddRun("existing content")
newParagraphs, _ := p.InsertMarkdownAfter("## Inserted\n\nNew content here")
```

## Supported Syntax

| Syntax | Example | OOXML Mapping |
|--------|---------|---------------|
| Headings | `# H1` through `###### H6` | Heading1-Heading6 styles |
| Bold | `**text**` | `<w:b/>` |
| Italic | `*text*` | `<w:i/>` |
| Bold+Italic | `***text***` | `<w:b/>` + `<w:i/>` |
| Strikethrough | `~~text~~` | `<w:strike/>` |
| Inline code | `` `code` `` | Courier New font |
| Links | `[text](url)` | `<w:hyperlink>` |
| Bullet lists | `- item` | ListBullet style |
| Numbered lists | `1. item` | ListNumber style |
| Blockquotes | `> text` | Quote style |
| Tables | `\| A \| B \|` | `<w:tbl>` |
| Code fences | ` ``` ` | Courier New font paragraphs |
| Horizontal rules | `---`, `***`, `___` | Empty separator paragraph |

## Limitations

- Nested inline formatting (e.g., `~~**bold strike**~~`) is not supported
- Images in Markdown are not converted
- List indentation levels are simplified
- Code fence language identifiers are ignored
