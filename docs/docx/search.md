# Search

Find text in documents with plain text or regex, then apply batch operations.

## Basic Search

```go
results := doc.Find("revenue")
for _, r := range results {
	fmt.Printf("Found %q in paragraph: %s\n", r.Text(), r.Paragraph.Text())
}
```

`Find` is case-insensitive by default and searches inside table cells.

## Regex Search

```go
pattern := regexp.MustCompile(`\d{3}-\d{4}`)
results := doc.FindRegex(pattern)
for _, r := range results {
	fmt.Println("Phone:", r.Text())
}
```

## Fluent Search API

`Search()` returns a `SearchQuery` with chainable filters:

```go
results := doc.Search("important").
	CaseSensitive(true).
	InStyle("Heading1").
	Results()
```

### Filters

| Method | Description |
|--------|-------------|
| `CaseSensitive(true)` | Exact case matching |
| `InStyle("Heading1")` | Only paragraphs with this style |
| `First()` | Return only the first match |
| `Nth(2)` | Return only the 3rd match (0-based) |

## Batch Operations

Apply changes to all matched text in one call:

### Replace Text

```go
count := doc.Search("old name").ReplaceText("new name")
fmt.Printf("Replaced %d occurrences\n", count)
```

Note: `ReplaceText` does not preserve the original formatting of the matched text.

### Replace Text (Format-Preserving)

Replace text while keeping the original formatting. The replacement inherits the union of all formatting from the matched runs — if any matched run is bold, the entire replacement is bold.

```go
// "ACME" is bold in the document. After replace, "SuperWidget" is also bold.
count := doc.Search("ACME").ReplaceTextFormatted("SuperWidget")
```

### Replace with Markdown (Format-Preserving)

Replace with inline Markdown while preserving original formatting as a base. Markdown formatting is additive — it merges with the inherited formatting.

```go
// "old text" is italic. After replace:
// "new" is bold+italic (markdown bold + inherited italic)
// " text" is italic (inherited only)
doc.Search("old text").ReplaceMarkdown("**new** text")
```

Supported inline Markdown: `**bold**`, `*italic*`, `~~strike~~`, `` `code` ``, `***bold+italic***`.

### Apply Formatting

```go
bold := true
doc.Search("WARNING").SetBold(&bold)
doc.Search("note").SetItalic(&bold)
```

### Set Paragraph Style

```go
doc.Search("Summary").InStyle("").SetStyle("Heading2")
```

### Add Comments to Matches

```go
doc.Search("TODO").AddComment("Please resolve",
	docx.WithAuthor("Bot"),
)
```

### Add Revisions to Matches

```go
doc.Search("draft").AddRevision("final",
	docx.WithRevisionAuthor("Editor"),
)
```

## SearchResult

Each result provides:

```go
result.Paragraph  // *Paragraph containing the match
result.Runs       // []*Run spanning the match
result.Start      // character offset in first run
result.End        // character offset in last run
result.Text()     // the matched text
result.Markdown()  // the matched text as Markdown (preserving bold/italic/etc.)
```
