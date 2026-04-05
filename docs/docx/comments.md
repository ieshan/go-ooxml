# Comments

Add, query, and remove review comments anchored to specific text ranges.

## Add a Comment

```go
doc, _ := docx.New(&docx.Config{Author: "Alice"})
defer doc.Close()

r := doc.Body().AddParagraph().AddRun("Check this claim")
comment := doc.AddComment(r, r, "Needs a citation")
```

The comment is anchored between `startRun` and `endRun`. They can be the same run (point comment) or different runs in the same or different paragraphs (range comment).

## Options

```go
doc.AddComment(startRun, endRun, "review note",
	docx.WithAuthor("Jane Doe"),
	docx.WithDate(time.Now()),
	docx.WithInitials("JD"),
)
```

### Markdown-Styled Comments

Pass `WithMarkdown()` to parse the comment text as inline Markdown:

```go
doc.AddComment(r, r, "This is **critical** — fix before release", docx.WithMarkdown())
```

This creates bold runs inside the comment body rather than storing literal `**` characters.

## Iterate Over All Comments

```go
for c := range doc.Comments() {
	fmt.Printf("[%s] %s: %s\n", c.Date().Format(time.RFC3339), c.Author(), c.Text())
}
```

## Find Comments by Text

Find comments whose anchored text range contains a substring:

```go
for c := range doc.CommentsForText("revenue") {
	fmt.Printf("Comment on 'revenue': %s\n", c.Text())
}
```

This works across paragraph boundaries — if a comment spans two paragraphs, the text from both is searched.

## Remove a Comment

```go
for c := range doc.Comments() {
	if c.Author() == "Old Reviewer" {
		doc.RemoveComment(c)
		break
	}
}
```

Removing a comment also removes its range markers (`commentRangeStart`, `commentRangeEnd`, `commentReference`) from the document body.

## Comment Paragraphs

A comment body can contain multiple paragraphs:

```go
for p := range comment.Paragraphs() {
	fmt.Println(p.Text())
	fmt.Println(p.Markdown())
}
```

## Search + Comment

Use the fluent search API to add comments to matched text:

```go
doc.Search("TODO").AddComment("Please resolve", docx.WithAuthor("Bot"))
```

## Key Types

- `Comment` — ID, Author, Date, Initials, Text, Markdown, Paragraphs
- `CommentOption` — WithAuthor, WithDate, WithInitials, WithMarkdown
- `Document.AddComment(startRun, endRun, text, ...opts)` -> `*Comment`
- `Document.RemoveComment(comment)` -> `error`
- `Document.Comments()` -> `iter.Seq[*Comment]`
- `Document.CommentsForText(text)` -> `iter.Seq[*Comment]`
