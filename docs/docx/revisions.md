# Revisions (Tracked Changes)

Add, inspect, accept, and reject suggested changes in a document.

## Add a Revision

A revision wraps the original text in `<w:del>` and the proposed replacement in `<w:ins>`:

```go
doc, _ := docx.New(&docx.Config{Author: "Editor"})
defer doc.Close()

r := doc.Body().AddParagraph().AddRun("draft text")
rev := doc.AddRevision(r, r, "final text")
```

The returned `*Revision` represents the insertion side. Both the deletion and insertion are visible when iterating.

## Options

```go
doc.AddRevision(startRun, endRun, "replacement",
	docx.WithRevisionAuthor("Jane"),
	docx.WithRevisionDate(time.Date(2025, 6, 1, 0, 0, 0, 0, time.UTC)),
)
```

### Markdown-Styled Revisions

```go
doc.AddRevision(r, r, "**bold** replacement", docx.WithMarkdownRevision())
```

## Iterate Over All Revisions

```go
for rev := range doc.Revisions() {
	switch rev.Type() {
	case docx.RevisionInsert:
		fmt.Printf("Insert by %s: %q\n", rev.Author(), rev.ProposedText())
	case docx.RevisionDelete:
		fmt.Printf("Delete by %s: %q\n", rev.Author(), rev.OriginalText())
	}
}
```

## Accept a Revision

Accepting makes the change permanent:

- **Insert revision** — the inserted runs become normal paragraph content.
- **Delete revision** — the deleted content is removed entirely.

```go
for rev := range doc.Revisions() {
	if rev.Type() == docx.RevisionInsert {
		doc.AcceptRevision(rev)
		break
	}
}
```

## Reject a Revision

Rejecting restores the original text:

- **Insert revision** — the insertion is removed.
- **Delete revision** — the deleted text is restored as normal content.

```go
doc.RejectRevision(rev)
```

## Override a Revision

Replace the proposed text of an existing revision:

```go
rev2 := doc.OverrideRevision(rev, "even better text")
```

## Find Revisions by Text

```go
for rev := range doc.RevisionsForText("important") {
	fmt.Println(rev.ProposedText())
}
```

## Remove a Revision

Remove both the insertion and deletion without accepting or rejecting:

```go
doc.RemoveRevision(rev)
```

## Markdown Access

```go
fmt.Println(rev.OriginalMarkdown())  // deleted text with formatting
fmt.Println(rev.ProposedMarkdown())  // inserted text with formatting
```

## Key Types

- `RevisionType` — `RevisionInsert`, `RevisionDelete`
- `Revision` — ID, Author, Date, Type, OriginalText, ProposedText, OriginalMarkdown, ProposedMarkdown
- `RevisionOption` — WithRevisionAuthor, WithRevisionDate, WithMarkdownRevision
