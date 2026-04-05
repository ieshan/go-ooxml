// Package main demonstrates comment features: add with options,
// query by text, markdown-styled comments, and removal.
package main

import (
	"fmt"
	"os"
	"time"

	"github.com/ieshan/go-ooxml/docx"
)

func main() {
	os.MkdirAll("output", 0o755)

	doc, err := docx.New(&docx.Config{Author: "Demo"})
	if err != nil {
		panic(err)
	}
	defer doc.Close()

	// Add paragraphs to comment on.
	r1 := doc.Body().AddParagraph().AddRun("The revenue figures need verification.")
	r2 := doc.Body().AddParagraph().AddRun("Customer satisfaction improved by 20%.")
	r3 := doc.Body().AddParagraph().AddRun("TODO: add executive summary.")

	// Basic comment.
	doc.AddComment(r1, r1, "Please double-check Q3 numbers")
	fmt.Println("Added basic comment")

	// Comment with all options.
	doc.AddComment(r2, r2, "Great improvement!",
		docx.WithAuthor("Jane Doe"),
		docx.WithDate(time.Date(2025, 6, 15, 10, 30, 0, 0, time.UTC)),
		docx.WithInitials("JD"),
	)
	fmt.Println("Added comment with options")

	// Markdown-styled comment.
	doc.AddComment(r3, r3, "This is **critical** — needs to be done before release",
		docx.WithAuthor("Bot"),
		docx.WithMarkdown(),
	)
	fmt.Println("Added markdown-styled comment")

	// Query comments for text containing "revenue".
	fmt.Println("\n--- Comments on 'revenue' ---")
	for c := range doc.CommentsForText("revenue") {
		fmt.Printf("  %s: %s\n", c.Author(), c.Text())
	}

	// List all comments.
	fmt.Println("\n--- All Comments ---")
	for c := range doc.Comments() {
		fmt.Printf("  [ID=%d] %s (%s): %s\n", c.ID(), c.Author(), c.Initials(), c.Text())
	}

	// Remove a comment — collect first, then remove (iterator holds a read lock).
	var toRemove *docx.Comment
	for c := range doc.Comments() {
		if c.Author() == "Bot" {
			toRemove = c
			break
		}
	}
	if toRemove != nil {
		doc.RemoveComment(toRemove)
		fmt.Println("\nRemoved Bot's comment")
	}

	f, err := os.Create("output/comments.docx")
	if err != nil {
		panic(err)
	}
	if _, err := doc.WriteTo(f); err != nil {
		f.Close()
		panic(err)
	}
	f.Close()
	fmt.Println("Created output/comments.docx")
}
