// Package main demonstrates a complete review workflow: adding comments,
// tracked changes, then accepting and rejecting revisions.
package main

import (
	"fmt"
	"os"

	"github.com/ieshan/go-ooxml/docx"
)

func main() {
	os.MkdirAll("output", 0o755)

	doc, err := docx.New(&docx.Config{Author: "Author"})
	if err != nil {
		panic(err)
	}
	defer doc.Close()

	// Write initial content.
	h := doc.Body().AddParagraph()
	h.SetStyle("Heading1")
	h.AddRun("Draft Report")

	p1 := doc.Body().AddParagraph()
	r1 := p1.AddRun("Revenue increased significantly last quarter.")

	p2 := doc.Body().AddParagraph()
	r2 := p2.AddRun("We plan to expand into new markets next year.")

	p3 := doc.Body().AddParagraph()
	r3 := p3.AddRun("The team performed adequately.")

	// Add comments.
	doc.AddComment(r1, r1, "Can we add specific numbers?",
		docx.WithAuthor("Reviewer"),
		docx.WithInitials("R"),
	)
	doc.AddComment(r2, r2, "Which markets?", docx.WithAuthor("Reviewer"))
	fmt.Println("Added 2 comments")

	// Add tracked changes (revisions).
	doc.AddRevision(r3, r3, "The team performed exceptionally well.",
		docx.WithRevisionAuthor("Editor"),
	)
	fmt.Println("Added 1 revision")

	// Write with comments and revisions.
	f, err := os.Create("output/review-pending.docx")
	if err != nil {
		panic(err)
	}
	if _, err := doc.WriteTo(f); err != nil {
		f.Close()
		panic(err)
	}
	f.Close()
	fmt.Println("Created output/review-pending.docx")

	// List all comments.
	fmt.Println("\n--- Comments ---")
	for c := range doc.Comments() {
		fmt.Printf("  [%s] %s\n", c.Author(), c.Text())
	}

	// List all revisions.
	fmt.Println("\n--- Revisions ---")
	for rev := range doc.Revisions() {
		switch rev.Type() {
		case docx.RevisionInsert:
			fmt.Printf("  INSERT by %s: %q\n", rev.Author(), rev.ProposedText())
		case docx.RevisionDelete:
			fmt.Printf("  DELETE by %s: %q\n", rev.Author(), rev.OriginalText())
		}
	}

	// Accept the revision — collect first, then accept (iterator holds a read lock).
	var toAccept *docx.Revision
	for rev := range doc.Revisions() {
		if rev.Type() == docx.RevisionInsert {
			toAccept = rev
			break
		}
	}
	if toAccept != nil {
		doc.AcceptRevision(toAccept)
		fmt.Println("\nAccepted revision")
	}

	// Write the final version.
	f2, err := os.Create("output/review-accepted.docx")
	if err != nil {
		panic(err)
	}
	if _, err := doc.WriteTo(f2); err != nil {
		f2.Close()
		panic(err)
	}
	f2.Close()
	fmt.Println("Created output/review-accepted.docx")
	fmt.Println("\nFinal text:")
	fmt.Println(doc.Text(nil))
}
