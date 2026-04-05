// Package main demonstrates tracked changes: add, inspect, accept, reject, override.
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

	// Original content.
	r1 := doc.Body().AddParagraph().AddRun("The project is on track.")
	r2 := doc.Body().AddParagraph().AddRun("Budget is within limits.")
	r3 := doc.Body().AddParagraph().AddRun("Timeline needs adjustment.")

	// Add revisions (suggested changes).
	rev1 := doc.AddRevision(r1, r1, "The project is ahead of schedule.",
		docx.WithRevisionAuthor("Editor"),
	)
	rev2 := doc.AddRevision(r2, r2, "Budget is under control.",
		docx.WithRevisionAuthor("Editor"),
	)
	doc.AddRevision(r3, r3, "Timeline is on target.",
		docx.WithRevisionAuthor("Editor"),
	)

	// Inspect revisions.
	fmt.Println("--- All Revisions ---")
	for rev := range doc.Revisions() {
		switch rev.Type() {
		case docx.RevisionInsert:
			fmt.Printf("  INSERT: %q (by %s)\n", rev.ProposedText(), rev.Author())
		case docx.RevisionDelete:
			fmt.Printf("  DELETE: %q (by %s)\n", rev.OriginalText(), rev.Author())
		}
	}

	// Accept rev1 — makes the insertion permanent.
	doc.AcceptRevision(rev1)
	fmt.Println("\nAccepted revision 1")

	// Reject rev2 — restores original text.
	doc.RejectRevision(rev2)
	fmt.Println("Rejected revision 2")

	// Override the third revision — collect first, then override (iterator holds a read lock).
	var toOverride *docx.Revision
	for rev := range doc.Revisions() {
		if rev.Type() == docx.RevisionInsert && rev.ProposedText() == "Timeline is on target." {
			toOverride = rev
			break
		}
	}
	if toOverride != nil {
		doc.OverrideRevision(toOverride, "Timeline has been revised.",
			docx.WithRevisionAuthor("Senior Editor"),
		)
		fmt.Println("Overridden revision 3")
	}

	f, err := os.Create("output/revisions.docx")
	if err != nil {
		panic(err)
	}
	if _, err := doc.WriteTo(f); err != nil {
		f.Close()
		panic(err)
	}
	f.Close()
	fmt.Println("\nCreated output/revisions.docx")
	fmt.Println("\nFinal text:")
	fmt.Println(doc.Text(nil))
}
