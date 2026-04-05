// Package main demonstrates text search, regex search, and batch operations.
package main

import (
	"fmt"
	"os"
	"regexp"

	"github.com/ieshan/go-ooxml/docx"
)

func main() {
	os.MkdirAll("output", 0o755)

	doc, err := docx.New(&docx.Config{Author: "Demo"})
	if err != nil {
		panic(err)
	}
	defer doc.Close()

	// Build a document with searchable content.
	h := doc.Body().AddParagraph()
	h.SetStyle("Heading1")
	h.AddRun("Contact Information")

	doc.Body().AddParagraph().AddRun("Call us at 555-1234 or 555-5678.")
	doc.Body().AddParagraph().AddRun("Email: support@example.com")
	doc.Body().AddParagraph().AddRun("The old product name was ACME Widget.")
	doc.Body().AddParagraph().AddRun("The ACME Widget is our best seller.")
	doc.Body().AddParagraph().AddRun("TODO: update pricing page.")
	doc.Body().AddParagraph().AddRun("TODO: review shipping policy.")

	// --- Basic Find ---
	fmt.Println("--- Find 'ACME' ---")
	results := doc.Find("acme") // case-insensitive by default
	for _, r := range results {
		fmt.Printf("  Found: %q\n", r.Text())
	}

	// --- Regex Find ---
	fmt.Println("\n--- Regex: phone numbers ---")
	pattern := regexp.MustCompile(`\d{3}-\d{4}`)
	phones := doc.FindRegex(pattern)
	for _, r := range phones {
		fmt.Printf("  Phone: %s\n", r.Text())
	}

	// --- Batch Replace ---
	count := doc.Search("ACME Widget").ReplaceText("SuperWidget Pro")
	fmt.Printf("\nReplaced %d occurrences of 'ACME Widget'\n", count)

	// --- Batch Formatting ---
	bold := true
	n := doc.Search("TODO").SetBold(&bold)
	fmt.Printf("Made %d 'TODO' runs bold\n", n)

	// --- Add Comments to Matches ---
	comments := doc.Search("SuperWidget Pro").AddComment("Verify new branding",
		docx.WithAuthor("Bot"),
	)
	fmt.Printf("Added %d comments on 'SuperWidget Pro'\n", len(comments))

	// --- Format-Preserving Replace ---
	// Make "SuperWidget Pro" bold, then replace while preserving bold.
	boldTrue := true
	doc.Search("SuperWidget Pro").SetBold(&boldTrue)
	n2 := doc.Search("SuperWidget Pro").ReplaceTextFormatted("MegaWidget")
	fmt.Printf("Format-preserving replace: %d occurrences (bold preserved)\n", n2)

	// --- Replace with Markdown (inherits formatting) ---
	// "MegaWidget" is bold. Replace with markdown — bold is inherited,
	// and markdown adds italic to "Ultra".
	n3 := doc.Search("MegaWidget").ReplaceMarkdown("*Ultra* Widget")
	fmt.Printf("Markdown replace: %d occurrences (bold inherited + italic from markdown)\n", n3)

	f, err := os.Create("output/search.docx")
	if err != nil {
		panic(err)
	}
	if _, err := doc.WriteTo(f); err != nil {
		f.Close()
		panic(err)
	}
	f.Close()
	fmt.Println("\nCreated output/search.docx")
	fmt.Println("\n--- Final Text ---")
	fmt.Println(doc.Text(nil))
}
