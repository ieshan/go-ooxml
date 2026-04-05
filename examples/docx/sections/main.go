// Package main demonstrates page layout: size, orientation, margins, columns.
package main

import (
	"fmt"
	"os"

	"github.com/ieshan/go-ooxml/common"
	"github.com/ieshan/go-ooxml/docx"
)

func main() {
	os.MkdirAll("output", 0o755)

	doc, err := docx.New(nil)
	if err != nil {
		panic(err)
	}
	defer doc.Close()

	// Add content.
	h := doc.Body().AddParagraph()
	h.SetStyle("Heading1")
	h.AddRun("Page Layout Demo")

	doc.Body().AddParagraph().AddRun(
		"This document demonstrates page size, margins, orientation, and column layout.",
	)

	doc.Body().AddParagraph().AddRun(
		"Lorem ipsum dolor sit amet, consectetur adipiscing elit. " +
			"Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. " +
			"Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris.",
	)

	// Configure the section (page layout).
	for section := range doc.Sections() {
		// Set page size to US Letter.
		section.SetPageSize(common.Inches(8.5), common.Inches(11))

		// Set landscape orientation.
		section.SetOrientation(docx.OrientLandscape)

		// Set custom margins.
		section.SetMargins(docx.SectionMargins{
			Top:    common.Inches(0.75),
			Right:  common.Inches(0.75),
			Bottom: common.Inches(0.75),
			Left:   common.Inches(1.0),
		})

		// Set two-column layout.
		section.SetColumns(&docx.ColumnLayout{
			Count:      2,
			Space:      common.Inches(0.5),
			EqualWidth: true,
			Separator:  true,
		})

		// Print current settings.
		w, ht := section.PageSize()
		fmt.Printf("Page size: %.1f x %.1f inches\n", w.Inches(), ht.Inches())
		fmt.Printf("Orientation: %v\n", section.Orientation())
		m := section.Margins()
		fmt.Printf("Margins: T=%.2f R=%.2f B=%.2f L=%.2f inches\n",
			m.Top.Inches(), m.Right.Inches(), m.Bottom.Inches(), m.Left.Inches())
		if cols := section.Columns(); cols != nil {
			fmt.Printf("Columns: %d (space=%.2f in)\n", cols.Count, cols.Space.Inches())
		}
	}

	f, err := os.Create("output/sections.docx")
	if err != nil {
		panic(err)
	}
	if _, err := doc.WriteTo(f); err != nil {
		f.Close()
		panic(err)
	}
	f.Close()
	fmt.Println("\nCreated output/sections.docx")
}
