// Package main demonstrates table creation, population, and extraction.
package main

import (
	"fmt"
	"os"

	"github.com/ieshan/go-ooxml/docx"
)

func main() {
	os.MkdirAll("output", 0o755)

	doc, err := docx.New(nil)
	if err != nil {
		panic(err)
	}
	defer doc.Close()

	h := doc.Body().AddParagraph()
	h.SetStyle("Heading1")
	h.AddRun("Sales Data")

	// Create a 4x3 table.
	tbl := doc.Body().AddTable(4, 3)
	tbl.SetStyle("TableGrid")

	// Populate headers.
	headers := []string{"Product", "Q1 Sales", "Q2 Sales"}
	for i, h := range headers {
		tbl.Cell(0, i).AddParagraph().AddRun(h)
	}

	// Populate data rows.
	data := [][]string{
		{"Widget A", "$1,200", "$1,500"},
		{"Widget B", "$800", "$950"},
		{"Widget C", "$2,100", "$2,400"},
	}
	for r, row := range data {
		for c, val := range row {
			tbl.Cell(r+1, c).AddParagraph().AddRun(val)
		}
	}

	// Add a dynamic row.
	newRow := tbl.AddRow()
	newRow.AddCell().AddParagraph().AddRun("Widget D")
	newRow.AddCell().AddParagraph().AddRun("$500")
	newRow.AddCell().AddParagraph().AddRun("$600")

	// Set cell content using Markdown.
	tbl.Cell(0, 0).SetMarkdown("**Product**")

	// Extract as text.
	fmt.Println("--- Table Text ---")
	fmt.Println(tbl.Text())

	// Extract as Markdown.
	fmt.Println("\n--- Table Markdown ---")
	fmt.Println(tbl.Markdown())

	// Show dimensions.
	fmt.Printf("\nRows: %d, Cols: %d\n", tbl.RowCount(), tbl.ColCount())

	f, err := os.Create("output/tables.docx")
	if err != nil {
		panic(err)
	}
	if _, err := doc.WriteTo(f); err != nil {
		f.Close()
		panic(err)
	}
	f.Close()
	fmt.Println("\nCreated output/tables.docx")
}
