// Package main demonstrates ProseMirror JSON export and import: create a rich
// document, export to ProseMirror JSON, re-import, and create from hand-crafted JSON.
package main

import (
	"encoding/json"
	"fmt"
	"os"

	"github.com/ieshan/go-ooxml/common"
	"github.com/ieshan/go-ooxml/docx"
)

func main() {
	os.MkdirAll("output", 0o755)

	// --- Create a rich document ---
	doc, err := docx.New(&docx.Config{Author: "go-ooxml"})
	if err != nil {
		panic(err)
	}
	defer doc.Close()

	doc.Properties().SetTitle("ProseMirror Demo")

	h := doc.Body().AddParagraph()
	h.SetStyle("Heading1")
	h.AddRun("Document Title")

	p := doc.Body().AddParagraph()
	p.AddRun("This is ")
	bold := true
	r := p.AddRun("bold")
	r.SetBold(&bold)
	p.AddRun(" and ")
	italic := true
	r2 := p.AddRun("italic")
	r2.SetItalic(&italic)
	p.AddRun(" and ")
	r3 := p.AddRun("colored")
	r3.SetFontColor(common.RGB(255, 0, 0))
	r3.SetFontSize(18.0)
	r3.SetFontName("Arial")
	p.AddRun(".")

	p2 := doc.Body().AddParagraph()
	p2.AddRun("Visit ")
	p2.AddHyperlink("https://example.com", "our website")
	p2.AddRun(" for details.")

	tbl := doc.Body().AddTable(2, 2)
	tbl.Cell(0, 0).AddParagraph().AddRun("Name")
	tbl.Cell(0, 1).AddParagraph().AddRun("Value")
	tbl.Cell(1, 0).AddParagraph().AddRun("Alpha")
	tbl.Cell(1, 1).AddParagraph().AddRun("100")

	// --- Export to ProseMirror JSON ---
	pmData, err := doc.ProseMirror(nil)
	if err != nil {
		panic(err)
	}

	fmt.Println("--- ProseMirror JSON (pretty) ---")
	var pretty json.RawMessage = pmData
	formatted, _ := json.MarshalIndent(pretty, "", "  ")
	fmt.Println(string(formatted))

	// --- Re-import the JSON to a new document ---
	doc2, err := docx.FromProseMirror(pmData, nil)
	if err != nil {
		panic(err)
	}
	defer doc2.Close()

	f, err := os.Create("output/from-prosemirror.docx")
	if err != nil {
		panic(err)
	}
	if _, err := doc2.WriteTo(f); err != nil {
		f.Close()
		panic(err)
	}
	f.Close()
	fmt.Println("\nCreated output/from-prosemirror.docx")

	// --- Create from hand-crafted ProseMirror JSON ---
	handCrafted := []byte(`{
		"type": "doc",
		"attrs": {"title": "Hand-Crafted"},
		"content": [
			{
				"type": "heading",
				"attrs": {"level": 1},
				"content": [{"type": "text", "text": "Hello from ProseMirror"}]
			},
			{
				"type": "paragraph",
				"content": [
					{"type": "text", "text": "This is "},
					{"type": "text", "text": "bold and blue", "marks": [
						{"type": "bold"},
						{"type": "textStyle", "attrs": {"color": "#0000FF", "fontSize": "16pt"}}
					]},
					{"type": "text", "text": " text."}
				]
			},
			{
				"type": "bulletList",
				"content": [
					{"type": "listItem", "content": [
						{"type": "paragraph", "content": [{"type": "text", "text": "First item"}]}
					]},
					{"type": "listItem", "content": [
						{"type": "paragraph", "content": [{"type": "text", "text": "Second item"}]}
					]}
				]
			}
		]
	}`)

	doc3, err := docx.FromProseMirror(handCrafted, nil)
	if err != nil {
		panic(err)
	}
	defer doc3.Close()

	f2, err := os.Create("output/hand-crafted.docx")
	if err != nil {
		panic(err)
	}
	if _, err := doc3.WriteTo(f2); err != nil {
		f2.Close()
		panic(err)
	}
	f2.Close()
	fmt.Println("Created output/hand-crafted.docx")

	fmt.Println("\n--- Round-trip text ---")
	fmt.Println(doc2.Text(nil))
}
