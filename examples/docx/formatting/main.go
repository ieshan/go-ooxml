// Package main demonstrates run-level formatting, paragraph styles, and alignment.
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

	// Heading styles.
	for i, text := range []string{"Title", "Subtitle", "Section"} {
		p := doc.Body().AddParagraph()
		p.SetStyle(fmt.Sprintf("Heading%d", i+1))
		p.AddRun(text)
	}

	// Bold.
	p := doc.Body().AddParagraph()
	r := p.AddRun("This is bold text.")
	r.SetBold(new(true))

	// Italic.
	p2 := doc.Body().AddParagraph()
	r2 := p2.AddRun("This is italic text.")
	r2.SetItalic(new(true))

	// Underline.
	p3 := doc.Body().AddParagraph()
	r3 := p3.AddRun("This is underlined text.")
	r3.SetUnderline(new(true))

	// Strikethrough.
	p4 := doc.Body().AddParagraph()
	r4 := p4.AddRun("This is strikethrough text.")
	r4.SetStrikethrough(new(true))

	// Custom font and size.
	p5 := doc.Body().AddParagraph()
	r5 := p5.AddRun("Custom font: Arial 18pt")
	r5.SetFontName("Arial")
	r5.SetFontSize(18.0)

	// Font color.
	p6 := doc.Body().AddParagraph()
	r6 := p6.AddRun("Red text")
	r6.SetFontColor(common.RGB(255, 0, 0))

	p7 := doc.Body().AddParagraph()
	r7 := p7.AddRun("Blue text")
	r7.SetFontColor(common.RGB(0, 0, 255))

	// Paragraph alignment.
	center := doc.Body().AddParagraph()
	center.AddRun("Centered paragraph")
	center.Format().SetAlignment("center")

	right := doc.Body().AddParagraph()
	right.AddRun("Right-aligned paragraph")
	right.Format().SetAlignment("right")

	// Special content in a run.
	p8 := doc.Body().AddParagraph()
	r8 := p8.AddRun("Before break")
	r8.AddLineBreak()
	p8.AddRun("After line break")

	f, err := os.Create("output/formatting.docx")
	if err != nil {
		panic(err)
	}
	if _, err := doc.WriteTo(f); err != nil {
		f.Close()
		panic(err)
	}
	f.Close()
	fmt.Println("Created output/formatting.docx")
}
