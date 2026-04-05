# Formatting

Control text appearance, paragraph styles, and page layout.

## Run-Level Formatting

A **Run** is a span of text with uniform formatting. Set properties using pointer arguments — `nil` clears (inherits from style), non-nil sets explicitly.

```go
p := doc.Body().AddParagraph()
r := p.AddRun("styled text")

bold := true
r.SetBold(&bold)
r.SetItalic(&bold)
r.SetStrikethrough(&bold)
r.SetUnderline(&bold)

r.SetFontName("Arial")
r.SetFontSize(14.0)            // points
r.SetFontColor(common.RGB(255, 0, 0))  // red
```

### Reading Format State

```go
if r.Bold() != nil && *r.Bold() {
	fmt.Println("run is bold")
}
if name := r.FontName(); name != nil {
	fmt.Println("font:", *name)
}
if size := r.FontSize(); size != nil {
	fmt.Printf("size: %.1f pt\n", *size)
}
```

### Special Content

```go
r.AddLineBreak()
r.AddPageBreak()
r.AddTab()
```

## Paragraph Styles

```go
p := doc.Body().AddParagraph()
p.SetStyle("Heading1")        // built-in heading
p.SetStyle("Quote")           // blockquote
p.SetStyle("ListBullet")      // bullet list item
fmt.Println(p.Style())        // "Heading1"
```

## Paragraph Alignment

```go
pf := p.Format()
pf.SetAlignment("center")     // "left", "right", "center", "both"
fmt.Println(pf.Alignment())   // "center"
```

## Sections and Page Layout

Sections control page dimensions, margins, orientation, and columns.

```go
for section := range doc.Sections() {
	w, h := section.PageSize()
	fmt.Printf("Page: %.1f x %.1f inches\n", w.Inches(), h.Inches())
}
```

### Set Page Size and Orientation

```go
for section := range doc.Sections() {
	section.SetPageSize(common.Inches(8.5), common.Inches(11))
	section.SetOrientation(docx.OrientLandscape)
}
```

### Margins

```go
section.SetMargins(docx.SectionMargins{
	Top:    common.Inches(1),
	Right:  common.Inches(1),
	Bottom: common.Inches(1),
	Left:   common.Inches(1.5),
})
```

### Multi-Column Layout

```go
section.SetColumns(&docx.ColumnLayout{
	Count:      2,
	Space:      common.Inches(0.5),
	EqualWidth: true,
	Separator:  true,
})
```

### Section Breaks

```go
section.SetBreakType(docx.BreakContinuous)
// Options: BreakNextPage, BreakContinuous, BreakEvenPage, BreakOddPage
```

## Measurement Units

The `common` package provides length conversions. All OOXML measurements use EMU (English Metric Units) internally.

```go
import "github.com/ieshan/go-ooxml/common"

l := common.Inches(1.5)
fmt.Println(l.Cm())     // ~3.81
fmt.Println(l.Pt())     // 108.0
fmt.Println(l.Twips())  // 2160
fmt.Println(l.EMU())    // 1371600
```
