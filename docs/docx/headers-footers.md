# Headers & Footers

Add, read, and manipulate page headers and footers. Each section can have up to three header and three footer variants: default (odd pages), first page, and even pages.

## Add a Header

```go
for section := range doc.Sections() {
    hdr := section.AddHeader(docx.HdrFtrDefault)
    hdr.AddParagraph().AddRun("Company Name")
}
```

## Add a Footer

```go
for section := range doc.Sections() {
    ftr := section.AddFooter(docx.HdrFtrDefault)
    ftr.AddParagraph().AddRun("Page 1 of 10")
}
```

## Header/Footer Variants

| Type | Constant | Used for |
|------|----------|----------|
| Default | `HdrFtrDefault` | All pages (or odd pages when even is set) |
| First | `HdrFtrFirst` | First page only |
| Even | `HdrFtrEven` | Even-numbered pages |

```go
section.AddHeader(docx.HdrFtrFirst)  // title page header
section.AddHeader(docx.HdrFtrEven)   // even page header
```

## Read Headers and Footers

```go
for section := range doc.Sections() {
    for hdr := range section.Headers() {
        fmt.Printf("Header (%d): %s\n", hdr.Type(), hdr.Text())
    }
    for ftr := range section.Footers() {
        fmt.Printf("Footer (%d): %s\n", ftr.Type(), ftr.Text())
    }
}
```

## Content Model

Headers and footers have the same content model as the document body — they contain paragraphs and tables.

```go
hdr := section.AddHeader(docx.HdrFtrDefault)

// Add paragraphs.
p := hdr.AddParagraph()
p.AddRun("Left text")
r := p.AddRun("Bold text")
b := true
r.SetBold(&b)

// Read content.
for p := range hdr.Paragraphs() {
    fmt.Println(p.Text())
}

// Export.
fmt.Println(hdr.Text())
fmt.Println(hdr.Markdown())
```

## Key Types

- `HeaderFooterType` — `HdrFtrDefault`, `HdrFtrFirst`, `HdrFtrEven`
- `HeaderFooter` — Type, Paragraphs, Tables, AddParagraph, Text, Markdown, ProseMirror
- `Section.Headers()` — iterate headers for the section
- `Section.Footers()` — iterate footers for the section
- `Section.AddHeader(typ)` — create a new header
- `Section.AddFooter(typ)` — create a new footer
