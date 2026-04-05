# ProseMirror JSON Conversion

Bidirectional conversion between `.docx` documents and ProseMirror JSON. Unlike Markdown, ProseMirror JSON preserves rich formatting detail — font size, color, font family, paragraph alignment, table structure, and page layout.

Compatible with both [TipTap](https://tiptap.dev) and vanilla [ProseMirror](https://prosemirror.net) naming conventions.

## Export to ProseMirror JSON

```go
doc, _ := docx.Open("report.docx", nil)
defer doc.Close()

data, err := doc.ProseMirror(nil)
// data is a JSON-encoded PMNode of type "doc"
```

### Options

```go
data, err := doc.ProseMirror(&docx.ProseMirrorOptions{
    IncludeHeaders: true,   // include header/footer content in doc attrs
    UseTipTapNames: false,  // use vanilla ProseMirror names (strong/em/ordered_list)
})
```

By default, TipTap names are used (`bold`, `italic`, `orderedList`).

### Per-Element Export

Every content element supports `.ProseMirror()`:

```go
body.ProseMirror()            // full body as doc node
paragraph.ProseMirror()       // single paragraph node
run.ProseMirror()             // text node with marks
table.ProseMirror()           // table node
cell.ProseMirror()            // table cell node
comment.ProseMirror()         // comment body as doc node
rev.ProposedProseMirror()     // revision proposed text
rev.OriginalProseMirror()     // revision deleted text
searchResult.ProseMirror()    // matched text with marks
headerFooter.ProseMirror()    // header/footer content
```

## Create Document from ProseMirror JSON

```go
pmJSON := []byte(`{
    "type": "doc",
    "content": [
        {"type": "heading", "attrs": {"level": 1}, "content": [
            {"type": "text", "text": "Hello"}
        ]},
        {"type": "paragraph", "content": [
            {"type": "text", "text": "This is "},
            {"type": "text", "text": "bold", "marks": [{"type": "bold"}]}
        ]}
    ]
}`)

doc, err := docx.FromProseMirror(pmJSON, nil)
defer doc.Close()
```

## Import into Existing Document

```go
doc.ImportProseMirror(pmJSON)
```

## Set Paragraph/Cell Content

```go
p := doc.Body().AddParagraph()
p.SetProseMirror([]byte(`{"type":"paragraph","content":[
    {"type":"text","text":"styled","marks":[{"type":"bold"}]}
]}`))

cell.SetProseMirror([]byte(`{"type":"tableCell","content":[
    {"type":"paragraph","content":[{"type":"text","text":"cell"}]}
]}`))
```

## Naming Conventions

Both TipTap and vanilla ProseMirror names are accepted on import:

| TipTap (default export) | Vanilla ProseMirror |
|---|---|
| `bold` | `strong` |
| `italic` | `em` |
| `orderedList` | `ordered_list` |
| `bulletList` | `bullet_list` |
| `listItem` | `list_item` |
| `codeBlock` | `code_block` |
| `hardBreak` | `hard_break` |
| `horizontalRule` | `horizontal_rule` |
| `tableRow` | `table_row` |
| `tableCell` | `table_cell` |
| `tableHeader` | `table_header` |

## Supported Marks

| Mark | Attrs | Maps to OOXML |
|---|---|---|
| `bold` / `strong` | — | `<w:b/>` |
| `italic` / `em` | — | `<w:i/>` |
| `underline` | — | `<w:u/>` |
| `strike` | — | `<w:strike/>` |
| `code` | — | Courier New font |
| `link` | `href` | `<w:hyperlink>` |
| `textStyle` | `color`, `fontSize`, `fontFamily` | `<w:color>`, `<w:sz>`, `<w:rFonts>` |

## Supported Nodes

| Node | Attrs | Maps to OOXML |
|---|---|---|
| `doc` | `title`, `author`, `pageWidth`, `pageHeight`, `orientation`, `margins`, `columns` | Document + section properties |
| `paragraph` | `textAlign`, `styleId` | `<w:p>` |
| `heading` | `level` (1-6) | Heading1-Heading6 styles |
| `blockquote` | — | Quote style |
| `codeBlock` | — | Courier New font paragraphs |
| `bulletList` / `orderedList` | `start`, `listStyleType` | List styles |
| `listItem` | — | List paragraph |
| `table` | `tableStyle` | `<w:tbl>` |
| `tableRow` | — | `<w:tr>` |
| `tableCell` / `tableHeader` | `colspan`, `rowspan`, `background`, `verticalAlign` | `<w:tc>` |
| `hardBreak` | — | `<w:br/>` |
| `pageBreak` | — | `<w:br type="page"/>` |
| `horizontalRule` | — | Empty separator paragraph |

## Document Properties in JSON

Section and document properties appear as attrs on the `doc` node:

```json
{
    "type": "doc",
    "attrs": {
        "title": "Report",
        "author": "Alice",
        "pageWidth": 612,
        "pageHeight": 792,
        "orientation": "landscape",
        "margins": {"top": 72, "right": 72, "bottom": 72, "left": 72},
        "columns": {"count": 2, "space": 36}
    },
    "content": [...]
}
```

## Limitations

- Images are exported as placeholder text (full DrawingML support not yet implemented)
- Floating image positioning not supported
- Numbered list level text patterns (`%1.%2.%3`) simplified to list-style-type
- Art borders and animated text effects not supported
- Office Math not supported
