# Architecture

## Package Layers

```
┌──────────────────────────────────────────┐
│  Your Code                               │
├──────────────────────────────────────────┤
│  docx/  — Document API                   │
│  (Document, Paragraph, Run, Comment...)  │
├──────────────────────────────────────────┤
│  docx/wml/  — WordprocessingML structs   │
│  (CT_P, CT_R, CT_Comment, CT_Ins...)     │
├──────────────────────────────────────────┤
│  xmlutil/  — XML safety + namespaces     │
├──────────────────────────────────────────┤
│  opc/  — ZIP I/O, relationships          │
├──────────────────────────────────────────┤
│  common/  — Length, Color                │
└──────────────────────────────────────────┘
```

Each layer depends only on layers below it. No cycles.

| Package | Responsibility |
|---------|----------------|
| `common/` | Shared value types: `Length` (EMU, inches, cm, points, twips), `Color` (RGB). |
| `opc/` | ZIP archive I/O, `[Content_Types].xml`, `.rels` relationship graph, security checks. Format-agnostic — shared across docx/pptx/xlsx. |
| `xmlutil/` | Namespace prefix preservation, `RawXML` for unknown element round-trip, secure marshal/unmarshal with XXE prevention. |
| `docx/wml/` | Go structs mapping to WordprocessingML XML elements. Internal to the library. |
| `docx/` | Public API. Proxy objects wrapping WML structs with thread-safe methods. |

## Round-Trip Fidelity

When you open a `.docx`, modify one paragraph, and save — every other element in the file should be preserved exactly. The library achieves this through:

- **`RawXML` preservation** — unknown or unsupported XML elements are stored as raw bytes during parsing and re-emitted during marshaling. No data is silently dropped.
- **Namespace preservation** — XML namespace prefixes are maintained through a `NamespaceRegistry`.
- **Relationship stability** — existing relationship IDs are preserved; new ones are generated using `NextRelID` to avoid collisions.

## Thread Safety

`Document` uses a `sync.RWMutex`:

- **Read operations** (Text, Markdown, Paragraphs, Comments, etc.) acquire a shared read lock — multiple goroutines can read concurrently.
- **Write operations** (AddParagraph, AddComment, Save, etc.) acquire an exclusive write lock.
- All proxy types (Paragraph, Run, Comment, etc.) acquire locks through their parent Document.

## Extensibility

The architecture supports future formats:

```
go-ooxml/
├── common/     ← shared across all formats
├── xmlutil/    ← shared across all formats
├── opc/        ← shared across all formats
├── docx/       ← Word (.docx)
├── pptx/       ← PowerPoint (.pptx) — future
└── xlsx/       ← Excel (.xlsx) — future
```

The `opc/` layer is format-agnostic. New formats add their own package with format-specific XML struct types and proxy APIs, reusing `common/`, `xmlutil/`, and `opc/`.
