# Tables

Create, populate, and manipulate tables in a document.

## Create a Table

```go
doc, _ := docx.New(nil)
defer doc.Close()

// Create a 3-row, 4-column table.
tbl := doc.Body().AddTable(3, 4)
```

Every cell is initialized with one empty paragraph (required by OOXML).

## Access Cells

```go
cell := tbl.Cell(0, 0) // row 0, column 0
cell.AddParagraph().AddRun("Header")
```

## Iterate Over Rows and Cells

```go
for row := range tbl.Rows() {
	for cell := range row.Cells() {
		fmt.Print(cell.Text() + "\t")
	}
	fmt.Println()
}
```

## Add Rows Dynamically

```go
row := tbl.AddRow()
cell := row.AddCell()
cell.AddParagraph().AddRun("New cell")
```

## Remove a Row

```go
row := tbl.Row(2)
tbl.RemoveRow(row)
```

## Table Style

```go
tbl.SetStyle("TableGrid")
fmt.Println(tbl.Style()) // "TableGrid"
```

## Cell Content with Markdown

```go
cell := tbl.Cell(0, 0)
cell.SetMarkdown("**Bold** header")
```

## Extract Table as Markdown

```go
fmt.Println(tbl.Markdown())
// | H1  | H2  |
// | --- | --- |
// | A   | B   |
```

## Text Extraction

```go
fmt.Println(tbl.Text())
// Rows separated by \n, cells by \t
```

## Table Dimensions

```go
fmt.Println(tbl.RowCount()) // 3
fmt.Println(tbl.ColCount()) // 4
```

## Key Types

- `Table` — Rows, Row, AddRow, RemoveRow, Cell, RowCount, ColCount, Style, SetStyle, Text, Markdown
- `Row` — Cells, AddCell
- `Cell` — Paragraphs, Tables (nested), AddParagraph, Text, Markdown, SetMarkdown
