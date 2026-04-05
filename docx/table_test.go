package docx

import (
	"bytes"
	"strings"
	"testing"
)

func TestTable_AddRow(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	tbl := doc.Body().AddTable(2, 3)
	if tbl == nil {
		t.Fatal("nil table")
	}
	if tbl.RowCount() != 2 {
		t.Errorf("rows = %d", tbl.RowCount())
	}
	if tbl.ColCount() != 3 {
		t.Errorf("cols = %d", tbl.ColCount())
	}
}

func TestTable_Cell(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	tbl := doc.Body().AddTable(2, 2)
	cell := tbl.Cell(0, 0)
	if cell == nil {
		t.Fatal("nil cell")
	}
	cell.AddParagraph().AddRun("test")
	if cell.Text() != "test" {
		t.Errorf("text = %q", cell.Text())
	}
}

func TestTable_Rows_Iterator(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	tbl := doc.Body().AddTable(3, 2)
	count := 0
	for range tbl.Rows() {
		count++
	}
	if count != 3 {
		t.Errorf("count = %d", count)
	}
}

func TestRow_Cells_Iterator(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	tbl := doc.Body().AddTable(1, 4)
	row := tbl.Row(0)
	count := 0
	for range row.Cells() {
		count++
	}
	if count != 4 {
		t.Errorf("count = %d", count)
	}
}

func TestCell_Paragraphs(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	tbl := doc.Body().AddTable(1, 1)
	cell := tbl.Cell(0, 0)
	cell.AddParagraph().AddRun("p1")
	cell.AddParagraph().AddRun("p2")
	count := 0
	for range cell.Paragraphs() {
		count++
	}
	if count < 2 {
		t.Errorf("expected at least 2 paragraphs, got %d", count)
	}
}

func TestTable_Style(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	tbl := doc.Body().AddTable(1, 1)
	tbl.SetStyle("TableGrid")
	if tbl.Style() != "TableGrid" {
		t.Errorf("style = %q", tbl.Style())
	}
}

func TestTable_Text(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	tbl := doc.Body().AddTable(2, 2)
	tbl.Cell(0, 0).AddParagraph().AddRun("a")
	tbl.Cell(0, 1).AddParagraph().AddRun("b")
	tbl.Cell(1, 0).AddParagraph().AddRun("c")
	tbl.Cell(1, 1).AddParagraph().AddRun("d")
	text := tbl.Text()
	for _, s := range []string{"a", "b", "c", "d"} {
		if !containsSubstring(text, s) {
			t.Errorf("text missing %q: %q", s, text)
		}
	}
}

func TestBody_Tables_Iterator(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	doc.Body().AddTable(1, 1)
	doc.Body().AddTable(1, 1)
	count := 0
	for range doc.Body().Tables() {
		count++
	}
	if count != 2 {
		t.Errorf("count = %d", count)
	}
}

func TestTable_RoundTrip(t *testing.T) {
	doc, _ := New(nil)
	tbl := doc.Body().AddTable(1, 2)
	tbl.Cell(0, 0).AddParagraph().AddRun("left")
	tbl.Cell(0, 1).AddParagraph().AddRun("right")
	var buf bytes.Buffer
	doc.Write(&buf)
	doc2, _ := OpenReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()), nil)
	defer doc2.Close()
	count := 0
	for range doc2.Body().Tables() {
		count++
	}
	if count != 1 {
		t.Errorf("tables = %d", count)
	}
}

func TestTable_AddRow_Dynamic(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	tbl := doc.Body().AddTable(1, 2)
	if tbl.RowCount() != 1 {
		t.Fatalf("initial rows = %d", tbl.RowCount())
	}
	row := tbl.AddRow()
	if row == nil {
		t.Fatal("nil row from AddRow")
	}
	if tbl.RowCount() != 2 {
		t.Errorf("rows after AddRow = %d", tbl.RowCount())
	}
}

func TestTable_RemoveRow(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	tbl := doc.Body().AddTable(3, 1)
	row := tbl.Row(1)
	if err := tbl.RemoveRow(row); err != nil {
		t.Fatalf("RemoveRow: %v", err)
	}
	if tbl.RowCount() != 2 {
		t.Errorf("rows after remove = %d", tbl.RowCount())
	}
}

func TestTable_Cell_OutOfBounds(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	tbl := doc.Body().AddTable(2, 2)
	if tbl.Cell(-1, 0) != nil {
		t.Error("expected nil for negative row")
	}
	if tbl.Cell(0, 5) != nil {
		t.Error("expected nil for out-of-bounds col")
	}
	if tbl.Cell(5, 0) != nil {
		t.Error("expected nil for out-of-bounds row")
	}
}

// helper used in TestTable_Text
func containsSubstring(s, sub string) bool {
	return strings.Contains(s, sub)
}

func ExampleBody_AddTable() {
	doc, _ := New(nil)
	defer doc.Close()
	tbl := doc.Body().AddTable(2, 3)
	tbl.SetStyle("TableGrid")
	tbl.Cell(0, 0).AddParagraph().AddRun("Header 1")
	tbl.Cell(0, 1).AddParagraph().AddRun("Header 2")
	tbl.Cell(0, 2).AddParagraph().AddRun("Header 3")
	tbl.Cell(1, 0).AddParagraph().AddRun("Value A")
	tbl.Cell(1, 1).AddParagraph().AddRun("Value B")
	tbl.Cell(1, 2).AddParagraph().AddRun("Value C")
	// doc.WriteTo(w)
}

func ExampleTable_Rows() {
	doc, _ := New(nil)
	defer doc.Close()
	tbl := doc.Body().AddTable(2, 2)
	tbl.Cell(0, 0).AddParagraph().AddRun("A")
	tbl.Cell(0, 1).AddParagraph().AddRun("B")
	for row := range tbl.Rows() {
		for cell := range row.Cells() {
			_ = cell.Text()
		}
	}
}

func ExampleCell_Text() {
	doc, _ := New(nil)
	defer doc.Close()
	tbl := doc.Body().AddTable(1, 1)
	tbl.Cell(0, 0).AddParagraph().AddRun("hello")
	_ = tbl.Cell(0, 0).Text() // "hello"
}

func ExampleCell_AddParagraph() {
	doc, _ := New(nil)
	defer doc.Close()
	tbl := doc.Body().AddTable(1, 1)
	c := tbl.Cell(0, 0)
	c.AddParagraph().AddRun("first paragraph")
	c.AddParagraph().AddRun("second paragraph")
}

func ExampleRow_AddCell() {
	doc, _ := New(nil)
	defer doc.Close()
	tbl := doc.Body().AddTable(1, 1)
	row := tbl.Row(0)
	cell := row.AddCell()
	cell.AddParagraph().AddRun("new cell")
}

func ExampleTable_Markdown() {
	doc, _ := New(nil)
	defer doc.Close()
	tbl := doc.Body().AddTable(2, 2)
	tbl.Cell(0, 0).AddParagraph().AddRun("H1")
	tbl.Cell(0, 1).AddParagraph().AddRun("H2")
	tbl.Cell(1, 0).AddParagraph().AddRun("A")
	tbl.Cell(1, 1).AddParagraph().AddRun("B")
	_ = tbl.Markdown() // pipe table
}

func ExampleCell_Markdown() {
	doc, _ := New(nil)
	defer doc.Close()
	tbl := doc.Body().AddTable(1, 1)
	tbl.Cell(0, 0).AddParagraph().AddRun("content")
	_ = tbl.Cell(0, 0).Markdown() // "content"
}
