package docx

import (
	"encoding/xml"
	"iter"
	"strings"

	"github.com/ieshan/go-ooxml/docx/wml"
)

// Table wraps a wml.CT_Tbl element.
type Table struct {
	doc *Document
	el  *wml.CT_Tbl
}

// Rows returns an iterator over all rows in the table.
func (t *Table) Rows() iter.Seq[*Row] {
	return func(yield func(*Row) bool) {
		t.doc.mu.RLock()
		defer t.doc.mu.RUnlock()
		for _, r := range t.el.Rows {
			if !yield(&Row{doc: t.doc, el: r}) {
				return
			}
		}
	}
}

// Row returns the row at the given index, or nil if out of bounds.
func (t *Table) Row(index int) *Row {
	t.doc.mu.RLock()
	defer t.doc.mu.RUnlock()
	if index < 0 || index >= len(t.el.Rows) {
		return nil
	}
	return &Row{doc: t.doc, el: t.el.Rows[index]}
}

// AddRow appends a new empty row to the table and returns it.
func (t *Table) AddRow() *Row {
	t.doc.mu.Lock()
	defer t.doc.mu.Unlock()
	r := &wml.CT_Row{XMLName: xml.Name{Space: wml.Ns, Local: "tr"}}
	t.el.Rows = append(t.el.Rows, r)
	return &Row{doc: t.doc, el: r}
}

// RemoveRow removes the given row from the table.
// Returns an error if the row is not found.
func (t *Table) RemoveRow(row *Row) error {
	t.doc.mu.Lock()
	defer t.doc.mu.Unlock()
	for i, r := range t.el.Rows {
		if r == row.el {
			t.el.Rows = append(t.el.Rows[:i], t.el.Rows[i+1:]...)
			return nil
		}
	}
	return ErrNotFound
}

// Cell returns the cell at (row, col), or nil if indices are out of bounds.
func (t *Table) Cell(row, col int) *Cell {
	t.doc.mu.RLock()
	defer t.doc.mu.RUnlock()
	if row < 0 || row >= len(t.el.Rows) {
		return nil
	}
	r := t.el.Rows[row]
	if col < 0 || col >= len(r.Cells) {
		return nil
	}
	return &Cell{doc: t.doc, el: r.Cells[col]}
}

// RowCount returns the number of rows in the table.
func (t *Table) RowCount() int {
	t.doc.mu.RLock()
	defer t.doc.mu.RUnlock()
	return len(t.el.Rows)
}

// ColCount returns the number of columns based on the first row's cell count.
func (t *Table) ColCount() int {
	t.doc.mu.RLock()
	defer t.doc.mu.RUnlock()
	if len(t.el.Rows) == 0 {
		return 0
	}
	return len(t.el.Rows[0].Cells)
}

// Style returns the table style name, or empty string if none is set.
func (t *Table) Style() string {
	t.doc.mu.RLock()
	defer t.doc.mu.RUnlock()
	if t.el.TblPr != nil && t.el.TblPr.Style != nil {
		return *t.el.TblPr.Style
	}
	return ""
}

// SetStyle sets the table style name.
func (t *Table) SetStyle(name string) {
	t.doc.mu.Lock()
	defer t.doc.mu.Unlock()
	if t.el.TblPr == nil {
		t.el.TblPr = &wml.CT_TblPr{}
	}
	t.el.TblPr.Style = &name
}

// Text returns a concatenation of all cell texts, rows separated by newlines,
// cells within a row separated by tabs.
func (t *Table) Text() string {
	t.doc.mu.RLock()
	defer t.doc.mu.RUnlock()
	var sb strings.Builder
	for i, r := range t.el.Rows {
		if i > 0 {
			sb.WriteByte('\n')
		}
		for j, tc := range r.Cells {
			if j > 0 {
				sb.WriteByte('\t')
			}
			sb.WriteString(cellText(tc))
		}
	}
	return sb.String()
}

// Markdown returns the table formatted as a Markdown pipe table, with the
// first row as the header and a --- separator line.
func (t *Table) Markdown() string {
	t.doc.mu.RLock()
	defer t.doc.mu.RUnlock()
	return tableToMarkdown(t.el, t.doc.hyperlinkURLMap())
}

// cellText returns the text content of a CT_Tc without acquiring the lock
// (caller must already hold at least RLock).
func cellText(tc *wml.CT_Tc) string {
	var sb strings.Builder
	for _, c := range tc.Content {
		if c.Paragraph != nil {
			sb.WriteString(c.Paragraph.Text())
		} else if c.Table != nil {
			// Nested table text inline
			for _, r := range c.Table.Rows {
				for _, nested := range r.Cells {
					sb.WriteString(cellText(nested))
				}
			}
		}
	}
	return sb.String()
}

// Row wraps a wml.CT_Row element.
type Row struct {
	doc *Document
	el  *wml.CT_Row
}

// Cells returns an iterator over all cells in the row.
func (r *Row) Cells() iter.Seq[*Cell] {
	return func(yield func(*Cell) bool) {
		r.doc.mu.RLock()
		defer r.doc.mu.RUnlock()
		for _, tc := range r.el.Cells {
			if !yield(&Cell{doc: r.doc, el: tc}) {
				return
			}
		}
	}
}

// AddCell appends a new empty cell to the row and returns it.
func (r *Row) AddCell() *Cell {
	r.doc.mu.Lock()
	defer r.doc.mu.Unlock()
	tc := &wml.CT_Tc{XMLName: xml.Name{Space: wml.Ns, Local: "tc"}}
	// A cell must contain at least one paragraph to be valid OOXML.
	p := &wml.CT_P{XMLName: xml.Name{Space: wml.Ns, Local: "p"}}
	tc.Content = append(tc.Content, wml.BlockLevelContent{Paragraph: p})
	r.el.Cells = append(r.el.Cells, tc)
	return &Cell{doc: r.doc, el: tc}
}

// Cell wraps a wml.CT_Tc element (block container).
type Cell struct {
	doc *Document
	el  *wml.CT_Tc
}

// Paragraphs returns an iterator over all top-level paragraphs in the cell.
func (c *Cell) Paragraphs() iter.Seq[*Paragraph] {
	return func(yield func(*Paragraph) bool) {
		c.doc.mu.RLock()
		defer c.doc.mu.RUnlock()
		for _, blk := range c.el.Content {
			if blk.Paragraph != nil {
				if !yield(&Paragraph{doc: c.doc, el: blk.Paragraph}) {
					return
				}
			}
		}
	}
}

// Tables returns an iterator over nested tables in the cell.
func (c *Cell) Tables() iter.Seq[*Table] {
	return func(yield func(*Table) bool) {
		c.doc.mu.RLock()
		defer c.doc.mu.RUnlock()
		for _, blk := range c.el.Content {
			if blk.Table != nil {
				if !yield(&Table{doc: c.doc, el: blk.Table}) {
					return
				}
			}
		}
	}
}

// AddParagraph appends a new empty paragraph to the cell and returns it.
func (c *Cell) AddParagraph() *Paragraph {
	c.doc.mu.Lock()
	defer c.doc.mu.Unlock()
	p := &wml.CT_P{XMLName: xml.Name{Space: wml.Ns, Local: "p"}}
	c.el.Content = append(c.el.Content, wml.BlockLevelContent{Paragraph: p})
	return &Paragraph{doc: c.doc, el: p}
}

// Text returns the concatenated text of all paragraphs in the cell.
func (c *Cell) Text() string {
	c.doc.mu.RLock()
	defer c.doc.mu.RUnlock()
	return cellText(c.el)
}

// Markdown returns the cell content formatted as Markdown.
func (c *Cell) Markdown() string {
	c.doc.mu.RLock()
	defer c.doc.mu.RUnlock()
	return cellToMarkdown(c.el, c.doc.hyperlinkURLMap())
}
