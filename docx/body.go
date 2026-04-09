package docx

import (
	"encoding/xml"
	"iter"
	"strings"

	"github.com/ieshan/go-ooxml/docx/wml"
)

// Body represents the document body -- the container for all block-level content.
type Body struct {
	doc *Document
	el  *wml.CT_Body
}

// Paragraphs returns an iterator over all paragraphs in the body.
func (b *Body) Paragraphs() iter.Seq[*Paragraph] {
	return func(yield func(*Paragraph) bool) {
		b.doc.mu.RLock()
		defer b.doc.mu.RUnlock()
		for _, c := range b.el.Content {
			if c.Paragraph != nil {
				if !yield(&Paragraph{doc: b.doc, el: c.Paragraph}) {
					return
				}
			}
		}
	}
}

// Tables returns an iterator over all tables in the body.
func (b *Body) Tables() iter.Seq[*Table] {
	return func(yield func(*Table) bool) {
		b.doc.mu.RLock()
		defer b.doc.mu.RUnlock()
		for _, c := range b.el.Content {
			if c.Table != nil {
				if !yield(&Table{doc: b.doc, el: c.Table}) {
					return
				}
			}
		}
	}
}

// AddParagraph appends a new empty paragraph to the body and returns it.
func (b *Body) AddParagraph() *Paragraph {
	b.doc.mu.Lock()
	defer b.doc.mu.Unlock()

	p := &wml.CT_P{XMLName: xml.Name{Space: wml.Ns, Local: "p"}}
	b.el.Content = append(b.el.Content, wml.BlockLevelContent{Paragraph: p})
	return &Paragraph{doc: b.doc, el: p}
}

// AddTable appends a new table with the given number of rows and columns to the
// body. Each cell is initialised with one empty paragraph as required by OOXML.
func (b *Body) AddTable(rows, cols int) *Table {
	b.doc.mu.Lock()
	defer b.doc.mu.Unlock()

	tbl := &wml.CT_Tbl{XMLName: xml.Name{Space: wml.Ns, Local: "tbl"}}
	// Build grid columns.
	tbl.TblGrid = &wml.CT_TblGrid{}
	for i := 0; i < cols; i++ {
		tbl.TblGrid.Cols = append(tbl.TblGrid.Cols, &wml.CT_TblGridCol{W: "0"})
	}
	// Build rows and cells.
	for i := 0; i < rows; i++ {
		row := &wml.CT_Row{XMLName: xml.Name{Space: wml.Ns, Local: "tr"}}
		for j := 0; j < cols; j++ {
			tc := &wml.CT_Tc{XMLName: xml.Name{Space: wml.Ns, Local: "tc"}}
			// Each cell must contain at least one paragraph.
			p := &wml.CT_P{XMLName: xml.Name{Space: wml.Ns, Local: "p"}}
			tc.Content = append(tc.Content, wml.BlockLevelContent{Paragraph: p})
			row.Cells = append(row.Cells, tc)
		}
		tbl.Rows = append(tbl.Rows, row)
	}
	b.el.Content = append(b.el.Content, wml.BlockLevelContent{Table: tbl})
	return &Table{doc: b.doc, el: tbl}
}

// Clear removes all block-level content (paragraphs, tables, and unknown
// elements) from the body. Section properties (page size, margins, orientation)
// are preserved.
func (b *Body) Clear() {
	b.doc.mu.Lock()
	defer b.doc.mu.Unlock()
	b.el.Content = nil
}

// Text concatenates the text of all block-level content (paragraphs and tables),
// separated by newlines. Table rows are newline-separated; cells within a row are tab-separated.
func (b *Body) Text() string {
	b.doc.mu.RLock()
	defer b.doc.mu.RUnlock()

	return bodyText(b.el)
}

// bodyText returns the text of a CT_Body without acquiring any lock.
// Caller must hold at least RLock.
func bodyText(body *wml.CT_Body) string {
	var sb strings.Builder
	first := true
	for _, c := range body.Content {
		var chunk string
		switch {
		case c.Paragraph != nil:
			chunk = c.Paragraph.Text()
		case c.Table != nil:
			// Table text: rows separated by \n, cells separated by \t.
			var tsb strings.Builder
			firstRow := true
			for _, r := range c.Table.Rows {
				if !firstRow {
					tsb.WriteByte('\n')
				}
				for j, tc := range r.Cells {
					if j > 0 {
						tsb.WriteByte('\t')
					}
					tsb.WriteString(cellText(tc))
				}
				firstRow = false
			}
			chunk = tsb.String()
		default:
			continue
		}
		if !first {
			sb.WriteByte('\n')
		}
		sb.WriteString(chunk)
		first = false
	}
	return sb.String()
}

// Markdown returns the body content formatted as Markdown, with paragraphs
// separated by blank lines and tables rendered as pipe tables.
func (b *Body) Markdown() string {
	b.doc.mu.RLock()
	defer b.doc.mu.RUnlock()
	return bodyToMarkdown(b.el, b.doc.hyperlinkURLMap())
}
