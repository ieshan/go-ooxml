package docx

import (
	"encoding/json"
	"encoding/xml"
	"fmt"
	"strconv"

	"github.com/ieshan/go-ooxml/docx/wml"
)

// FromProseMirror creates a new Document from ProseMirror JSON. The data should
// be a JSON-encoded PMNode of type "doc". Both TipTap and vanilla ProseMirror
// naming conventions are accepted.
func FromProseMirror(data []byte, cfg *Config) (*Document, error) {
	doc, err := New(cfg)
	if err != nil {
		return nil, err
	}
	if err := doc.ImportProseMirror(data); err != nil {
		_ = doc.Close()
		return nil, err
	}
	return doc, nil
}

// ImportProseMirror appends ProseMirror JSON content to the document body. The
// data should be a JSON-encoded PMNode. If the root type is "doc", its children
// are appended. Otherwise, the single node is imported.
func (d *Document) ImportProseMirror(data []byte) error {
	var root PMNode
	if err := json.Unmarshal(data, &root); err != nil {
		return fmt.Errorf("docx: parse ProseMirror JSON: %w", err)
	}

	d.mu.Lock()
	defer d.mu.Unlock()

	root.Type = normalizeTypeName(root.Type)

	if root.Type == "doc" {
		// Apply doc-level attrs (section properties, document properties).
		if root.Attrs != nil {
			d.applyDocAttrs(root.Attrs)
		}
		for _, child := range root.Content {
			d.importPMNodeLocked(child)
		}
	} else {
		d.importPMNodeLocked(root)
	}
	return nil
}

// SetProseMirror replaces the paragraph content with ProseMirror JSON inline content.
// The data should be a JSON-encoded PMNode of type "paragraph" or similar.
func (p *Paragraph) SetProseMirror(data []byte) error {
	var node PMNode
	if err := json.Unmarshal(data, &node); err != nil {
		return err
	}

	p.doc.mu.Lock()
	defer p.doc.mu.Unlock()

	p.el.Content = nil
	node.Type = normalizeTypeName(node.Type)
	for _, child := range node.Content {
		ic := p.doc.pmNodeToInlineContent(child)
		p.el.Content = append(p.el.Content, ic...)
	}
	return nil
}

// SetProseMirror replaces the cell content with ProseMirror JSON block content.
func (c *Cell) SetProseMirror(data []byte) error {
	var node PMNode
	if err := json.Unmarshal(data, &node); err != nil {
		return err
	}

	c.doc.mu.Lock()
	defer c.doc.mu.Unlock()

	c.el.Content = nil
	node.Type = normalizeTypeName(node.Type)
	for _, child := range node.Content {
		child.Type = normalizeTypeName(child.Type)
		switch child.Type {
		case "paragraph", "heading":
			p := importPMParagraph(c.doc, child)
			c.el.Content = append(c.el.Content, wml.BlockLevelContent{Paragraph: p})
		default:
			// Ensure cell has at least one paragraph.
		}
	}
	// OOXML requires at least one paragraph in a cell.
	if len(c.el.Content) == 0 {
		p := &wml.CT_P{XMLName: xml.Name{Space: wml.Ns, Local: "p"}}
		c.el.Content = append(c.el.Content, wml.BlockLevelContent{Paragraph: p})
	}
	return nil
}

// ---------------------------------------------------------------------------
// Internal import helpers
// ---------------------------------------------------------------------------

// importPMNodeLocked dispatches a PM node to the appropriate importer.
// Caller must hold the write lock.
func (d *Document) importPMNodeLocked(node PMNode) {
	node.Type = normalizeTypeName(node.Type)
	body := d.doc.Body

	switch node.Type {
	case "paragraph", "heading":
		p := importPMParagraph(d, node)
		body.Content = append(body.Content, wml.BlockLevelContent{Paragraph: p})

	case "blockquote":
		for _, child := range node.Content {
			child.Type = normalizeTypeName(child.Type)
			p := importPMParagraph(d, child)
			if p.PPr == nil {
				p.PPr = &wml.CT_PPr{XMLName: xml.Name{Space: wml.Ns, Local: "pPr"}}
			}
			p.PPr.Style = new("Quote")
			body.Content = append(body.Content, wml.BlockLevelContent{Paragraph: p})
		}

	case "codeBlock":
		for _, child := range node.Content {
			if child.Type == "text" && child.Text != "" {
				p := &wml.CT_P{XMLName: xml.Name{Space: wml.Ns, Local: "p"}}
				r := &wml.CT_R{XMLName: xml.Name{Space: wml.Ns, Local: "r"}}
				r.AddText(child.Text)
				r.RPr = &wml.CT_RPr{FontName: new("Courier New")}
				p.Content = append(p.Content, wml.InlineContent{Run: r})
				body.Content = append(body.Content, wml.BlockLevelContent{Paragraph: p})
			}
		}

	case "bulletList", "orderedList":
		importPMList(d, node, false)

	case "taskList":
		importPMList(d, node, true)

	case "table":
		tbl := importPMTable(d, node)
		body.Content = append(body.Content, wml.BlockLevelContent{Table: tbl})

	case "horizontalRule":
		// Empty separator paragraph.
		p := &wml.CT_P{XMLName: xml.Name{Space: wml.Ns, Local: "p"}}
		body.Content = append(body.Content, wml.BlockLevelContent{Paragraph: p})

	case "pageBreak":
		// Page break as a paragraph containing a page break run.
		p := &wml.CT_P{XMLName: xml.Name{Space: wml.Ns, Local: "p"}}
		r := &wml.CT_R{XMLName: xml.Name{Space: wml.Ns, Local: "r"}}
		r.Content = append(r.Content, wml.RunContent{Break: &wml.CT_Break{Type: "page"}})
		p.Content = append(p.Content, wml.InlineContent{Run: r})
		body.Content = append(body.Content, wml.BlockLevelContent{Paragraph: p})

	case "hardBreak":
		// Standalone hard break — unlikely at block level, but handle gracefully.
		p := &wml.CT_P{XMLName: xml.Name{Space: wml.Ns, Local: "p"}}
		r := &wml.CT_R{XMLName: xml.Name{Space: wml.Ns, Local: "r"}}
		r.Content = append(r.Content, wml.RunContent{Break: &wml.CT_Break{}})
		p.Content = append(p.Content, wml.InlineContent{Run: r})
		body.Content = append(body.Content, wml.BlockLevelContent{Paragraph: p})
	}
}

// importPMParagraph creates a CT_P from a paragraph or heading PM node.
func importPMParagraph(d *Document, node PMNode) *wml.CT_P {
	node.Type = normalizeTypeName(node.Type)
	p := &wml.CT_P{XMLName: xml.Name{Space: wml.Ns, Local: "p"}}

	if node.Type == "heading" {
		level := 1
		if node.Attrs != nil {
			if v, ok := node.Attrs["level"]; ok {
				if f, ok := v.(float64); ok {
					level = int(f)
				}
			}
		}
		if level < 1 {
			level = 1
		}
		if level > 6 {
			level = 6
		}
		p.PPr = &wml.CT_PPr{
			XMLName: xml.Name{Space: wml.Ns, Local: "pPr"},
			Style:   new("Heading" + strconv.Itoa(level)),
		}
	}

	// Apply paragraph attrs.
	if node.Attrs != nil {
		if align, ok := node.Attrs["textAlign"]; ok {
			if s, ok := align.(string); ok && s != "" {
				if s == "justify" {
					s = "both"
				}
				if p.PPr == nil {
					p.PPr = &wml.CT_PPr{XMLName: xml.Name{Space: wml.Ns, Local: "pPr"}}
				}
				p.PPr.Alignment = &s
			}
		}
		if sid, ok := node.Attrs["styleId"]; ok {
			if s, ok := sid.(string); ok && s != "" {
				if p.PPr == nil {
					p.PPr = &wml.CT_PPr{XMLName: xml.Name{Space: wml.Ns, Local: "pPr"}}
				}
				if p.PPr.Style == nil {
					p.PPr.Style = &s
				}
			}
		}
	}

	// Import inline content.
	for _, child := range node.Content {
		ics := d.pmNodeToInlineContent(child)
		p.Content = append(p.Content, ics...)
	}

	return p
}

// pmNodeToInlineContent converts a PM inline node to WML InlineContent entries.
// Caller must hold write lock.
func (d *Document) pmNodeToInlineContent(node PMNode) []wml.InlineContent {
	node.Type = normalizeTypeName(node.Type)

	switch node.Type {
	case "text":
		// Check if this text has a link mark — if so, create a hyperlink.
		var linkHref string
		for _, m := range node.Marks {
			mt := normalizeTypeName(m.Type)
			if mt == "link" && m.Attrs != nil {
				if href, ok := m.Attrs["href"].(string); ok {
					linkHref = href
				}
			}
		}

		if linkHref != "" {
			// Create hyperlink with relationship.
			relID := d.addHyperlinkRelationship(linkHref)
			r := &wml.CT_R{XMLName: xml.Name{Space: wml.Ns, Local: "r"}}
			r.AddText(node.Text)
			applyPMMarksToRun(r, node.Marks)
			h := &wml.CT_Hyperlink{
				XMLName: xml.Name{Space: wml.Ns, Local: "hyperlink"},
				ID:      relID,
				Runs:    []*wml.CT_R{r},
			}
			return []wml.InlineContent{{Hyperlink: h}}
		}

		r := &wml.CT_R{XMLName: xml.Name{Space: wml.Ns, Local: "r"}}
		r.AddText(node.Text)
		applyPMMarksToRun(r, node.Marks)
		return []wml.InlineContent{{Run: r}}

	case "hardBreak":
		r := &wml.CT_R{XMLName: xml.Name{Space: wml.Ns, Local: "r"}}
		r.Content = append(r.Content, wml.RunContent{Break: &wml.CT_Break{}})
		return []wml.InlineContent{{Run: r}}

	case "image":
		// Images are not fully implemented; create a placeholder text run.
		alt := ""
		if node.Attrs != nil {
			if v, ok := node.Attrs["alt"].(string); ok {
				alt = v
			}
		}
		if alt == "" {
			alt = "[image]"
		}
		r := &wml.CT_R{XMLName: xml.Name{Space: wml.Ns, Local: "r"}}
		r.AddText(alt)
		return []wml.InlineContent{{Run: r}}

	default:
		return nil
	}
}

// applyPMMarksToRun applies PM marks to a CT_R's run properties.
func applyPMMarksToRun(r *wml.CT_R, marks []PMMark) {
	if len(marks) == 0 {
		return
	}

	ensureRPr := func() *wml.CT_RPr {
		if r.RPr == nil {
			r.RPr = &wml.CT_RPr{}
		}
		return r.RPr
	}

	for _, m := range marks {
		mt := normalizeTypeName(m.Type)
		switch mt {
		case "bold":
			ensureRPr().Bold = new(true)
		case "italic":
			ensureRPr().Italic = new(true)
		case "underline":
			ensureRPr().Underline = new("single")
		case "strike":
			ensureRPr().Strike = new(true)
		case "code":
			ensureRPr().FontName = new("Courier New")
		case "superscript":
			// Stored in Extra for now (vertAlign not in CT_RPr fields).
		case "subscript":
			// Stored in Extra for now.
		case "textStyle":
			if m.Attrs == nil {
				continue
			}
			rpr := ensureRPr()
			if color, ok := m.Attrs["color"].(string); ok && color != "" {
				rpr.Color = new(trimColorHash(color))
			}
			if fs, ok := m.Attrs["fontSize"].(string); ok && fs != "" {
				rpr.FontSize = new(pmPtStringToHalfPoints(fs))
			}
			if ff, ok := m.Attrs["fontFamily"].(string); ok && ff != "" {
				rpr.FontName = &ff
			}
		case "highlight":
			// Named highlight color — store in Extra for now.
		case "link":
			// Handled at the pmNodeToInlineContent level.
		}
	}
}

// importPMTable creates a CT_Tbl from a table PM node.
func importPMTable(d *Document, node PMNode) *wml.CT_Tbl {
	tbl := &wml.CT_Tbl{XMLName: xml.Name{Space: wml.Ns, Local: "tbl"}}

	// Table style from attrs.
	if node.Attrs != nil {
		if style, ok := node.Attrs["tableStyle"].(string); ok && style != "" {
			tbl.TblPr = &wml.CT_TblPr{Style: &style}
		}
	}

	for _, rowNode := range node.Content {
		rowNode.Type = normalizeTypeName(rowNode.Type)
		if rowNode.Type != "tableRow" {
			continue
		}
		row := &wml.CT_Row{XMLName: xml.Name{Space: wml.Ns, Local: "tr"}}
		for _, cellNode := range rowNode.Content {
			cellNode.Type = normalizeTypeName(cellNode.Type)
			if cellNode.Type != "tableCell" && cellNode.Type != "tableHeader" {
				continue
			}
			tc := &wml.CT_Tc{XMLName: xml.Name{Space: wml.Ns, Local: "tc"}}

			// Cell attrs.
			if cellNode.Attrs != nil {
				if cs, ok := cellNode.Attrs["colspan"]; ok {
					if v := pmAttrToInt(cs); v > 1 {
						tc.TcPr = ensureTcPr(tc)
						tc.TcPr.GridSpan = &v
					}
				}
				if bg, ok := cellNode.Attrs["background"].(string); ok && bg != "" {
					tc.TcPr = ensureTcPr(tc)
					fill := trimColorHash(bg)
					tc.TcPr.Shading = &wml.CT_Shd{Val: "clear", Fill: fill}
				}
			}

			// Cell content (paragraphs).
			for _, child := range cellNode.Content {
				child.Type = normalizeTypeName(child.Type)
				if child.Type == "paragraph" || child.Type == "heading" {
					p := importPMParagraph(d, child)
					tc.Content = append(tc.Content, wml.BlockLevelContent{Paragraph: p})
				}
			}
			// OOXML requires at least one paragraph.
			if len(tc.Content) == 0 {
				p := &wml.CT_P{XMLName: xml.Name{Space: wml.Ns, Local: "p"}}
				tc.Content = append(tc.Content, wml.BlockLevelContent{Paragraph: p})
			}

			row.Cells = append(row.Cells, tc)
		}
		tbl.Rows = append(tbl.Rows, row)
	}

	// Build grid.
	if len(tbl.Rows) > 0 {
		cols := len(tbl.Rows[0].Cells)
		tbl.TblGrid = &wml.CT_TblGrid{}
		for i := 0; i < cols; i++ {
			tbl.TblGrid.Cols = append(tbl.TblGrid.Cols, &wml.CT_TblGridCol{W: "0"})
		}
	}

	return tbl
}

// importPMList converts a bulletList/orderedList/taskList PM node into
// paragraphs with list styles (OOXML lists are flat, not nested).
func importPMList(d *Document, node PMNode, isTask bool) {
	node.Type = normalizeTypeName(node.Type)
	isBullet := node.Type == "bulletList" || node.Type == "taskList"

	style := "ListBullet"
	if !isBullet {
		style = "ListNumber"
	}

	for _, item := range node.Content {
		item.Type = normalizeTypeName(item.Type)
		if item.Type != "listItem" && item.Type != "taskItem" {
			continue
		}
		for _, child := range item.Content {
			child.Type = normalizeTypeName(child.Type)
			switch child.Type {
			case "paragraph", "heading":
				p := importPMParagraph(d, child)
				if p.PPr == nil {
					p.PPr = &wml.CT_PPr{XMLName: xml.Name{Space: wml.Ns, Local: "pPr"}}
				}
				if p.PPr.Style == nil {
					p.PPr.Style = &style
				}
				d.doc.Body.Content = append(d.doc.Body.Content, wml.BlockLevelContent{Paragraph: p})
			case "bulletList", "orderedList":
				// Nested list — recurse.
				importPMList(d, child, false)
			case "taskList":
				importPMList(d, child, true)
			}
		}
	}
}

// applyDocAttrs applies doc-level PM attrs to the Document (section properties, properties).
// Caller must hold write lock.
func (d *Document) applyDocAttrs(attrs map[string]any) {
	// Document properties — access directly without re-acquiring lock.
	needProps := false
	if _, ok := attrs["title"].(string); ok {
		needProps = true
	}
	if _, ok := attrs["author"].(string); ok {
		needProps = true
	}
	if _, ok := attrs["description"].(string); ok {
		needProps = true
	}

	if needProps {
		if d.props == nil {
			d.props = &Properties{doc: d}
		}
		if title, ok := attrs["title"].(string); ok && title != "" {
			d.props.cp.Title = title
		}
		if author, ok := attrs["author"].(string); ok && author != "" {
			d.props.cp.Creator = author
		}
		if desc, ok := attrs["description"].(string); ok && desc != "" {
			d.props.cp.Description = desc
		}
	}
}

// ensureTcPr ensures the cell has a TcPr element.
func ensureTcPr(tc *wml.CT_Tc) *wml.CT_TcPr {
	if tc.TcPr == nil {
		tc.TcPr = &wml.CT_TcPr{}
	}
	return tc.TcPr
}

// pmAttrToInt converts a PM attr value (float64 from JSON) to int.
func pmAttrToInt(v any) int {
	switch val := v.(type) {
	case float64:
		return int(val)
	case int:
		return val
	case string:
		n, _ := strconv.Atoi(val)
		return n
	}
	return 0
}
