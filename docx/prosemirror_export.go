package docx

import (
	"encoding/json"
	"strconv"
	"strings"

	"github.com/ieshan/go-ooxml/docx/wml"
)

// ProseMirror renders the full document as ProseMirror JSON. The returned bytes
// are a JSON-encoded PMNode of type "doc". Section properties (page size,
// margins, orientation, columns) and document properties (title, author) are
// included as attrs on the doc node.
func (d *Document) ProseMirror(opts *ProseMirrorOptions) ([]byte, error) {
	d.mu.RLock()
	defer d.mu.RUnlock()

	o := defaultPMOptions(opts)
	doc := PMNode{
		Type:  "doc",
		Attrs: map[string]any{},
	}

	// Section properties.
	if d.sectPr != nil {
		addSectionAttrs(doc.Attrs, d.sectPr)
	}

	// Document properties — check props and also look at Config author.
	if d.props != nil {
		if d.props.cp.Title != "" {
			doc.Attrs["title"] = d.props.cp.Title
		}
		if d.props.cp.Creator != "" {
			doc.Attrs["author"] = d.props.cp.Creator
		}
		if d.props.cp.Description != "" {
			doc.Attrs["description"] = d.props.cp.Description
		}
	}
	// Fallback: use Config author if props haven't been loaded.
	if doc.Attrs["author"] == nil && d.cfg.Author != "" {
		doc.Attrs["author"] = d.cfg.Author
	}

	// Headers/footers.
	if o.IncludeHeaders {
		addHdrFtrAttrs(d, &doc, o)
	}

	if len(doc.Attrs) == 0 {
		doc.Attrs = nil
	}

	// Body content.
	urls := d.hyperlinkURLMap()
	doc.Content = bodyToPMNodes(d.doc.Body, urls, o)

	return json.Marshal(doc)
}

// Body.ProseMirror returns the body content as ProseMirror JSON.
func (b *Body) ProseMirror() ([]byte, error) {
	b.doc.mu.RLock()
	defer b.doc.mu.RUnlock()
	urls := b.doc.hyperlinkURLMap()
	o := &ProseMirrorOptions{UseTipTapNames: true}
	doc := PMNode{
		Type:    "doc",
		Content: bodyToPMNodes(b.el, urls, o),
	}
	return json.Marshal(doc)
}

// Paragraph.ProseMirror returns the paragraph as ProseMirror JSON.
func (p *Paragraph) ProseMirror() ([]byte, error) {
	p.doc.mu.RLock()
	defer p.doc.mu.RUnlock()
	urls := p.doc.hyperlinkURLMap()
	o := &ProseMirrorOptions{UseTipTapNames: true}
	node := paragraphToPMNode(p.el, urls, o)
	return json.Marshal(node)
}

// Run.ProseMirror returns the run as ProseMirror JSON (a text node with marks).
func (r *Run) ProseMirror() ([]byte, error) {
	r.doc.mu.RLock()
	defer r.doc.mu.RUnlock()
	node := runToPMNode(r.el, true)
	return json.Marshal(node)
}

// Table.ProseMirror returns the table as ProseMirror JSON.
func (t *Table) ProseMirror() ([]byte, error) {
	t.doc.mu.RLock()
	defer t.doc.mu.RUnlock()
	urls := t.doc.hyperlinkURLMap()
	o := &ProseMirrorOptions{UseTipTapNames: true}
	node := tableToPMNode(t.el, urls, o)
	return json.Marshal(node)
}

// Cell.ProseMirror returns the table cell as ProseMirror JSON.
func (c *Cell) ProseMirror() ([]byte, error) {
	c.doc.mu.RLock()
	defer c.doc.mu.RUnlock()
	urls := c.doc.hyperlinkURLMap()
	o := &ProseMirrorOptions{UseTipTapNames: true}
	node := cellToPMNode(c.el, urls, o, false)
	return json.Marshal(node)
}

// Comment.ProseMirror returns the comment body as ProseMirror JSON.
func (c *Comment) ProseMirror() ([]byte, error) {
	c.doc.mu.RLock()
	defer c.doc.mu.RUnlock()
	o := &ProseMirrorOptions{UseTipTapNames: true}
	var content []PMNode
	for _, bc := range c.el.Content {
		if bc.Paragraph != nil {
			content = append(content, paragraphToPMNode(bc.Paragraph, nil, o))
		}
	}
	doc := PMNode{Type: "doc", Content: content}
	return json.Marshal(doc)
}

// ProposedProseMirror returns the proposed (inserted) text of a revision as ProseMirror JSON.
func (r *Revision) ProposedProseMirror() ([]byte, error) {
	r.doc.mu.RLock()
	defer r.doc.mu.RUnlock()
	if r.typ != RevisionInsert {
		return json.Marshal(PMNode{Type: "doc"})
	}
	var content []PMNode
	for _, run := range r.el.Runs {
		n := runToPMNode(run, true)
		if n.Text != "" {
			content = append(content, n)
		}
	}
	p := PMNode{Type: "paragraph", Content: content}
	return json.Marshal(PMNode{Type: "doc", Content: []PMNode{p}})
}

// OriginalProseMirror returns the original (deleted) text of a revision as ProseMirror JSON.
func (r *Revision) OriginalProseMirror() ([]byte, error) {
	r.doc.mu.RLock()
	defer r.doc.mu.RUnlock()
	if r.typ != RevisionDelete {
		return json.Marshal(PMNode{Type: "doc"})
	}
	var content []PMNode
	for _, run := range r.el.Runs {
		text := run.DelText()
		if text == "" {
			text = run.Text()
		}
		if text == "" {
			continue
		}
		n := PMNode{Type: "text", Text: text}
		marks := runPropsToPMMarks(run.RPr, true)
		if len(marks) > 0 {
			n.Marks = marks
		}
		content = append(content, n)
	}
	p := PMNode{Type: "paragraph", Content: content}
	return json.Marshal(PMNode{Type: "doc", Content: []PMNode{p}})
}

// ProseMirror returns the matched text as ProseMirror JSON with marks.
func (sr *SearchResult) ProseMirror() ([]byte, error) {
	if len(sr.Runs) == 0 {
		return json.Marshal(PMNode{Type: "text"})
	}
	// Use local copies of Start/End to avoid mutating the receiver.
	start, end := sr.Start, sr.End

	if len(sr.Runs) == 1 {
		sr.Runs[0].doc.mu.RLock()
		defer sr.Runs[0].doc.mu.RUnlock()
		text := sr.Runs[0].el.Text()
		if end > len(text) {
			end = len(text)
		}
		text = text[start:end]
		node := PMNode{Type: "text", Text: text}
		marks := runPropsToPMMarks(sr.Runs[0].el.RPr, true)
		if len(marks) > 0 {
			node.Marks = marks
		}
		return json.Marshal(node)
	}
	// Multi-run: build paragraph with sliced text nodes.
	sr.Runs[0].doc.mu.RLock()
	defer sr.Runs[0].doc.mu.RUnlock()
	var content []PMNode
	for i, r := range sr.Runs {
		text := r.el.Text()
		switch {
		case i == 0:
			text = text[start:]
		case i == len(sr.Runs)-1:
			if end > len(text) {
				end = len(text)
			}
			text = text[:end]
		}
		if text == "" {
			continue
		}
		n := PMNode{Type: "text", Text: text}
		marks := runPropsToPMMarks(r.el.RPr, true)
		if len(marks) > 0 {
			n.Marks = marks
		}
		content = append(content, n)
	}
	return json.Marshal(PMNode{Type: "paragraph", Content: content})
}

// HeaderFooter.ProseMirror returns the header/footer content as ProseMirror JSON.
func (hf *HeaderFooter) ProseMirror() ([]byte, error) {
	hf.doc.mu.RLock()
	defer hf.doc.mu.RUnlock()
	o := &ProseMirrorOptions{UseTipTapNames: true}
	if hf.part.body == nil {
		return json.Marshal(PMNode{Type: "doc"})
	}
	urls := hf.doc.hyperlinkURLMap()
	doc := PMNode{
		Type:    "doc",
		Content: bodyToPMNodes(hf.part.body, urls, o),
	}
	return json.Marshal(doc)
}

// ---------------------------------------------------------------------------
// Internal conversion helpers
// ---------------------------------------------------------------------------

// bodyToPMNodes converts a CT_Body's content to PM nodes.
func bodyToPMNodes(body *wml.CT_Body, urls map[string]string, o *ProseMirrorOptions) []PMNode {
	var nodes []PMNode
	for _, bc := range body.Content {
		switch {
		case bc.Paragraph != nil:
			nodes = append(nodes, paragraphToPMNode(bc.Paragraph, urls, o))
		case bc.Table != nil:
			nodes = append(nodes, tableToPMNode(bc.Table, urls, o))
		}
	}
	return nodes
}

// paragraphToPMNode converts a CT_P to a PM node (paragraph, heading, or blockquote wrapper).
func paragraphToPMNode(p *wml.CT_P, urls map[string]string, o *ProseMirrorOptions) PMNode {
	useTT := o == nil || o.UseTipTapNames

	// Determine node type from style.
	nodeType := exportName("paragraph", useTT)
	var attrs map[string]any
	style := ""
	if p.PPr != nil && p.PPr.Style != nil {
		style = *p.PPr.Style
	}

	if level := headingLevel(style); level > 0 {
		nodeType = exportName("heading", useTT)
		attrs = map[string]any{"level": level}
	}

	// Paragraph attrs (alignment, etc.).
	if p.PPr != nil && p.PPr.Alignment != nil {
		a := *p.PPr.Alignment
		if a == "both" {
			a = "justify"
		}
		if attrs == nil {
			attrs = map[string]any{}
		}
		attrs["textAlign"] = a
	}

	// Non-heading/non-quote style preservation.
	if style != "" && headingLevel(style) == 0 && !isBlockquoteStyle(style) && nodeType == exportName("paragraph", useTT) {
		if attrs == nil {
			attrs = map[string]any{}
		}
		attrs["styleId"] = style
	}

	// Build inline content.
	var content []PMNode
	for _, ic := range p.Content {
		switch {
		case ic.Run != nil:
			n := runToPMNode(ic.Run, useTT)
			if n.Text != "" {
				content = append(content, n)
			}
			// Emit hard breaks (line breaks) found in the run's content.
			for _, rc := range ic.Run.Content {
				if rc.Break != nil && rc.Break.Type == "" {
					content = append(content, PMNode{Type: exportName("hardBreak", useTT)})
				}
			}
		case ic.Ins != nil:
			for _, r := range ic.Ins.Runs {
				n := runToPMNode(r, useTT)
				if n.Text != "" {
					content = append(content, n)
				}
			}
		case ic.Hyperlink != nil:
			content = append(content, hyperlinkToPMNodes(ic.Hyperlink, urls, useTT)...)
		}
	}

	// Wrap blockquotes.
	if isBlockquoteStyle(style) {
		inner := PMNode{Type: exportName("paragraph", useTT), Content: content}
		return PMNode{
			Type:    exportName("blockquote", useTT),
			Content: []PMNode{inner},
		}
	}

	node := PMNode{Type: nodeType, Content: content}
	if len(attrs) > 0 {
		node.Attrs = attrs
	}
	return node
}

// runToPMNode converts a CT_R to a text node with marks.
func runToPMNode(r *wml.CT_R, useTipTap bool) PMNode {
	text := runPlainText(r)
	node := PMNode{Type: "text", Text: text}
	if r.RPr == nil {
		return node
	}
	marks := runPropsToPMMarks(r.RPr, useTipTap)
	if len(marks) > 0 {
		node.Marks = marks
	}
	return node
}

// runPropsToPMMarks converts CT_RPr to PM marks.
func runPropsToPMMarks(rpr *wml.CT_RPr, useTipTap bool) []PMMark {
	if rpr == nil {
		return nil
	}
	var marks []PMMark

	// Code detection (before other marks — code runs don't get textStyle).
	if isCodeRunProps(rpr) {
		marks = append(marks, PMMark{Type: "code"})
		return marks
	}

	if rpr.Bold != nil && *rpr.Bold {
		marks = append(marks, PMMark{Type: exportName("bold", useTipTap)})
	}
	if rpr.Italic != nil && *rpr.Italic {
		marks = append(marks, PMMark{Type: exportName("italic", useTipTap)})
	}
	if rpr.Underline != nil && *rpr.Underline != "none" {
		marks = append(marks, PMMark{Type: "underline"})
	}
	if rpr.Strike != nil && *rpr.Strike {
		marks = append(marks, PMMark{Type: "strike"})
	}

	// textStyle mark (combines color, fontSize, fontFamily).
	tsAttrs := map[string]any{}
	if rpr.Color != nil {
		tsAttrs["color"] = "#" + *rpr.Color
	}
	if rpr.FontSize != nil {
		tsAttrs["fontSize"] = pmHalfPointsToPtString(*rpr.FontSize)
	}
	if rpr.FontName != nil {
		tsAttrs["fontFamily"] = *rpr.FontName
	}
	if len(tsAttrs) > 0 {
		marks = append(marks, PMMark{Type: "textStyle", Attrs: tsAttrs})
	}

	return marks
}

// isCodeRunProps checks if run properties indicate code formatting.
func isCodeRunProps(rpr *wml.CT_RPr) bool {
	if rpr.RunStyle != nil && *rpr.RunStyle == "CodeChar" {
		return true
	}
	if rpr.FontName != nil {
		fn := *rpr.FontName
		if fn == "Courier New" || fn == "Consolas" {
			return true
		}
	}
	return false
}

// hyperlinkToPMNodes converts a CT_Hyperlink to PM text nodes with a link mark.
func hyperlinkToPMNodes(h *wml.CT_Hyperlink, urls map[string]string, useTipTap bool) []PMNode {
	url := ""
	if urls != nil && h.ID != "" {
		url = urls[h.ID]
	}
	if url == "" {
		url = h.Anchor
	}

	var nodes []PMNode
	for _, r := range h.Runs {
		text := runPlainText(r)
		if text == "" {
			continue
		}
		n := PMNode{Type: "text", Text: text}
		// Start with run formatting marks.
		marks := runPropsToPMMarks(r.RPr, useTipTap)
		// Add link mark.
		linkMark := PMMark{Type: "link"}
		if url != "" {
			linkMark.Attrs = map[string]any{"href": url}
		}
		marks = append(marks, linkMark)
		n.Marks = marks
		nodes = append(nodes, n)
	}
	return nodes
}

// tableToPMNode converts a CT_Tbl to a PM table node.
func tableToPMNode(tbl *wml.CT_Tbl, urls map[string]string, o *ProseMirrorOptions) PMNode {
	useTT := o == nil || o.UseTipTapNames

	node := PMNode{Type: exportName("table", useTT)}

	// Table style attr.
	if tbl.TblPr != nil && tbl.TblPr.Style != nil {
		node.Attrs = map[string]any{"tableStyle": *tbl.TblPr.Style}
	}

	for i, row := range tbl.Rows {
		rowNode := PMNode{Type: exportName("tableRow", useTT)}
		for _, tc := range row.Cells {
			isHeader := i == 0
			rowNode.Content = append(rowNode.Content, cellToPMNode(tc, urls, o, isHeader))
		}
		node.Content = append(node.Content, rowNode)
	}
	return node
}

// cellToPMNode converts a CT_Tc to a PM tableCell or tableHeader node.
func cellToPMNode(tc *wml.CT_Tc, urls map[string]string, o *ProseMirrorOptions, isHeader bool) PMNode {
	useTT := o == nil || o.UseTipTapNames

	nodeType := exportName("tableCell", useTT)
	if isHeader {
		nodeType = exportName("tableHeader", useTT)
	}

	attrs := map[string]any{
		"colspan": 1,
		"rowspan": 1,
	}

	if tc.TcPr != nil {
		if tc.TcPr.GridSpan != nil {
			attrs["colspan"] = *tc.TcPr.GridSpan
		}
		if tc.TcPr.Shading != nil && tc.TcPr.Shading.Fill != "" {
			attrs["background"] = "#" + tc.TcPr.Shading.Fill
		}
	}

	var content []PMNode
	for _, bc := range tc.Content {
		if bc.Paragraph != nil {
			content = append(content, paragraphToPMNode(bc.Paragraph, urls, o))
		}
	}

	return PMNode{
		Type:    nodeType,
		Attrs:   attrs,
		Content: content,
	}
}

// addSectionAttrs adds section properties to a doc node's attrs.
func addSectionAttrs(attrs map[string]any, sp *wml.CT_SectPr) {
	if sp.PgSz != nil {
		if sp.PgSz.W != "" {
			if w, err := strconv.ParseInt(sp.PgSz.W, 10, 64); err == nil {
				attrs["pageWidth"] = pmTwipsToPoints(w)
			}
		}
		if sp.PgSz.H != "" {
			if h, err := strconv.ParseInt(sp.PgSz.H, 10, 64); err == nil {
				attrs["pageHeight"] = pmTwipsToPoints(h)
			}
		}
		if sp.PgSz.Orient != "" {
			attrs["orientation"] = sp.PgSz.Orient
		}
	}
	if sp.PgMar != nil {
		margins := map[string]any{}
		for _, pair := range []struct{ key, val string }{
			{"top", sp.PgMar.Top}, {"right", sp.PgMar.Right},
			{"bottom", sp.PgMar.Bottom}, {"left", sp.PgMar.Left},
		} {
			if pair.val != "" {
				if v, err := strconv.ParseInt(pair.val, 10, 64); err == nil {
					margins[pair.key] = pmTwipsToPoints(v)
				}
			}
		}
		if len(margins) > 0 {
			attrs["margins"] = margins
		}
	}
	if sp.Cols != nil && sp.Cols.Num != nil && *sp.Cols.Num > 1 {
		cols := map[string]any{"count": *sp.Cols.Num}
		if sp.Cols.Space != nil {
			if v, err := strconv.ParseInt(*sp.Cols.Space, 10, 64); err == nil {
				cols["space"] = pmTwipsToPoints(v)
			}
		}
		attrs["columns"] = cols
	}
}

// addHdrFtrAttrs adds header/footer content to doc attrs.
func addHdrFtrAttrs(d *Document, doc *PMNode, o *ProseMirrorOptions) {
	urls := d.hyperlinkURLMap()
	header := map[string]any{"default": nil, "first": nil, "even": nil}
	footer := map[string]any{"default": nil, "first": nil, "even": nil}

	for i := range d.hdrFtrs {
		hf := &d.hdrFtrs[i]
		if hf.body == nil {
			continue
		}
		content := bodyToPMNodes(hf.body, urls, o)
		pmDoc := PMNode{Type: "doc", Content: content}

		key := "default"
		switch hf.typ {
		case HdrFtrFirst:
			key = "first"
		case HdrFtrEven:
			key = "even"
		}

		if hf.isHdr {
			header[key] = pmDoc
		} else {
			footer[key] = pmDoc
		}
	}

	hasHdr := header["default"] != nil || header["first"] != nil || header["even"] != nil
	hasFtr := footer["default"] != nil || footer["first"] != nil || footer["even"] != nil

	if hasHdr {
		doc.Attrs["header"] = header
	}
	if hasFtr {
		doc.Attrs["footer"] = footer
	}
}

// trimColorHash removes "#" prefix from a color string if present.
func trimColorHash(color string) string {
	return strings.TrimPrefix(color, "#")
}
