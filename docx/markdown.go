package docx

import (
	"encoding/xml"
	"fmt"
	"strings"

	"github.com/ieshan/go-ooxml/docx/wml"
)

// wmlRunName is the XML name for a <w:r> element, used when building run
// elements without going through the proxy layer (e.g. inside locked contexts).
var wmlRunName = xml.Name{Space: wml.Ns, Local: "r"}

// MarkdownOptions configures Document.Markdown() output.
type MarkdownOptions struct {
	IncludeHeaders  bool
	IncludeFooters  bool
	IncludeComments bool   // Render comments as footnotes
	HorizontalRule  string // Default: "---"
}

func (o *MarkdownOptions) hrule() string {
	if o != nil && o.HorizontalRule != "" {
		return o.HorizontalRule
	}
	return "---"
}

// Markdown renders the full document as Markdown.
func (d *Document) Markdown(opts *MarkdownOptions) string {
	d.mu.RLock()
	defer d.mu.RUnlock()

	var sb strings.Builder

	// Body content.
	sb.WriteString(bodyToMarkdown(d.doc.Body, d.hyperlinkURLMap()))

	// Comments as footnotes.
	if opts != nil && opts.IncludeComments && d.comments != nil && len(d.comments.Comments) > 0 {
		// Build comment ID -> footnote index mapping.
		// Also collect comment range info from body paragraphs.
		type commentRef struct {
			id     int
			author string
			text   string
		}
		var refs []commentRef
		for _, c := range d.comments.Comments {
			text := commentBodyText(c)
			refs = append(refs, commentRef{id: c.ID, author: c.Author, text: text})
		}

		if len(refs) > 0 {
			sb.WriteString("\n")
			for i, ref := range refs {
				sb.WriteString(fmt.Sprintf("\n[^%d]: %s: %s", i+1, ref.author, ref.text))
			}
		}
	}

	return sb.String()
}

// commentBodyText extracts plain text from a CT_Comment's body.
func commentBodyText(c *wml.CT_Comment) string {
	var sb strings.Builder
	first := true
	for _, bc := range c.Content {
		if bc.Paragraph != nil {
			if !first {
				sb.WriteByte('\n')
			}
			sb.WriteString(bc.Paragraph.Text())
			first = false
		}
	}
	return sb.String()
}

// runToMarkdown converts a CT_R to Markdown text.
func runToMarkdown(r *wml.CT_R) string {
	text := runPlainText(r)
	if text == "" {
		return ""
	}

	// Check for code style.
	if isCodeRun(r) {
		return "`" + text + "`"
	}

	bold := r.RPr != nil && r.RPr.Bold != nil && *r.RPr.Bold
	italic := r.RPr != nil && r.RPr.Italic != nil && *r.RPr.Italic
	strike := r.RPr != nil && r.RPr.Strike != nil && *r.RPr.Strike

	if bold && italic {
		text = "***" + text + "***"
	} else if bold {
		text = "**" + text + "**"
	} else if italic {
		text = "*" + text + "*"
	}

	if strike {
		text = "~~" + text + "~~"
	}

	return text
}

// runPlainText extracts only Text content from a run (not DelText).
func runPlainText(r *wml.CT_R) string {
	var sb strings.Builder
	for _, c := range r.Content {
		if c.Text != nil {
			sb.WriteString(c.Text.Value)
		} else if c.Tab != nil {
			sb.WriteByte('\t')
		} else if c.Break != nil {
			sb.WriteByte('\n')
		}
	}
	return sb.String()
}

// runDeleteMarkdown converts a CT_R from a delete revision to Markdown text.
// Unlike runToMarkdown it reads DelText instead of Text.
func runDeleteMarkdown(r *wml.CT_R) string {
	text := r.DelText()
	if text == "" {
		// Fall back to Text in case the run uses plain Text nodes.
		text = runPlainText(r)
	}
	if text == "" {
		return ""
	}

	if isCodeRun(r) {
		return "`" + text + "`"
	}

	bold := r.RPr != nil && r.RPr.Bold != nil && *r.RPr.Bold
	italic := r.RPr != nil && r.RPr.Italic != nil && *r.RPr.Italic
	strike := r.RPr != nil && r.RPr.Strike != nil && *r.RPr.Strike

	if bold && italic {
		text = "***" + text + "***"
	} else if bold {
		text = "**" + text + "**"
	} else if italic {
		text = "*" + text + "*"
	}

	if strike {
		text = "~~" + text + "~~"
	}

	return text
}

// isCodeRun checks if the run should be rendered as inline code.
func isCodeRun(r *wml.CT_R) bool {
	if r.RPr == nil {
		return false
	}
	if r.RPr.RunStyle != nil && *r.RPr.RunStyle == "CodeChar" {
		return true
	}
	if r.RPr.FontName != nil {
		fn := *r.RPr.FontName
		if fn == "Courier New" || fn == "Consolas" {
			return true
		}
	}
	return false
}

// paragraphToMarkdown converts a CT_P to Markdown.
// hyperlinkURLs is an optional map from relationship ID to target URL used to
// render hyperlinks as [text](url). Pass nil to fall back to text-only rendering.
func paragraphToMarkdown(p *wml.CT_P, hyperlinkURLs map[string]string) string {
	// Concatenate run markdown.
	var sb strings.Builder
	for _, c := range p.Content {
		if c.Run != nil {
			sb.WriteString(runToMarkdown(c.Run))
		} else if c.Ins != nil {
			for _, r := range c.Ins.Runs {
				sb.WriteString(runToMarkdown(r))
			}
		} else if c.Hyperlink != nil {
			// Collect the text of all runs inside the hyperlink.
			var linkText strings.Builder
			for _, r := range c.Hyperlink.Runs {
				linkText.WriteString(runToMarkdown(r))
			}
			lt := linkText.String()

			// Resolve URL via relationship map when available.
			url := ""
			if hyperlinkURLs != nil && c.Hyperlink.ID != "" {
				url = hyperlinkURLs[c.Hyperlink.ID]
			}
			if url == "" {
				url = c.Hyperlink.Anchor
			}

			if url != "" {
				sb.WriteString("[" + lt + "](" + url + ")")
			} else {
				sb.WriteString(lt)
			}
		}
	}
	text := sb.String()

	// Apply paragraph-level formatting.
	style := ""
	if p.PPr != nil && p.PPr.Style != nil {
		style = *p.PPr.Style
	}

	// Heading styles.
	if level := headingLevel(style); level > 0 {
		return strings.Repeat("#", level) + " " + text
	}

	// Blockquote styles.
	if isBlockquoteStyle(style) {
		return "> " + text
	}

	// List items.
	if p.PPr != nil && p.PPr.NumPr != nil {
		return listItemMarkdown(p.PPr.NumPr, text)
	}

	return text
}

// headingLevel returns the heading level (1-9) for a style name, or 0 if not a heading.
func headingLevel(style string) int {
	switch style {
	case "Title":
		return 1
	case "Heading1":
		return 1
	case "Heading2":
		return 2
	case "Heading3":
		return 3
	case "Heading4":
		return 4
	case "Heading5":
		return 5
	case "Heading6":
		return 6
	case "Heading7":
		return 7
	case "Heading8":
		return 8
	case "Heading9":
		return 9
	}
	return 0
}

// isBlockquoteStyle returns true if the style indicates a blockquote.
func isBlockquoteStyle(style string) bool {
	switch style {
	case "Quote", "BlockQuote", "Blockquote", "IntenseQuote":
		return true
	}
	return false
}

// listItemMarkdown formats a list item with indentation.
func listItemMarkdown(numPr *wml.CT_NumPr, text string) string {
	level := 0
	if numPr.ILvl != nil {
		level = *numPr.ILvl
	}
	indent := strings.Repeat("  ", level)

	// NumID-based heuristic: odd numIDs are typically bullets, even are decimal.
	// This is a simplification; real numbering would require numbering.xml parsing.
	// Default to bullet list.
	marker := "- "
	if numPr.NumID != nil && *numPr.NumID%2 == 0 {
		marker = "1. "
	}

	return indent + marker + text
}

// cellToMarkdown converts a CT_Tc to inline Markdown (no block separators).
// Empty paragraphs are skipped.
func cellToMarkdown(tc *wml.CT_Tc, hyperlinkURLs map[string]string) string {
	var parts []string
	for _, c := range tc.Content {
		if c.Paragraph != nil {
			md := paragraphToMarkdown(c.Paragraph, hyperlinkURLs)
			if md != "" {
				parts = append(parts, md)
			}
		}
	}
	return strings.Join(parts, " ")
}

// tableToMarkdown converts a CT_Tbl to a pipe-table Markdown string.
func tableToMarkdown(tbl *wml.CT_Tbl, hyperlinkURLs map[string]string) string {
	if len(tbl.Rows) == 0 {
		return ""
	}

	var sb strings.Builder

	// Render the first row as a header.
	firstRow := tbl.Rows[0]
	sb.WriteByte('|')
	widths := make([]int, len(firstRow.Cells))
	headerCells := make([]string, len(firstRow.Cells))
	for j, tc := range firstRow.Cells {
		cell := cellToMarkdown(tc, hyperlinkURLs)
		headerCells[j] = cell
		w := len(cell)
		if w < 3 {
			w = 3
		}
		widths[j] = w
	}
	for j, cell := range headerCells {
		sb.WriteString(fmt.Sprintf(" %-*s |", widths[j], cell))
	}
	sb.WriteByte('\n')

	// Separator row.
	sb.WriteByte('|')
	for _, w := range widths {
		sb.WriteString(" " + strings.Repeat("-", w) + " |")
	}
	sb.WriteByte('\n')

	// Data rows.
	for i := 1; i < len(tbl.Rows); i++ {
		row := tbl.Rows[i]
		sb.WriteByte('|')
		for j, tc := range row.Cells {
			cell := cellToMarkdown(tc, hyperlinkURLs)
			w := 3
			if j < len(widths) {
				w = widths[j]
			}
			sb.WriteString(fmt.Sprintf(" %-*s |", w, cell))
		}
		sb.WriteByte('\n')
	}

	return strings.TrimRight(sb.String(), "\n")
}

// bodyToMarkdown converts a CT_Body to Markdown.
func bodyToMarkdown(body *wml.CT_Body, hyperlinkURLs map[string]string) string {
	var parts []string
	for _, c := range body.Content {
		switch {
		case c.Paragraph != nil:
			parts = append(parts, paragraphToMarkdown(c.Paragraph, hyperlinkURLs))
		case c.Table != nil:
			parts = append(parts, tableToMarkdown(c.Table, hyperlinkURLs))
		}
	}
	return strings.Join(parts, "\n\n")
}

// ---------------------------------------------------------------------------
// Markdown → DOCX import
// ---------------------------------------------------------------------------

// FromMarkdown creates a new Document from markdown content.
// cfg may be nil to use defaults.
func FromMarkdown(md string, cfg *Config) (*Document, error) {
	doc, err := New(cfg)
	if err != nil {
		return nil, err
	}
	if err := doc.ImportMarkdown(md); err != nil {
		_ = doc.Close()
		return nil, err
	}
	return doc, nil
}

// ImportMarkdown appends markdown content to the document body.
func (d *Document) ImportMarkdown(md string) error {
	return importMarkdownIntoBody(d, md)
}

// importMarkdownIntoBody parses block-level markdown and appends to document body.
func importMarkdownIntoBody(d *Document, md string) error {
	lines := strings.Split(md, "\n")
	body := d.Body()

	i := 0
	for i < len(lines) {
		line := lines[i]
		trimmed := strings.TrimSpace(line)

		// Code fence (```...```)
		if strings.HasPrefix(line, "```") {
			i++
			for i < len(lines) {
				codeLine := lines[i]
				if strings.HasPrefix(codeLine, "```") {
					i++
					break
				}
				p := body.AddParagraph()
				r := p.AddRun(codeLine)
				r.SetFontName("Courier New")
				i++
			}
			continue
		}

		// Horizontal rule: ---, ***, ___
		if isHorizontalRule(trimmed) {
			// Add an empty paragraph as a section separator
			body.AddParagraph()
			i++
			continue
		}

		// Heading
		if strings.HasPrefix(line, "#") {
			level := 0
			for level < len(line) && line[level] == '#' {
				level++
			}
			if level <= 6 && (len(line) <= level || line[level] == ' ') {
				text := strings.TrimSpace(line[level:])
				p := body.AddParagraph()
				p.SetStyle(headingStyleName(level))
				if err := applyInlineMarkdown(p, text); err != nil {
					return err
				}
				i++
				continue
			}
		}

		// Blockquote
		if strings.HasPrefix(line, ">") {
			text := strings.TrimSpace(strings.TrimPrefix(line, ">"))
			p := body.AddParagraph()
			p.SetStyle("Quote")
			if err := applyInlineMarkdown(p, text); err != nil {
				return err
			}
			i++
			continue
		}

		// Bullet list item
		if isBulletItem(line) {
			text := strings.TrimSpace(line[2:])
			p := body.AddParagraph()
			p.SetStyle("ListBullet")
			if err := applyInlineMarkdown(p, "• "+text); err != nil {
				return err
			}
			i++
			continue
		}

		// Numbered list item (simplified: "1. " or "2. " etc.)
		if isNumberedListItem(line) {
			dotIdx := strings.Index(line, ". ")
			text := strings.TrimSpace(line[dotIdx+2:])
			num := strings.TrimSpace(line[:dotIdx])
			p := body.AddParagraph()
			p.SetStyle("ListNumber")
			if err := applyInlineMarkdown(p, num+". "+text); err != nil {
				return err
			}
			i++
			continue
		}

		// Table: lines starting with |
		if strings.HasPrefix(trimmed, "|") {
			tableLines := []string{}
			for i < len(lines) && strings.HasPrefix(strings.TrimSpace(lines[i]), "|") {
				tableLines = append(tableLines, lines[i])
				i++
			}
			if err := parseMarkdownTable(d, tableLines); err != nil {
				return err
			}
			continue
		}

		// Blank line: skip
		if trimmed == "" {
			i++
			continue
		}

		// Normal paragraph: accumulate consecutive non-blank, non-special lines
		paraLines := []string{}
		for i < len(lines) {
			l := lines[i]
			lt := strings.TrimSpace(l)
			if lt == "" {
				break
			}
			if isMarkdownBlockStart(l) {
				break
			}
			paraLines = append(paraLines, l)
			i++
		}
		text := strings.Join(paraLines, " ")
		p := body.AddParagraph()
		if err := applyInlineMarkdown(p, text); err != nil {
			return err
		}
	}
	return nil
}

// isMarkdownBlockStart returns true if a line starts a block-level element that
// should not be merged into a continuing paragraph.
func isMarkdownBlockStart(line string) bool {
	trimmed := strings.TrimSpace(line)
	if trimmed == "" {
		return true
	}
	if strings.HasPrefix(line, "#") {
		return true
	}
	if strings.HasPrefix(line, ">") {
		return true
	}
	if isBulletItem(line) {
		return true
	}
	if isNumberedListItem(line) {
		return true
	}
	if strings.HasPrefix(trimmed, "|") {
		return true
	}
	if strings.HasPrefix(line, "```") {
		return true
	}
	if isHorizontalRule(trimmed) {
		return true
	}
	return false
}

// isBulletItem returns true for lines starting with "- ", "* ", or "+ ".
func isBulletItem(line string) bool {
	return strings.HasPrefix(line, "- ") || strings.HasPrefix(line, "* ") || strings.HasPrefix(line, "+ ")
}

// isHorizontalRule returns true for lines like ---, ***, ___ (3+ identical chars, no others).
func isHorizontalRule(s string) bool {
	if len(s) < 3 {
		return false
	}
	ch := s[0]
	if ch != '-' && ch != '*' && ch != '_' {
		return false
	}
	for i := 0; i < len(s); i++ {
		if s[i] != ch {
			return false
		}
	}
	return true
}

// isNumberedListItem returns true for lines like "1. text".
func isNumberedListItem(line string) bool {
	for i, c := range line {
		if c >= '0' && c <= '9' {
			continue
		}
		if c == '.' && i > 0 {
			return i+1 < len(line) && line[i+1] == ' '
		}
		return false
	}
	return false
}

// headingStyleName maps a heading level to an OOXML style name.
func headingStyleName(level int) string {
	switch level {
	case 1:
		return "Heading1"
	case 2:
		return "Heading2"
	case 3:
		return "Heading3"
	case 4:
		return "Heading4"
	case 5:
		return "Heading5"
	case 6:
		return "Heading6"
	default:
		return "Heading1"
	}
}

// parseMarkdownTable parses pipe-separated table lines and creates a Table in the document.
func parseMarkdownTable(d *Document, lines []string) error {
	// Filter separator rows (|---|---|)
	dataRows := [][]string{}
	for _, line := range lines {
		cells := splitTableRow(line)
		if isSeparatorRow(cells) {
			continue
		}
		dataRows = append(dataRows, cells)
	}
	if len(dataRows) == 0 {
		return nil
	}

	// Determine column count
	cols := 0
	for _, row := range dataRows {
		if len(row) > cols {
			cols = len(row)
		}
	}
	if cols == 0 {
		return nil
	}

	tbl := d.Body().AddTable(len(dataRows), cols)
	for r, row := range dataRows {
		for c := 0; c < cols; c++ {
			cell := tbl.Cell(r, c)
			if cell == nil {
				continue
			}
			text := ""
			if c < len(row) {
				text = strings.TrimSpace(row[c])
			}
			// Apply inline markdown to first paragraph in cell
			// Access WML struct directly to avoid lock re-entrancy
			if len(cell.el.Content) > 0 && cell.el.Content[0].Paragraph != nil {
				p := &Paragraph{doc: d, el: cell.el.Content[0].Paragraph}
				if err := applyInlineMarkdown(p, text); err != nil {
					return err
				}
			}
		}
	}
	return nil
}

// splitTableRow splits a pipe-delimited table row into cell strings.
func splitTableRow(line string) []string {
	line = strings.TrimSpace(line)
	line = strings.TrimPrefix(line, "|")
	line = strings.TrimSuffix(line, "|")
	return strings.Split(line, "|")
}

// isSeparatorRow returns true if all cells match the separator pattern (---, :---:, etc.).
func isSeparatorRow(cells []string) bool {
	if len(cells) == 0 {
		return false
	}
	for _, c := range cells {
		s := strings.TrimSpace(c)
		s = strings.Trim(s, ":")
		if len(s) == 0 {
			return false
		}
		allDash := true
		for _, ch := range s {
			if ch != '-' {
				allDash = false
				break
			}
		}
		if !allDash {
			return false
		}
	}
	return true
}

// applyInlineMarkdown parses inline markdown from text and appends runs to the paragraph.
// This acquires the document lock via Paragraph methods and must NOT be called while
// the document lock is already held. Use applyInlineMarkdownWML for lock-held contexts.
func applyInlineMarkdown(p *Paragraph, text string) error {
	runs := parseInlineMarkdown(text)
	for _, run := range runs {
		if run.link != "" {
			// Hyperlink: add as a proper hyperlink element with the URL.
			p.AddHyperlink(run.link, run.text)
			continue
		}
		r := p.AddRun(run.text)
		if run.bold {
			v := true
			r.SetBold(&v)
		}
		if run.italic {
			v := true
			r.SetItalic(&v)
		}
		if run.strike {
			v := true
			r.SetStrikethrough(&v)
		}
		if run.code {
			r.SetFontName("Courier New")
		}
	}
	return nil
}

// applyInlineMarkdownWML is like applyInlineMarkdown but operates directly on
// a *wml.CT_P without acquiring any lock. Use this when the document lock is
// already held. Hyperlinks are rendered as plain text runs (no relationship
// registration is possible without releasing the lock).
func applyInlineMarkdownWML(p *wml.CT_P, text string) {
	runs := parseInlineMarkdown(text)
	for _, run := range runs {
		r := &wml.CT_R{XMLName: wmlRunName}
		r.AddText(run.text)
		if run.bold || run.italic || run.strike || run.code {
			r.RPr = &wml.CT_RPr{}
			if run.bold {
				v := true
				r.RPr.Bold = &v
			}
			if run.italic {
				v := true
				r.RPr.Italic = &v
			}
			if run.strike {
				v := true
				r.RPr.Strike = &v
			}
			if run.code {
				fn := "Courier New"
				r.RPr.FontName = &fn
			}
		}
		p.Content = append(p.Content, wml.InlineContent{Run: r})
	}
}

// inlineRun holds parsed inline content.
type inlineRun struct {
	text   string
	bold   bool
	italic bool
	strike bool
	code   bool
	link   string // non-empty for [text](url) links
}

// parseInlineMarkdown parses inline markdown tokens from a string.
// Handles: **bold**, *italic*, ***bold+italic***, `code`, ~~strike~~, [text](url)
func parseInlineMarkdown(s string) []inlineRun {
	var runs []inlineRun
	i := 0
	var buf strings.Builder

	flushPlain := func() {
		if buf.Len() > 0 {
			runs = append(runs, inlineRun{text: buf.String()})
			buf.Reset()
		}
	}

	for i < len(s) {
		// Backtick code span
		if s[i] == '`' {
			flushPlain()
			i++
			start := i
			for i < len(s) && s[i] != '`' {
				i++
			}
			if i < len(s) {
				runs = append(runs, inlineRun{text: s[start:i], code: true})
				i++ // consume closing `
			} else {
				// No closing backtick — treat as literal
				runs = append(runs, inlineRun{text: "`" + s[start:]})
			}
			continue
		}

		// Strikethrough ~~
		if i+1 < len(s) && s[i] == '~' && s[i+1] == '~' {
			flushPlain()
			i += 2
			start := i
			for i+1 < len(s) && !(s[i] == '~' && s[i+1] == '~') {
				i++
			}
			if i+1 < len(s) {
				runs = append(runs, inlineRun{text: s[start:i], strike: true})
				i += 2 // consume closing ~~
			} else {
				runs = append(runs, inlineRun{text: "~~" + s[start:i]})
			}
			continue
		}

		// Bold+italic ***
		if i+2 < len(s) && s[i] == '*' && s[i+1] == '*' && s[i+2] == '*' {
			flushPlain()
			i += 3
			start := i
			for i+2 < len(s) && !(s[i] == '*' && s[i+1] == '*' && s[i+2] == '*') {
				i++
			}
			if i+2 < len(s) {
				runs = append(runs, inlineRun{text: s[start:i], bold: true, italic: true})
				i += 3
			} else {
				runs = append(runs, inlineRun{text: "***" + s[start:i]})
			}
			continue
		}

		// Bold **
		if i+1 < len(s) && s[i] == '*' && s[i+1] == '*' {
			flushPlain()
			i += 2
			start := i
			for i+1 < len(s) && !(s[i] == '*' && s[i+1] == '*') {
				i++
			}
			if i+1 < len(s) {
				runs = append(runs, inlineRun{text: s[start:i], bold: true})
				i += 2
			} else {
				runs = append(runs, inlineRun{text: "**" + s[start:i]})
			}
			continue
		}

		// Italic *
		if s[i] == '*' {
			flushPlain()
			i++
			start := i
			for i < len(s) && s[i] != '*' {
				i++
			}
			if i < len(s) {
				runs = append(runs, inlineRun{text: s[start:i], italic: true})
				i++ // consume closing *
			} else {
				runs = append(runs, inlineRun{text: "*" + s[start:]})
			}
			continue
		}

		// Link [text](url)
		if s[i] == '[' {
			flushPlain()
			i++
			linkTextStart := i
			for i < len(s) && s[i] != ']' {
				i++
			}
			linkText := s[linkTextStart:i]
			if i < len(s) && s[i] == ']' {
				i++
				if i < len(s) && s[i] == '(' {
					i++
					urlStart := i
					// consume url
					for i < len(s) && s[i] != ')' {
						i++
					}
					url := s[urlStart:i]
					if i < len(s) {
						i++ // consume ')'
					}
					runs = append(runs, inlineRun{text: linkText, link: url})
				} else {
					runs = append(runs, inlineRun{text: "[" + linkText + "]"})
				}
			} else {
				runs = append(runs, inlineRun{text: "[" + linkText})
			}
			continue
		}

		buf.WriteByte(s[i])
		i++
	}

	flushPlain()
	return runs
}

// SetMarkdown replaces paragraph content with inline-markdown-styled runs.
// Only handles inline markdown (bold, italic, code, links, strikethrough).
func (p *Paragraph) SetMarkdown(md string) error {
	p.doc.mu.Lock()
	p.el.Content = nil
	p.doc.mu.Unlock()

	return applyInlineMarkdown(p, md)
}

// InsertMarkdownAfter inserts block-level markdown after this paragraph.
// Returns the newly created paragraphs (tables are not included in the return slice).
func (p *Paragraph) InsertMarkdownAfter(md string) ([]*Paragraph, error) {
	// Parse into a temp document to get the block-level wml elements.
	tmp, err := New(nil)
	if err != nil {
		return nil, err
	}
	defer tmp.Close()

	if err := tmp.ImportMarkdown(md); err != nil {
		return nil, err
	}

	// Collect wml content blocks from the temp document.
	tmp.mu.RLock()
	newBlocks := make([]wml.BlockLevelContent, len(tmp.doc.Body.Content))
	copy(newBlocks, tmp.doc.Body.Content)
	tmp.mu.RUnlock()

	if len(newBlocks) == 0 {
		return nil, nil
	}

	// Find insertion point in this document's body.
	p.doc.mu.Lock()
	defer p.doc.mu.Unlock()

	body := p.doc.doc.Body
	idx := -1
	for i, c := range body.Content {
		if c.Paragraph == p.el {
			idx = i
			break
		}
	}

	newContent := make([]wml.BlockLevelContent, 0, len(body.Content)+len(newBlocks))
	if idx >= 0 {
		newContent = append(newContent, body.Content[:idx+1]...)
		newContent = append(newContent, newBlocks...)
		newContent = append(newContent, body.Content[idx+1:]...)
	} else {
		newContent = append(newContent, body.Content...)
		newContent = append(newContent, newBlocks...)
	}
	body.Content = newContent

	// Collect returned paragraphs (those backed by wml.CT_P).
	var result []*Paragraph
	for _, blk := range newBlocks {
		if blk.Paragraph != nil {
			result = append(result, &Paragraph{doc: p.doc, el: blk.Paragraph})
		}
	}
	return result, nil
}

// SetMarkdown replaces cell content with markdown.
func (c *Cell) SetMarkdown(md string) error {
	c.doc.mu.Lock()
	c.el.Content = nil
	c.doc.mu.Unlock()

	return importMarkdownIntoCell(c, md)
}

// importMarkdownIntoCell parses block-level markdown and populates a cell.
func importMarkdownIntoCell(c *Cell, md string) error {
	lines := strings.Split(md, "\n")
	i := 0
	for i < len(lines) {
		line := lines[i]
		trimmed := strings.TrimSpace(line)

		if trimmed == "" {
			i++
			continue
		}

		if strings.HasPrefix(line, "#") {
			level := 0
			for level < len(line) && line[level] == '#' {
				level++
			}
			if level <= 6 && (len(line) <= level || line[level] == ' ') {
				text := strings.TrimSpace(line[level:])
				p := c.AddParagraph()
				p.SetStyle(headingStyleName(level))
				if err := applyInlineMarkdown(p, text); err != nil {
					return err
				}
				i++
				continue
			}
		}

		p := c.AddParagraph()
		if err := applyInlineMarkdown(p, trimmed); err != nil {
			return err
		}
		i++
	}

	// OOXML requires at least one paragraph in a cell.
	hasContent := false
	for range c.Paragraphs() {
		hasContent = true
		break
	}
	if !hasContent {
		c.AddParagraph()
	}
	return nil
}
