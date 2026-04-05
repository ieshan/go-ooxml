package docx

import "strings"

// TextOptions controls what content is included in document-level text extraction.
type TextOptions struct {
	// IncludeHeaders includes header text. Currently a no-op (headers not yet implemented).
	IncludeHeaders bool
	// IncludeFooters includes footer text. Currently a no-op (footers not yet implemented).
	IncludeFooters bool
	// IncludeComments appends comment text after the body text.
	IncludeComments bool
}

// Text returns the full document text.
// If opts is nil, only body text is extracted.
// Body paragraphs are separated by newlines; table rows by newlines, cells by tabs.
// When IncludeComments is set, comment text is appended (newline-separated) after the body.
// IncludeHeaders and IncludeFooters are no-ops until headers/footers are implemented.
func (d *Document) Text(opts *TextOptions) string {
	d.mu.RLock()
	defer d.mu.RUnlock()

	var sb strings.Builder

	// Extract body text.
	if d.doc != nil && d.doc.Body != nil {
		sb.WriteString(bodyText(d.doc.Body))
	}

	if opts == nil {
		return sb.String()
	}

	// IncludeHeaders and IncludeFooters are no-ops until Phase 8 implements them.

	// Append comment text if requested.
	if opts.IncludeComments && d.comments != nil {
		for _, c := range d.comments.Comments {
			// Build text from comment paragraphs directly (no lock needed — we hold RLock).
			var csb strings.Builder
			first := true
			for _, bc := range c.Content {
				if bc.Paragraph != nil {
					if !first {
						csb.WriteByte('\n')
					}
					csb.WriteString(bc.Paragraph.Text())
					first = false
				}
			}
			text := csb.String()
			if text == "" {
				continue
			}
			if sb.Len() > 0 {
				sb.WriteByte('\n')
			}
			sb.WriteString(text)
		}
	}

	return sb.String()
}
