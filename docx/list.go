package docx

// ListFormat describes the bullet or numbering style of a list item.
type ListFormat int

const (
	// ListBullet is the default list format (unordered / bulleted).
	ListBullet ListFormat = iota
)

// ListInfo holds list item metadata read from a paragraph's numbering properties.
type ListInfo struct {
	// Level is the zero-based indentation level of the list item.
	Level int
	// NumID is the identifier of the abstract numbering definition.
	NumID int
	// Format is the detected list format. It defaults to ListBullet because
	// full numbering resolution requires parsing numbering.xml.
	Format ListFormat
}

// ListInfo returns list metadata for the paragraph, or nil if the paragraph is
// not a list item (i.e. it has no <w:numPr> with a numId).
func (p *Paragraph) ListInfo() *ListInfo {
	p.doc.mu.RLock()
	defer p.doc.mu.RUnlock()

	if p.el.PPr == nil || p.el.PPr.NumPr == nil {
		return nil
	}
	np := p.el.PPr.NumPr
	if np.NumID == nil {
		return nil
	}

	level := 0
	if np.ILvl != nil {
		level = *np.ILvl
	}

	return &ListInfo{
		Level:  level,
		NumID:  *np.NumID,
		Format: ListBullet, // simplified — full resolution needs numbering.xml
	}
}

// IsListItem reports whether the paragraph is a list item.
func (p *Paragraph) IsListItem() bool {
	return p.ListInfo() != nil
}
