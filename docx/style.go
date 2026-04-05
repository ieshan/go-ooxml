package docx

import (
	"iter"
	"strings"

	"github.com/ieshan/go-ooxml/docx/wml"
)

// Style wraps a wml.CT_Style element.
type Style struct {
	doc *Document
	el  *wml.CT_Style
}

// ID returns the style's unique identifier (e.g., "Heading1").
func (s *Style) ID() string {
	return s.el.StyleID
}

// Name returns the display name of the style (e.g., "heading 1").
func (s *Style) Name() string {
	return s.el.Name
}

// Type returns the style type: "paragraph", "character", "table", or "numbering".
func (s *Style) Type() string {
	return s.el.Type
}

// BasedOn returns the parent style ID, or "" if none.
func (s *Style) BasedOn() string {
	if s.el.BasedOn == nil {
		return ""
	}
	return *s.el.BasedOn
}

// IsHeading returns true if the style is a heading style (Heading1-Heading9 or Title).
func (s *Style) IsHeading() bool {
	return s.HeadingLevel() > 0
}

// HeadingLevel returns 1-9 for Heading1-Heading9, 0 for non-heading styles.
// Detection is based on the style ID (e.g., "Heading1") or display name (e.g., "heading 1").
func (s *Style) HeadingLevel() int {
	// Check by style ID first (most reliable): "Heading1" .. "Heading9"
	id := s.el.StyleID
	if strings.HasPrefix(id, "Heading") && len(id) == 8 {
		ch := id[7]
		if ch >= '1' && ch <= '9' {
			return int(ch - '0')
		}
	}

	// Check by display name: "heading 1" .. "heading 9"
	name := strings.ToLower(strings.TrimSpace(s.el.Name))
	if strings.HasPrefix(name, "heading ") && len(name) == 9 {
		ch := name[8]
		if ch >= '1' && ch <= '9' {
			return int(ch - '0')
		}
	}

	return 0
}

// Styles returns an iterator over all styles defined in the document.
func (d *Document) Styles() iter.Seq[*Style] {
	return func(yield func(*Style) bool) {
		d.mu.RLock()
		defer d.mu.RUnlock()

		if d.styles == nil {
			return
		}
		for _, el := range d.styles.Styles {
			s := &Style{doc: d, el: el}
			if !yield(s) {
				return
			}
		}
	}
}

// StyleByID finds and returns a style by its styleId attribute, or nil if not found.
func (d *Document) StyleByID(id string) *Style {
	d.mu.RLock()
	defer d.mu.RUnlock()

	if d.styles == nil {
		return nil
	}
	for _, el := range d.styles.Styles {
		if el.StyleID == id {
			return &Style{doc: d, el: el}
		}
	}
	return nil
}

// StyleByName finds and returns a style by its display name, or nil if not found.
func (d *Document) StyleByName(name string) *Style {
	d.mu.RLock()
	defer d.mu.RUnlock()

	if d.styles == nil {
		return nil
	}
	for _, el := range d.styles.Styles {
		if el.Name == name {
			return &Style{doc: d, el: el}
		}
	}
	return nil
}
