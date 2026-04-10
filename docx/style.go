package docx

import (
	"encoding/xml"
	"fmt"
	"iter"
	"strconv"
	"strings"

	"github.com/ieshan/go-ooxml/common"
	"github.com/ieshan/go-ooxml/docx/wml"
)

// Style wraps a wml.CT_Style element.
type Style struct {
	doc *Document
	el  *wml.CT_Style
}

// ID returns the style's unique identifier (e.g., "Heading1").
func (s *Style) ID() string {
	s.doc.mu.RLock()
	defer s.doc.mu.RUnlock()
	return s.el.StyleID
}

// Name returns the display name of the style (e.g., "heading 1").
func (s *Style) Name() string {
	s.doc.mu.RLock()
	defer s.doc.mu.RUnlock()
	return s.el.Name
}

// Type returns the style type: "paragraph", "character", "table", or "numbering".
func (s *Style) Type() string {
	s.doc.mu.RLock()
	defer s.doc.mu.RUnlock()
	return s.el.Type
}

// BasedOn returns the parent style ID, or "" if none.
func (s *Style) BasedOn() string {
	s.doc.mu.RLock()
	defer s.doc.mu.RUnlock()
	if s.el.BasedOn == nil {
		return ""
	}
	return *s.el.BasedOn
}

// RPr returns the style's run properties, or nil if not set.
func (s *Style) RPr() *wml.CT_RPr {
	s.doc.mu.RLock()
	defer s.doc.mu.RUnlock()
	return s.el.RPr
}

// PPr returns the style's paragraph properties, or nil if not set.
func (s *Style) PPr() *wml.CT_PPr {
	s.doc.mu.RLock()
	defer s.doc.mu.RUnlock()
	return s.el.PPr
}

// IsHeading returns true if the style is a heading style (Heading1-Heading9 or Title).
func (s *Style) IsHeading() bool {
	s.doc.mu.RLock()
	defer s.doc.mu.RUnlock()
	return s.headingLevel() > 0
}

// HeadingLevel returns 1-9 for Heading1-Heading9, 1 for the Title style,
// or 0 for non-heading styles.
// Detection is based on the style ID (e.g., "Heading1", "Title") or display
// name (e.g., "heading 1", "title").
func (s *Style) HeadingLevel() int {
	s.doc.mu.RLock()
	defer s.doc.mu.RUnlock()
	return s.headingLevel()
}

// headingLevel is the lock-free implementation of HeadingLevel.
// Caller must hold at least RLock.
func (s *Style) headingLevel() int {
	id := s.el.StyleID

	// "Title" maps to level 1, consistent with Markdown rendering.
	if id == "Title" {
		return 1
	}

	// Check by style ID: "Heading1" .. "Heading9"
	if strings.HasPrefix(id, "Heading") && len(id) == 8 {
		ch := id[7]
		if ch >= '1' && ch <= '9' {
			return int(ch - '0')
		}
	}

	// Check by display name: "heading 1" .. "heading 9", "title"
	name := strings.ToLower(strings.TrimSpace(s.el.Name))
	if name == "title" {
		return 1
	}
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

// AddStyle creates a new style definition with the given ID and type,
// adds it to the document, and returns the Style for further configuration.
// The id must be non-empty; styleType should be one of "paragraph", "character",
// "table", or "numbering". Panics if id is empty.
func (d *Document) AddStyle(id string, styleType string) *Style {
	if id == "" {
		panic("docx: AddStyle requires a non-empty style ID")
	}

	d.mu.Lock()
	defer d.mu.Unlock()

	if d.styles == nil {
		d.initDefaultStyles()
	}

	el := &wml.CT_Style{
		XMLName: xml.Name{Space: wml.Ns, Local: "style"},
		Type:    styleType,
		StyleID: id,
		Name:    id,
	}
	d.styles.Styles = append(d.styles.Styles, el)
	return &Style{doc: d, el: el}
}

// EnsureBuiltinStyle ensures that a built-in style definition exists.
// Returns the Style if found or created, nil if the ID is not a known built-in.
func (d *Document) EnsureBuiltinStyle(id string) *Style {
	d.mu.Lock()
	defer d.mu.Unlock()

	if d.styles == nil {
		d.initDefaultStyles()
	}

	for _, el := range d.styles.Styles {
		if el.StyleID == id {
			return &Style{doc: d, el: el}
		}
	}

	def, ok := builtinRegistry[id]
	if !ok {
		return nil
	}

	el := buildCTStyle(def)
	d.styles.Styles = append(d.styles.Styles, el)
	return &Style{doc: d, el: el}
}

// --- Style property setters ---

func (s *Style) ensurePPr() *wml.CT_PPr {
	if s.el.PPr == nil {
		s.el.PPr = &wml.CT_PPr{XMLName: xml.Name{Space: wml.Ns, Local: "pPr"}}
	}
	return s.el.PPr
}

func (s *Style) ensureRPr() *wml.CT_RPr {
	if s.el.RPr == nil {
		s.el.RPr = &wml.CT_RPr{}
	}
	return s.el.RPr
}

// SetBasedOn sets the parent style ID that this style inherits properties from.
// For example, heading styles typically inherit from "Normal".
func (s *Style) SetBasedOn(id string) {
	s.doc.mu.Lock()
	defer s.doc.mu.Unlock()
	s.el.BasedOn = &id
}

// SetNext sets the style ID that is automatically applied to the next paragraph
// when the user presses Enter. For example, a heading style typically sets next
// to "Normal" so body text follows the heading.
func (s *Style) SetNext(id string) {
	s.doc.mu.Lock()
	defer s.doc.mu.Unlock()
	s.el.Next = &id
}

// SetFont sets the font family name for both ASCII/Latin and High-ANSI character
// ranges (e.g., "Calibri", "Times New Roman").
func (s *Style) SetFont(name string) {
	s.doc.mu.Lock()
	defer s.doc.mu.Unlock()
	rpr := s.ensureRPr()
	rpr.FontName = &name
	rpr.FontHAnsi = &name
}

// SetFontSize sets the font size in points (e.g., 12 for 12pt text).
// OOXML stores font sizes as half-points; this method converts from points.
func (s *Style) SetFontSize(points float64) {
	s.doc.mu.Lock()
	defer s.doc.mu.Unlock()
	rpr := s.ensureRPr()
	rpr.FontSize = new(fmt.Sprintf("%g", points*2))
}

// SetColor sets the font color for text rendered with this style.
func (s *Style) SetColor(c common.Color) {
	s.doc.mu.Lock()
	defer s.doc.mu.Unlock()
	s.ensureRPr().Color = new(c.Hex())
}

// SetBold enables or disables bold formatting for this style's run properties.
func (s *Style) SetBold(v bool) {
	s.doc.mu.Lock()
	defer s.doc.mu.Unlock()
	s.ensureRPr().Bold = &v
}

// SetItalic enables or disables italic formatting for this style's run properties.
func (s *Style) SetItalic(v bool) {
	s.doc.mu.Lock()
	defer s.doc.mu.Unlock()
	s.ensureRPr().Italic = &v
}

// SetAlignment sets the paragraph horizontal alignment for this style.
// Common values: "left", "center", "right", "both" (justified).
func (s *Style) SetAlignment(align string) {
	s.doc.mu.Lock()
	defer s.doc.mu.Unlock()
	s.ensurePPr().Alignment = &align
}

// SetSpacingBefore sets the spacing above a paragraph in twips (1/1440 inch).
// For example, 240 twips equals roughly 1/6 inch or one line at 12pt.
func (s *Style) SetSpacingBefore(twips int) {
	s.doc.mu.Lock()
	defer s.doc.mu.Unlock()
	ppr := s.ensurePPr()
	if ppr.Spacing == nil {
		ppr.Spacing = &wml.CT_Spacing{}
	}
	ppr.Spacing.Before = new(strconv.Itoa(twips))
}

// SetSpacingAfter sets the spacing below a paragraph in twips (1/1440 inch).
func (s *Style) SetSpacingAfter(twips int) {
	s.doc.mu.Lock()
	defer s.doc.mu.Unlock()
	ppr := s.ensurePPr()
	if ppr.Spacing == nil {
		ppr.Spacing = &wml.CT_Spacing{}
	}
	ppr.Spacing.After = new(strconv.Itoa(twips))
}

// SetLineSpacing sets the line spacing for paragraphs using this style.
// The val parameter is in twips for "exact" and "atLeast" rules, or in 240ths
// of a line for "auto" (e.g., 240 = single, 480 = double).
func (s *Style) SetLineSpacing(val int, rule string) {
	s.doc.mu.Lock()
	defer s.doc.mu.Unlock()
	ppr := s.ensurePPr()
	if ppr.Spacing == nil {
		ppr.Spacing = &wml.CT_Spacing{}
	}
	ppr.Spacing.Line = new(strconv.Itoa(val))
	ppr.Spacing.LineRule = &rule
}

// SetOutlineLevel sets the outline level (0-8) used by Word's navigation pane
// and table of contents. Level 0 corresponds to Heading 1, level 1 to Heading 2, etc.
func (s *Style) SetOutlineLevel(level int) {
	s.doc.mu.Lock()
	defer s.doc.mu.Unlock()
	s.ensurePPr().OutlineLvl = &level
}

// SetKeepNext prevents a page break between this paragraph and the next.
// This is commonly enabled for heading styles to avoid orphaned headings.
func (s *Style) SetKeepNext(v bool) {
	s.doc.mu.Lock()
	defer s.doc.mu.Unlock()
	s.ensurePPr().KeepNext = &v
}

// SetKeepLines prevents a page break within this paragraph, keeping all its
// lines on the same page.
func (s *Style) SetKeepLines(v bool) {
	s.doc.mu.Lock()
	defer s.doc.mu.Unlock()
	s.ensurePPr().KeepLines = &v
}

// SetIndentLeft sets the left indentation for paragraphs in twips (1/1440 inch).
// For example, 720 twips equals 0.5 inch.
func (s *Style) SetIndentLeft(twips int) {
	s.doc.mu.Lock()
	defer s.doc.mu.Unlock()
	ppr := s.ensurePPr()
	if ppr.Ind == nil {
		ppr.Ind = &wml.CT_Ind{}
	}
	ppr.Ind.Left = new(strconv.Itoa(twips))
}

// SetIndentRight sets the right indentation for paragraphs in twips (1/1440 inch).
func (s *Style) SetIndentRight(twips int) {
	s.doc.mu.Lock()
	defer s.doc.mu.Unlock()
	ppr := s.ensurePPr()
	if ppr.Ind == nil {
		ppr.Ind = &wml.CT_Ind{}
	}
	ppr.Ind.Right = new(strconv.Itoa(twips))
}
