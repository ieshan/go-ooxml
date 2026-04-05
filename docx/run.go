package docx

import (
	"fmt"
	"strconv"

	"github.com/ieshan/go-ooxml/common"
	"github.com/ieshan/go-ooxml/docx/wml"
)

// Run wraps a wml.CT_R element.
type Run struct {
	doc *Document
	el  *wml.CT_R
}

// Text returns the concatenated text content of the run.
func (r *Run) Text() string {
	r.doc.mu.RLock()
	defer r.doc.mu.RUnlock()
	return r.el.Text()
}

// Markdown returns the run's text content formatted as Markdown, applying
// bold, italic, strikethrough, and code formatting as appropriate.
func (r *Run) Markdown() string {
	r.doc.mu.RLock()
	defer r.doc.mu.RUnlock()
	return runToMarkdown(r.el)
}

// SetText replaces all text content in the run with the given string.
func (r *Run) SetText(text string) {
	r.doc.mu.Lock()
	defer r.doc.mu.Unlock()

	r.el.Content = nil
	r.el.AddText(text)
}

// ensureRPr ensures the run has a run properties element, creating one if needed.
// Caller must hold the write lock.
func (r *Run) ensureRPr() *wml.CT_RPr {
	if r.el.RPr == nil {
		r.el.RPr = &wml.CT_RPr{}
	}
	return r.el.RPr
}

// Bold returns the bold state: nil means inherit, non-nil is explicit.
func (r *Run) Bold() *bool {
	r.doc.mu.RLock()
	defer r.doc.mu.RUnlock()
	if r.el.RPr == nil {
		return nil
	}
	return r.el.RPr.Bold
}

// SetBold sets the bold state. Pass nil to clear (inherit from style).
func (r *Run) SetBold(v *bool) {
	r.doc.mu.Lock()
	defer r.doc.mu.Unlock()
	if v == nil {
		if r.el.RPr != nil {
			r.el.RPr.Bold = nil
		}
		return
	}
	r.ensureRPr().Bold = v
}

// Italic returns the italic state: nil means inherit.
func (r *Run) Italic() *bool {
	r.doc.mu.RLock()
	defer r.doc.mu.RUnlock()
	if r.el.RPr == nil {
		return nil
	}
	return r.el.RPr.Italic
}

// SetItalic sets the italic state. Pass nil to clear.
func (r *Run) SetItalic(v *bool) {
	r.doc.mu.Lock()
	defer r.doc.mu.Unlock()
	if v == nil {
		if r.el.RPr != nil {
			r.el.RPr.Italic = nil
		}
		return
	}
	r.ensureRPr().Italic = v
}

// Underline returns the underline state as a bool pointer.
// nil means inherit; true means underlined; false means explicitly not underlined.
func (r *Run) Underline() *bool {
	r.doc.mu.RLock()
	defer r.doc.mu.RUnlock()
	if r.el.RPr == nil || r.el.RPr.Underline == nil {
		return nil
	}
	v := *r.el.RPr.Underline != "none"
	return &v
}

// SetUnderline sets the underline state. nil clears, true sets "single", false sets "none".
func (r *Run) SetUnderline(v *bool) {
	r.doc.mu.Lock()
	defer r.doc.mu.Unlock()
	if v == nil {
		if r.el.RPr != nil {
			r.el.RPr.Underline = nil
		}
		return
	}
	rpr := r.ensureRPr()
	if *v {
		s := "single"
		rpr.Underline = &s
	} else {
		s := "none"
		rpr.Underline = &s
	}
}

// Strikethrough returns the strikethrough state: nil means inherit.
func (r *Run) Strikethrough() *bool {
	r.doc.mu.RLock()
	defer r.doc.mu.RUnlock()
	if r.el.RPr == nil {
		return nil
	}
	return r.el.RPr.Strike
}

// SetStrikethrough sets the strikethrough state. Pass nil to clear.
func (r *Run) SetStrikethrough(v *bool) {
	r.doc.mu.Lock()
	defer r.doc.mu.Unlock()
	if v == nil {
		if r.el.RPr != nil {
			r.el.RPr.Strike = nil
		}
		return
	}
	r.ensureRPr().Strike = v
}

// FontName returns the font name, or nil if not set.
func (r *Run) FontName() *string {
	r.doc.mu.RLock()
	defer r.doc.mu.RUnlock()
	if r.el.RPr == nil {
		return nil
	}
	return r.el.RPr.FontName
}

// SetFontName sets the font name.
func (r *Run) SetFontName(name string) {
	r.doc.mu.Lock()
	defer r.doc.mu.Unlock()
	r.ensureRPr().FontName = &name
}

// FontSize returns the font size in points, or nil if not set.
// OOXML stores font sizes as half-points; this method converts to points.
func (r *Run) FontSize() *float64 {
	r.doc.mu.RLock()
	defer r.doc.mu.RUnlock()
	if r.el.RPr == nil || r.el.RPr.FontSize == nil {
		return nil
	}
	hp, err := strconv.ParseFloat(*r.el.RPr.FontSize, 64)
	if err != nil {
		return nil
	}
	pts := hp / 2
	return &pts
}

// SetFontSize sets the font size in points.
// OOXML stores font sizes as half-points; this method converts from points.
func (r *Run) SetFontSize(points float64) {
	r.doc.mu.Lock()
	defer r.doc.mu.Unlock()
	hp := fmt.Sprintf("%g", points*2)
	r.ensureRPr().FontSize = &hp
}

// FontColor returns the font color, or nil if not set.
func (r *Run) FontColor() *common.Color {
	r.doc.mu.RLock()
	defer r.doc.mu.RUnlock()
	if r.el.RPr == nil || r.el.RPr.Color == nil {
		return nil
	}
	c, err := common.HexColor(*r.el.RPr.Color)
	if err != nil {
		return nil
	}
	return &c
}

// SetFontColor sets the font color.
func (r *Run) SetFontColor(c common.Color) {
	r.doc.mu.Lock()
	defer r.doc.mu.Unlock()
	hex := c.Hex()
	r.ensureRPr().Color = &hex
}

// AddPageBreak appends a page break to the run.
func (r *Run) AddPageBreak() {
	r.doc.mu.Lock()
	defer r.doc.mu.Unlock()
	r.el.Content = append(r.el.Content, wml.RunContent{
		Break: &wml.CT_Break{Type: "page"},
	})
}

// AddLineBreak appends a line break to the run.
func (r *Run) AddLineBreak() {
	r.doc.mu.Lock()
	defer r.doc.mu.Unlock()
	r.el.Content = append(r.el.Content, wml.RunContent{
		Break: &wml.CT_Break{},
	})
}

// AddTab appends a tab character to the run.
func (r *Run) AddTab() {
	r.doc.mu.Lock()
	defer r.doc.mu.Unlock()
	r.el.Content = append(r.el.Content, wml.RunContent{
		Tab: &wml.CT_Tab{},
	})
}
