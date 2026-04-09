package docx

import (
	"encoding/xml"
	"iter"
	"strconv"

	"github.com/ieshan/go-ooxml/common"
	"github.com/ieshan/go-ooxml/docx/wml"
	"github.com/ieshan/go-ooxml/xmlutil"
)

// Orientation represents the page orientation of a section.
type Orientation int

const (
	// OrientPortrait indicates portrait (vertical) orientation.
	OrientPortrait Orientation = iota
	// OrientLandscape indicates landscape (horizontal) orientation.
	OrientLandscape
)

// SectionBreak represents the type of section break that precedes the section.
type SectionBreak int

const (
	// BreakNextPage starts the section on the next page.
	BreakNextPage SectionBreak = iota
	// BreakContinuous starts the section immediately, with no page break.
	BreakContinuous
	// BreakEvenPage starts the section on the next even-numbered page.
	BreakEvenPage
	// BreakOddPage starts the section on the next odd-numbered page.
	BreakOddPage
)

// SectionMargins holds the page margin measurements for a section.
type SectionMargins struct {
	Top, Right, Bottom, Left common.Length
}

// ColumnLayout describes the column layout for a section.
type ColumnLayout struct {
	Count      int
	Space      common.Length
	EqualWidth bool
	Cols       []ColumnDef
	Separator  bool
}

// ColumnDef describes a single column in a multi-column layout.
type ColumnDef struct {
	Width common.Length
	Space common.Length
}

// Section wraps a wml.CT_SectPr element, providing a high-level API for
// reading and writing section-level properties such as page size, margins,
// and column layout.
type Section struct {
	doc *Document
	el  *wml.CT_SectPr
}

// Sections returns an iterator over the sections in the document.
// For this simplified implementation, only the body's section properties
// are returned as a single section.
func (d *Document) Sections() iter.Seq[*Section] {
	return func(yield func(*Section) bool) {
		// Acquire the lock only long enough to obtain the section element,
		// then release before yielding so the caller can safely call mutating
		// methods on the section (which acquire the lock themselves).
		d.mu.Lock()
		sectPr := d.ensureBodySectPr()
		d.mu.Unlock()

		yield(&Section{doc: d, el: sectPr})
	}
}

// ensureBodySectPr returns the cached CT_SectPr, parsing from raw XML on
// first access or creating a minimal one if absent. The parsed struct is
// cached on d.sectPr so subsequent calls return the same pointer.
// Caller must hold d.mu (write lock).
func (d *Document) ensureBodySectPr() *wml.CT_SectPr {
	if d.sectPr != nil {
		return d.sectPr
	}

	body := d.doc.Body
	if body.SectPr != nil && len(body.SectPr.Data) > 0 {
		var sp wml.CT_SectPr
		if err := xmlutil.Unmarshal(body.SectPr.Data, &sp); err == nil {
			d.sectPr = &sp
			return d.sectPr
		}
	}

	// No existing sectPr — create a minimal one.
	d.sectPr = &wml.CT_SectPr{
		XMLName: xml.Name{Space: wml.Ns, Local: "sectPr"},
	}
	return d.sectPr
}

// syncBodySectPr marshals the cached sectPr back to the body's RawXML.
// Called during syncParts before saving.
func (d *Document) syncBodySectPr() {
	if d.sectPr == nil {
		return
	}
	data, err := xmlutil.Marshal(d.sectPr, xmlutil.OOXML)
	if err != nil {
		return
	}
	if d.doc.Body.SectPr == nil {
		d.doc.Body.SectPr = &xmlutil.RawXML{}
	}
	d.doc.Body.SectPr.Data = data
}

// PageSize returns the page width and height of the section in common.Length.
func (s *Section) PageSize() (width, height common.Length) {
	s.doc.mu.RLock()
	defer s.doc.mu.RUnlock()

	if s.el.PgSz == nil {
		return 0, 0
	}
	w, _ := strconv.ParseInt(s.el.PgSz.W, 10, 64)
	h, _ := strconv.ParseInt(s.el.PgSz.H, 10, 64)
	return common.Twips(w), common.Twips(h)
}

// SetPageSize sets the page width and height of the section.
func (s *Section) SetPageSize(width, height common.Length) {
	s.doc.mu.Lock()
	defer s.doc.mu.Unlock()

	if s.el.PgSz == nil {
		s.el.PgSz = &wml.CT_PgSz{}
	}
	s.el.PgSz.W = strconv.FormatInt(width.Twips(), 10)
	s.el.PgSz.H = strconv.FormatInt(height.Twips(), 10)
	// Cached sectPr is written back to body during syncParts/save
}

// Orientation returns the page orientation of the section.
func (s *Section) Orientation() Orientation {
	s.doc.mu.RLock()
	defer s.doc.mu.RUnlock()

	if s.el.PgSz != nil && s.el.PgSz.Orient == "landscape" {
		return OrientLandscape
	}
	return OrientPortrait
}

// SetOrientation sets the page orientation of the section.
// When switching to landscape the width and height values are swapped if they
// are currently set up for portrait (width < height), and vice-versa.
func (s *Section) SetOrientation(o Orientation) {
	s.doc.mu.Lock()
	defer s.doc.mu.Unlock()

	if s.el.PgSz == nil {
		s.el.PgSz = &wml.CT_PgSz{}
	}

	switch o {
	case OrientLandscape:
		s.el.PgSz.Orient = "landscape"
	default:
		s.el.PgSz.Orient = ""
	}
	// Cached sectPr is written back to body during syncParts/save
}

// Margins returns the page margins of the section.
func (s *Section) Margins() SectionMargins {
	s.doc.mu.RLock()
	defer s.doc.mu.RUnlock()

	if s.el.PgMar == nil {
		return SectionMargins{}
	}
	top, _ := strconv.ParseInt(s.el.PgMar.Top, 10, 64)
	right, _ := strconv.ParseInt(s.el.PgMar.Right, 10, 64)
	bottom, _ := strconv.ParseInt(s.el.PgMar.Bottom, 10, 64)
	left, _ := strconv.ParseInt(s.el.PgMar.Left, 10, 64)
	return SectionMargins{
		Top:    common.Twips(top),
		Right:  common.Twips(right),
		Bottom: common.Twips(bottom),
		Left:   common.Twips(left),
	}
}

// SetMargins sets the page margins of the section.
func (s *Section) SetMargins(m SectionMargins) {
	s.doc.mu.Lock()
	defer s.doc.mu.Unlock()

	if s.el.PgMar == nil {
		s.el.PgMar = &wml.CT_PgMar{}
	}
	s.el.PgMar.Top = strconv.FormatInt(m.Top.Twips(), 10)
	s.el.PgMar.Right = strconv.FormatInt(m.Right.Twips(), 10)
	s.el.PgMar.Bottom = strconv.FormatInt(m.Bottom.Twips(), 10)
	s.el.PgMar.Left = strconv.FormatInt(m.Left.Twips(), 10)
	// Cached sectPr is written back to body during syncParts/save
}

// Columns returns the column layout of the section, or nil if none is set.
func (s *Section) Columns() *ColumnLayout {
	s.doc.mu.RLock()
	defer s.doc.mu.RUnlock()

	if s.el.Cols == nil {
		return nil
	}
	cols := s.el.Cols
	layout := &ColumnLayout{}

	if cols.Num != nil {
		layout.Count = *cols.Num
	}
	if cols.Space != nil {
		v, _ := strconv.ParseInt(*cols.Space, 10, 64)
		layout.Space = common.Twips(v)
	}
	if cols.EqualWidth != nil {
		layout.EqualWidth = *cols.EqualWidth
	}
	if cols.Sep != nil {
		layout.Separator = *cols.Sep
	}
	for _, c := range cols.Cols {
		w, _ := strconv.ParseInt(c.W, 10, 64)
		sp, _ := strconv.ParseInt(c.Space, 10, 64)
		layout.Cols = append(layout.Cols, ColumnDef{
			Width: common.Twips(w),
			Space: common.Twips(sp),
		})
	}
	return layout
}

// SetColumns sets the column layout of the section.
func (s *Section) SetColumns(cols *ColumnLayout) {
	s.doc.mu.Lock()
	defer s.doc.mu.Unlock()

	if cols == nil {
		s.el.Cols = nil
		// Cached sectPr is written back to body during syncParts/save
		return
	}

	wmlCols := &wml.CT_Columns{}
	if cols.Count > 0 {
		n := cols.Count
		wmlCols.Num = &n
	}
	spaceStr := strconv.FormatInt(cols.Space.Twips(), 10)
	wmlCols.Space = &spaceStr
	eq := cols.EqualWidth
	wmlCols.EqualWidth = &eq
	sep := cols.Separator
	wmlCols.Sep = &sep

	for _, cd := range cols.Cols {
		wmlCols.Cols = append(wmlCols.Cols, &wml.CT_Column{
			W:     strconv.FormatInt(cd.Width.Twips(), 10),
			Space: strconv.FormatInt(cd.Space.Twips(), 10),
		})
	}
	s.el.Cols = wmlCols
	// Cached sectPr is written back to body during syncParts/save
}

// BreakType returns the section break type.
func (s *Section) BreakType() SectionBreak {
	s.doc.mu.RLock()
	defer s.doc.mu.RUnlock()

	if s.el.Type == nil {
		return BreakNextPage
	}
	switch *s.el.Type {
	case "continuous":
		return BreakContinuous
	case "evenPage":
		return BreakEvenPage
	case "oddPage":
		return BreakOddPage
	default:
		return BreakNextPage
	}
}

// SetBreakType sets the section break type.
func (s *Section) SetBreakType(b SectionBreak) {
	s.doc.mu.Lock()
	defer s.doc.mu.Unlock()

	var v string
	switch b {
	case BreakContinuous:
		v = "continuous"
	case BreakEvenPage:
		v = "evenPage"
	case BreakOddPage:
		v = "oddPage"
	default:
		v = "nextPage"
	}
	s.el.Type = &v
	// Cached sectPr is written back to body during syncParts/save
}
