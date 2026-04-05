package docx

import (
	"encoding/xml"
	"fmt"
	"iter"
	"strings"
	"time"

	"github.com/ieshan/go-ooxml/docx/wml"
)

// RevisionType identifies the kind of tracked change.
type RevisionType int

const (
	// RevisionInsert represents inserted content (<w:ins>).
	RevisionInsert RevisionType = iota
	// RevisionDelete represents deleted content (<w:del>).
	RevisionDelete
)

// Revision represents a tracked change in the document.
type Revision struct {
	doc *Document
	el  *wml.CT_RunTrackChange // the ins or del element
	typ RevisionType
	par *wml.CT_P // owning paragraph
}

// ID returns the revision identifier.
func (r *Revision) ID() int {
	return r.el.ID
}

// Author returns the author who made the change.
func (r *Revision) Author() string {
	return r.el.Author
}

// Date returns the date of the revision. Returns zero time if not set or unparseable.
func (r *Revision) Date() time.Time {
	if r.el.Date == "" {
		return time.Time{}
	}
	t, err := time.Parse(time.RFC3339, r.el.Date)
	if err != nil {
		return time.Time{}
	}
	return t
}

// Type returns the revision type (insert, delete, or format change).
func (r *Revision) Type() RevisionType {
	return r.typ
}

// OriginalText returns the deleted text for delete revisions; empty for inserts.
func (r *Revision) OriginalText() string {
	if r.typ != RevisionDelete {
		return ""
	}
	var sb strings.Builder
	for _, run := range r.el.Runs {
		sb.WriteString(run.DelText())
	}
	return sb.String()
}

// ProposedText returns the inserted text for insert revisions; empty for deletes.
func (r *Revision) ProposedText() string {
	if r.typ != RevisionInsert {
		return ""
	}
	var sb strings.Builder
	for _, run := range r.el.Runs {
		sb.WriteString(run.Text())
	}
	return sb.String()
}

// OriginalMarkdown returns the deleted text for delete revisions formatted as
// Markdown (with bold/italic/code decoration). Returns empty string for
// insert revisions.
func (r *Revision) OriginalMarkdown() string {
	if r.typ != RevisionDelete {
		return ""
	}
	var sb strings.Builder
	for _, run := range r.el.Runs {
		sb.WriteString(runDeleteMarkdown(run))
	}
	return sb.String()
}

// ProposedMarkdown returns the inserted text for insert revisions formatted as
// Markdown (with bold/italic/code decoration). Returns empty string for delete
// revisions.
func (r *Revision) ProposedMarkdown() string {
	if r.typ != RevisionInsert {
		return ""
	}
	var sb strings.Builder
	for _, run := range r.el.Runs {
		sb.WriteString(runToMarkdown(run))
	}
	return sb.String()
}

// revisionConfig holds configuration for creating revisions.
type revisionConfig struct {
	author   string
	date     time.Time
	markdown bool
}

// RevisionOption configures AddRevision.
type RevisionOption func(*revisionConfig)

// WithRevisionAuthor sets the author for a revision.
func WithRevisionAuthor(author string) RevisionOption {
	return func(c *revisionConfig) {
		c.author = author
	}
}

// WithRevisionDate sets the date for a revision.
func WithRevisionDate(date time.Time) RevisionOption {
	return func(c *revisionConfig) {
		c.date = date
	}
}

// WithMarkdownRevision is a stub option for markdown-aware revisions.
func WithMarkdownRevision() RevisionOption {
	return func(c *revisionConfig) {
		c.markdown = true
	}
}

// nextRevisionID scans all paragraphs (including those inside table cells)
// for the highest existing revision ID and returns the next one.
// Caller must hold at least a read lock.
func (d *Document) nextRevisionID() int {
	maxID := 0
	for _, p := range d.collectAllParagraphs() {
		for _, ic := range p.Content {
			if ic.Ins != nil && ic.Ins.ID > maxID {
				maxID = ic.Ins.ID
			}
			if ic.Del != nil && ic.Del.ID > maxID {
				maxID = ic.Del.ID
			}
		}
	}
	return maxID + 1
}

// findParagraphForRun finds the CT_P that contains the given CT_R,
// including paragraphs inside table cells.
// Caller must hold at least a read lock.
func (d *Document) findParagraphForRun(r *wml.CT_R) *wml.CT_P {
	for _, p := range d.collectAllParagraphs() {
		for _, ic := range p.Content {
			if ic.Run == r {
				return p
			}
		}
	}
	return nil
}

// Revisions returns an iterator over all tracked changes in the document,
// including those inside table cells.
func (d *Document) Revisions() iter.Seq[*Revision] {
	return func(yield func(*Revision) bool) {
		d.mu.RLock()
		defer d.mu.RUnlock()
		for _, p := range d.collectAllParagraphs() {
			for _, ic := range p.Content {
				if ic.Ins != nil {
					rev := &Revision{doc: d, el: ic.Ins, typ: RevisionInsert, par: p}
					if !yield(rev) {
						return
					}
				}
				if ic.Del != nil {
					rev := &Revision{doc: d, el: ic.Del, typ: RevisionDelete, par: p}
					if !yield(rev) {
						return
					}
				}
			}
		}
	}
}

// AddRevision wraps the runs from startRun to endRun in a <w:del> and inserts
// a <w:ins> with newText. Returns the Revision for the insertion.
func (d *Document) AddRevision(startRun, endRun *Run, newText string, opts ...RevisionOption) *Revision {
	d.mu.Lock()
	defer d.mu.Unlock()

	cfg := revisionConfig{
		author: d.cfg.Author,
		date:   d.cfg.Date,
	}
	for _, o := range opts {
		o(&cfg)
	}
	if cfg.date.IsZero() {
		cfg.date = time.Now()
	}
	dateStr := cfg.date.UTC().Format(time.RFC3339)

	if cfg.author == "" {
		cfg.author = "Unknown"
	}

	// Find the paragraph containing startRun.
	par := d.findParagraphForRun(startRun.el)
	if par == nil {
		return nil
	}

	// Find indices of startRun and endRun in the paragraph content.
	startIdx, endIdx := -1, -1
	for i, ic := range par.Content {
		if ic.Run == startRun.el {
			startIdx = i
		}
		if ic.Run == endRun.el {
			endIdx = i
		}
	}
	if startIdx < 0 || endIdx < 0 || startIdx > endIdx {
		return nil
	}

	delID := d.nextRevisionID()

	// Collect the original runs and convert their Text to DelText.
	var delRuns []*wml.CT_R
	for i := startIdx; i <= endIdx; i++ {
		r := par.Content[i].Run
		if r == nil {
			continue
		}
		// Convert Text entries to DelText entries.
		for j := range r.Content {
			if r.Content[j].Text != nil {
				r.Content[j].DelText = r.Content[j].Text
				r.Content[j].Text = nil
			}
		}
		delRuns = append(delRuns, r)
	}

	// Build the <w:del> element.
	del := &wml.CT_RunTrackChange{
		XMLName: xml.Name{Space: wml.Ns, Local: "del"},
		ID:      delID,
		Author:  cfg.author,
		Date:    dateStr,
		Runs:    delRuns,
	}

	// Per ECMA-376 17.13.1, each revision mark must have a unique w:id.
	insID := delID + 1

	// Build the <w:ins> element with a new run (or multiple styled runs for markdown).
	var insRuns []*wml.CT_R
	if cfg.markdown {
		// Parse inline markdown and build styled runs.
		inlineRuns := parseInlineMarkdown(newText)
		for _, ir := range inlineRuns {
			r := &wml.CT_R{XMLName: xml.Name{Space: wml.Ns, Local: "r"}}
			r.AddText(ir.text)
			if ir.bold || ir.italic || ir.strike || ir.code {
				r.RPr = &wml.CT_RPr{}
				if ir.bold {
					v := true
					r.RPr.Bold = &v
				}
				if ir.italic {
					v := true
					r.RPr.Italic = &v
				}
				if ir.strike {
					v := true
					r.RPr.Strike = &v
				}
				if ir.code {
					fn := "Courier New"
					r.RPr.FontName = &fn
				}
			}
			insRuns = append(insRuns, r)
		}
		if len(insRuns) == 0 {
			r := &wml.CT_R{XMLName: xml.Name{Space: wml.Ns, Local: "r"}}
			r.AddText(newText)
			insRuns = []*wml.CT_R{r}
		}
	} else {
		insRun := &wml.CT_R{XMLName: xml.Name{Space: wml.Ns, Local: "r"}}
		insRun.AddText(newText)
		insRuns = []*wml.CT_R{insRun}
	}
	ins := &wml.CT_RunTrackChange{
		XMLName: xml.Name{Space: wml.Ns, Local: "ins"},
		ID:      insID,
		Author:  cfg.author,
		Date:    dateStr,
		Runs:    insRuns,
	}

	// Replace the original run entries [startIdx..endIdx] with del + ins.
	newContent := make([]wml.InlineContent, 0, len(par.Content)-endIdx+startIdx+1)
	newContent = append(newContent, par.Content[:startIdx]...)
	newContent = append(newContent, wml.InlineContent{Del: del})
	newContent = append(newContent, wml.InlineContent{Ins: ins})
	newContent = append(newContent, par.Content[endIdx+1:]...)
	par.Content = newContent

	return &Revision{doc: d, el: ins, typ: RevisionInsert, par: par}
}

// findRevisionPartner locates the index and adjacent partner revision for a given
// track change element. A partner is the adjacent ins/del pair (del immediately
// before ins). Returns (index of el, partner element, partner index, found).
// Caller must hold lock.
func findRevisionPartner(par *wml.CT_P, el *wml.CT_RunTrackChange) (elIdx int, partner *wml.CT_RunTrackChange, partnerIdx int, found bool) {
	elIdx = -1
	partnerIdx = -1
	for i, ic := range par.Content {
		if ic.Ins == el || ic.Del == el {
			elIdx = i
			break
		}
	}
	if elIdx < 0 {
		return
	}
	found = true

	// Look for the adjacent partner: del+ins pairs are always adjacent.
	if par.Content[elIdx].Ins == el {
		// This is an ins; partner del should be at elIdx-1.
		if elIdx > 0 && par.Content[elIdx-1].Del != nil {
			partner = par.Content[elIdx-1].Del
			partnerIdx = elIdx - 1
		}
	} else {
		// This is a del; partner ins should be at elIdx+1.
		if elIdx+1 < len(par.Content) && par.Content[elIdx+1].Ins != nil {
			partner = par.Content[elIdx+1].Ins
			partnerIdx = elIdx + 1
		}
	}
	return
}

// AcceptRevision accepts a tracked change, making it permanent.
// For inserts: unwraps the insertion runs into the paragraph.
// For deletes: removes the deletion and its content entirely.
func (d *Document) AcceptRevision(rev *Revision) error {
	d.mu.Lock()
	defer d.mu.Unlock()

	par := rev.par
	if par == nil {
		return fmt.Errorf("%w: revision has no parent paragraph", ErrNotFound)
	}

	elIdx, _, partnerIdx, found := findRevisionPartner(par, rev.el)
	if !found {
		return fmt.Errorf("%w: revision element not found in paragraph", ErrNotFound)
	}

	// Build new content slice.
	// We need to handle both the ins and its paired del.
	// Accept means: keep ins runs as normal, remove del entirely.
	var newContent []wml.InlineContent

	// Collect indices to skip.
	skip := map[int]bool{elIdx: true}
	if partnerIdx >= 0 {
		skip[partnerIdx] = true
	}

	for i, ic := range par.Content {
		if skip[i] {
			continue
		}
		newContent = append(newContent, ic)
	}

	// Now insert the accepted content at the right position.
	// For an insert revision: unwrap its runs as normal runs.
	// For a delete revision: accept means remove it (already skipped).
	insertAt := elIdx
	if partnerIdx >= 0 && partnerIdx < insertAt {
		insertAt-- // adjust for removed partner before this index
	}

	var unwrappedContent []wml.InlineContent
	if rev.typ == RevisionInsert {
		// Unwrap ins runs as normal content.
		for _, r := range rev.el.Runs {
			unwrappedContent = append(unwrappedContent, wml.InlineContent{Run: r})
		}
	}
	// For RevisionDelete: accepting a delete means the deletion is applied (content removed).
	// The del runs are discarded (already skipped above).

	// Insert unwrapped content at the position.
	if len(unwrappedContent) > 0 {
		// Adjust insertAt for skipped items before it.
		adjustedIdx := 0
		for i := range par.Content {
			if i == elIdx {
				break
			}
			if !skip[i] {
				adjustedIdx++
			}
		}

		result := make([]wml.InlineContent, 0, len(newContent)+len(unwrappedContent))
		result = append(result, newContent[:adjustedIdx]...)
		result = append(result, unwrappedContent...)
		result = append(result, newContent[adjustedIdx:]...)
		newContent = result
	}

	par.Content = newContent
	return nil
}

// RejectRevision rejects a tracked change, restoring original text.
// For inserts: removes the insertion entirely.
// For deletes: unwraps the deletion runs back as normal content (delText -> text).
func (d *Document) RejectRevision(rev *Revision) error {
	d.mu.Lock()
	defer d.mu.Unlock()

	par := rev.par
	if par == nil {
		return fmt.Errorf("%w: revision has no parent paragraph", ErrNotFound)
	}

	elIdx, partner, partnerIdx, found := findRevisionPartner(par, rev.el)
	if !found {
		return fmt.Errorf("%w: revision element not found in paragraph", ErrNotFound)
	}

	// Reject means: for ins, remove it; for del, unwrap it back to normal.
	// We also handle the partner: reject ins -> also need to unwrap del partner.
	skip := map[int]bool{elIdx: true}
	if partnerIdx >= 0 {
		skip[partnerIdx] = true
	}

	var newContent []wml.InlineContent
	for i, ic := range par.Content {
		if skip[i] {
			continue
		}
		newContent = append(newContent, ic)
	}

	// Determine what to unwrap.
	// If rejecting an insert: remove ins, unwrap del (partner).
	// If rejecting a delete: remove del, unwrap... no, reject delete means keep original,
	// so unwrap the del runs back as normal.

	var unwrappedContent []wml.InlineContent
	var unwrapSource *wml.CT_RunTrackChange
	var unwrapAt int

	if rev.typ == RevisionInsert {
		// Rejecting insert: restore del (partner) runs as normal.
		if partner != nil {
			unwrapSource = partner
			unwrapAt = partnerIdx
		}
	} else {
		// Rejecting delete: restore del runs as normal.
		unwrapSource = rev.el
		unwrapAt = elIdx
	}

	if unwrapSource != nil {
		for _, r := range unwrapSource.Runs {
			// Convert DelText back to Text.
			for j := range r.Content {
				if r.Content[j].DelText != nil {
					r.Content[j].Text = r.Content[j].DelText
					r.Content[j].DelText = nil
				}
			}
			unwrappedContent = append(unwrappedContent, wml.InlineContent{Run: r})
		}
	}

	if len(unwrappedContent) > 0 {
		// Calculate adjusted index for where to insert unwrapped content.
		adjustedIdx := 0
		for i := range par.Content {
			if i == unwrapAt {
				break
			}
			if !skip[i] {
				adjustedIdx++
			}
		}

		result := make([]wml.InlineContent, 0, len(newContent)+len(unwrappedContent))
		result = append(result, newContent[:adjustedIdx]...)
		result = append(result, unwrappedContent...)
		result = append(result, newContent[adjustedIdx:]...)
		newContent = result
	}

	par.Content = newContent
	return nil
}

// RemoveRevision removes both ins and del parts of a revision entirely,
// without accepting or rejecting.
func (d *Document) RemoveRevision(rev *Revision) error {
	d.mu.Lock()
	defer d.mu.Unlock()

	par := rev.par
	if par == nil {
		return fmt.Errorf("%w: revision has no parent paragraph", ErrNotFound)
	}

	elIdx, _, partnerIdx, found := findRevisionPartner(par, rev.el)
	if !found {
		return fmt.Errorf("%w: revision element not found in paragraph", ErrNotFound)
	}

	skip := map[int]bool{elIdx: true}
	if partnerIdx >= 0 {
		skip[partnerIdx] = true
	}

	var newContent []wml.InlineContent
	for i, ic := range par.Content {
		if skip[i] {
			continue
		}
		newContent = append(newContent, ic)
	}
	par.Content = newContent
	return nil
}

// OverrideRevision replaces the proposed text of an existing revision with new text.
// It removes the old revision and creates a new one.
func (d *Document) OverrideRevision(rev *Revision, text string, opts ...RevisionOption) *Revision {
	d.mu.Lock()
	defer d.mu.Unlock()

	par := rev.par
	if par == nil {
		return nil
	}

	cfg := revisionConfig{
		author: d.cfg.Author,
		date:   d.cfg.Date,
	}
	for _, o := range opts {
		o(&cfg)
	}
	if cfg.author == "" {
		cfg.author = "Unknown"
	}
	if cfg.date.IsZero() {
		cfg.date = time.Now()
	}
	dateStr := cfg.date.UTC().Format(time.RFC3339)

	// Find the existing ins and adjacent del using partner search.
	insIdx, existingDel, delIdx, _ := findRevisionPartner(par, rev.el)
	if insIdx < 0 {
		return nil
	}

	newDelID := d.nextRevisionID()
	newInsID := newDelID + 1

	// Build new ins with updated text.
	insRun := &wml.CT_R{XMLName: xml.Name{Space: wml.Ns, Local: "r"}}
	insRun.AddText(text)
	ins := &wml.CT_RunTrackChange{
		XMLName: xml.Name{Space: wml.Ns, Local: "ins"},
		ID:      newInsID,
		Author:  cfg.author,
		Date:    dateStr,
		Runs:    []*wml.CT_R{insRun},
	}

	// Update del with new IDs.
	if existingDel != nil {
		existingDel.ID = newDelID
		existingDel.Author = cfg.author
		existingDel.Date = dateStr
	}

	// Replace ins in the content; keep del.
	var newContent []wml.InlineContent
	for i, ic := range par.Content {
		if i == insIdx {
			newContent = append(newContent, wml.InlineContent{Ins: ins})
			continue
		}
		if i == delIdx && existingDel != nil {
			newContent = append(newContent, wml.InlineContent{Del: existingDel})
			continue
		}
		newContent = append(newContent, ic)
	}
	par.Content = newContent

	return &Revision{doc: d, el: ins, typ: RevisionInsert, par: par}
}

// RevisionsForText returns an iterator over revisions whose proposed (insert) text
// contains the given substring.
func (d *Document) RevisionsForText(text string) iter.Seq[*Revision] {
	return func(yield func(*Revision) bool) {
		for rev := range d.Revisions() {
			if rev.typ == RevisionInsert && strings.Contains(rev.ProposedText(), text) {
				if !yield(rev) {
					return
				}
			}
		}
	}
}
