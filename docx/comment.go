package docx

import (
	"encoding/xml"
	"fmt"
	"iter"
	"strings"
	"time"

	"github.com/ieshan/go-ooxml/docx/wml"
	"github.com/ieshan/go-ooxml/opc"
)

// Comment wraps a wml.CT_Comment.
type Comment struct {
	doc *Document
	el  *wml.CT_Comment
}

// ID returns the comment's unique identifier.
func (c *Comment) ID() int {
	return c.el.ID
}

// Author returns the comment author.
func (c *Comment) Author() string {
	return c.el.Author
}

// Date returns the comment date. Returns zero time if not set or unparseable.
func (c *Comment) Date() time.Time {
	if c.el.Date == "" {
		return time.Time{}
	}
	t, err := time.Parse(time.RFC3339, c.el.Date)
	if err != nil {
		return time.Time{}
	}
	return t
}

// Initials returns the comment author's initials.
func (c *Comment) Initials() string {
	return c.el.Initials
}

// Text returns the plain text of the comment body.
func (c *Comment) Text() string {
	c.doc.mu.RLock()
	defer c.doc.mu.RUnlock()
	var sb strings.Builder
	first := true
	for _, bc := range c.el.Content {
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

// Markdown returns the comment body formatted as Markdown.
func (c *Comment) Markdown() string {
	c.doc.mu.RLock()
	defer c.doc.mu.RUnlock()
	var parts []string
	for _, bc := range c.el.Content {
		if bc.Paragraph != nil {
			parts = append(parts, paragraphToMarkdown(bc.Paragraph, nil))
		}
	}
	return strings.Join(parts, "\n\n")
}

// Paragraphs returns an iterator over paragraphs in the comment body.
func (c *Comment) Paragraphs() iter.Seq[*Paragraph] {
	return func(yield func(*Paragraph) bool) {
		c.doc.mu.RLock()
		defer c.doc.mu.RUnlock()
		for _, bc := range c.el.Content {
			if bc.Paragraph != nil {
				if !yield(&Paragraph{doc: c.doc, el: bc.Paragraph}) {
					return
				}
			}
		}
	}
}

// commentConfig holds configuration for AddComment.
type commentConfig struct {
	author   string
	date     time.Time
	initials string
	markdown bool
}

// CommentOption configures AddComment.
type CommentOption func(*commentConfig)

// WithAuthor sets the comment author.
func WithAuthor(author string) CommentOption {
	return func(c *commentConfig) {
		c.author = author
	}
}

// WithDate sets the comment date.
func WithDate(date time.Time) CommentOption {
	return func(c *commentConfig) {
		c.date = date
	}
}

// WithInitials sets the comment author's initials.
func WithInitials(initials string) CommentOption {
	return func(c *commentConfig) {
		c.initials = initials
	}
}

// WithMarkdown is a stub option for future markdown support.
func WithMarkdown() CommentOption {
	return func(c *commentConfig) {
		c.markdown = true
	}
}

// Comments returns an iterator over all comments in the document.
func (d *Document) Comments() iter.Seq[*Comment] {
	return func(yield func(*Comment) bool) {
		d.mu.RLock()
		defer d.mu.RUnlock()
		if d.comments == nil {
			return
		}
		for _, c := range d.comments.Comments {
			if !yield(&Comment{doc: d, el: c}) {
				return
			}
		}
	}
}

// AddComment adds a comment anchored between startRun and endRun with the given text.
// Options can override the author, date, and initials.
func (d *Document) AddComment(startRun, endRun *Run, text string, opts ...CommentOption) *Comment {
	d.mu.Lock()
	defer d.mu.Unlock()

	// Build config from defaults + options.
	cfg := commentConfig{
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

	// Ensure comments container exists.
	if d.comments == nil {
		d.comments = &wml.CT_Comments{
			XMLName: xml.Name{Space: wml.Ns, Local: "comments"},
		}
	}

	// Generate next comment ID.
	nextID := 0
	for _, c := range d.comments.Comments {
		if c.ID >= nextID {
			nextID = c.ID + 1
		}
	}

	// Create the CT_Comment.
	comment := &wml.CT_Comment{
		XMLName:  xml.Name{Space: wml.Ns, Local: "comment"},
		ID:       nextID,
		Author:   cfg.author,
		Date:     cfg.date.Format(time.RFC3339),
		Initials: cfg.initials,
	}

	// Add a paragraph with the comment text.
	// Per ECMA-376 17.13.4.1, the first paragraph must contain an annotationRef run.
	p := &wml.CT_P{XMLName: xml.Name{Space: wml.Ns, Local: "p"}}
	annotRefRun := &wml.CT_R{XMLName: xml.Name{Space: wml.Ns, Local: "r"}}
	commentRefStyle := "CommentReference"
	annotRefRun.RPr = &wml.CT_RPr{RunStyle: &commentRefStyle}
	annotRefRun.Content = append(annotRefRun.Content, wml.RunContent{AnnotationRef: &wml.CT_AnnotationRef{}})
	p.Content = append(p.Content, wml.InlineContent{Run: annotRefRun})
	if cfg.markdown {
		// Parse the text as inline markdown and create styled runs.
		// Must use the WML-level helper to avoid re-entrancy on the document lock.
		applyInlineMarkdownWML(p, text)
	} else {
		p.AddRun(text)
	}
	comment.Content = append(comment.Content, wml.BlockLevelContent{Paragraph: p})

	d.comments.Comments = append(d.comments.Comments, comment)

	// Insert comment range markers and reference in the document body.
	d.insertCommentMarkers(startRun.el, endRun.el, nextID)

	// Ensure the comments relationship and content type exist.
	d.ensureCommentsRelationship()

	return &Comment{doc: d, el: comment}
}

// RemoveComment removes a comment from the document, including its range markers.
func (d *Document) RemoveComment(comment *Comment) error {
	d.mu.Lock()
	defer d.mu.Unlock()

	if d.comments == nil {
		return fmt.Errorf("docx: %w: no comments in document", ErrNotFound)
	}

	// Remove from comments list.
	found := false
	commentID := comment.el.ID
	for i, c := range d.comments.Comments {
		if c == comment.el {
			d.comments.Comments = append(d.comments.Comments[:i], d.comments.Comments[i+1:]...)
			found = true
			break
		}
	}
	if !found {
		return fmt.Errorf("docx: %w: comment ID %d", ErrNotFound, commentID)
	}

	// Remove comment range markers and references from body paragraphs.
	d.removeCommentMarkers(commentID)

	return nil
}

// CommentsForText finds comments whose range markers span text containing the search string.
// Handles both single-paragraph and cross-paragraph comment ranges.
func (d *Document) CommentsForText(text string) iter.Seq[*Comment] {
	return func(yield func(*Comment) bool) {
		d.mu.RLock()
		defer d.mu.RUnlock()

		if d.comments == nil || d.doc.Body == nil {
			return
		}

		// Build a map of comment ID -> *CT_Comment for quick lookup.
		commentMap := make(map[int]*wml.CT_Comment, len(d.comments.Comments))
		for _, c := range d.comments.Comments {
			commentMap[c.ID] = c
		}

		// Collect all paragraphs (including those inside table cells).
		allParas := d.collectAllParagraphs()

		// First pass: record which paragraph index each comment range starts in.
		// Map from comment ID -> paragraph index of commentRangeStart.
		startParaIdx := make(map[int]int)
		for pi, p := range allParas {
			for _, ic := range p.Content {
				if ic.CommentRangeStart != nil {
					startParaIdx[ic.CommentRangeStart.ID] = pi
				}
			}
		}

		// Second pass: when we find commentRangeEnd, collect text from start para to end para.
		yielded := make(map[int]bool)
		for pi, p := range allParas {
			for _, ic := range p.Content {
				if ic.CommentRangeEnd == nil {
					continue
				}
				cID := ic.CommentRangeEnd.ID
				spi, ok := startParaIdx[cID]
				if !ok || yielded[cID] {
					continue
				}
				// Collect text across all paragraphs from spi to pi.
				var sb strings.Builder
				for j := spi; j <= pi; j++ {
					for _, jc := range allParas[j].Content {
						if jc.Run != nil {
							sb.WriteString(jc.Run.Text())
						}
						if jc.Ins != nil {
							for _, r := range jc.Ins.Runs {
								sb.WriteString(r.Text())
							}
						}
						if jc.Hyperlink != nil {
							for _, r := range jc.Hyperlink.Runs {
								sb.WriteString(r.Text())
							}
						}
					}
				}
				if strings.Contains(sb.String(), text) {
					if c, ok := commentMap[cID]; ok {
						yielded[cID] = true
						if !yield(&Comment{doc: d, el: c}) {
							return
						}
					}
				}
			}
		}
	}
}

// insertCommentMarkers inserts commentRangeStart, commentRangeEnd, and a commentReference run
// into the paragraph(s) containing the start and end runs. Searches all paragraphs including
// those inside table cells.
func (d *Document) insertCommentMarkers(startRun, endRun *wml.CT_R, commentID int) {
	allParas := d.collectAllParagraphs()

	// Find the paragraph and index for startRun and endRun.
	var startPara, endPara *wml.CT_P
	startIdx, endIdx := -1, -1
	for _, p := range allParas {
		for i, ic := range p.Content {
			if ic.Run == startRun {
				startPara = p
				startIdx = i
			}
			if ic.Run == endRun {
				endPara = p
				endIdx = i
			}
		}
	}
	if startPara == nil {
		return
	}

	rangeStart := wml.InlineContent{
		CommentRangeStart: &wml.CT_MarkupRange{
			XMLName: xml.Name{Space: wml.Ns, Local: "commentRangeStart"},
			ID:      commentID,
		},
	}

	rangeEnd := wml.InlineContent{
		CommentRangeEnd: &wml.CT_MarkupRange{
			XMLName: xml.Name{Space: wml.Ns, Local: "commentRangeEnd"},
			ID:      commentID,
		},
	}

	// Build a comment reference run with CommentReference character style.
	refRun := &wml.CT_R{XMLName: xml.Name{Space: wml.Ns, Local: "r"}}
	crStyle := "CommentReference"
	refRun.RPr = &wml.CT_RPr{RunStyle: &crStyle}
	refRun.Content = append(refRun.Content, wml.RunContent{
		CommentReference: &wml.CT_CommentReference{ID: commentID},
	})
	refRunIC := wml.InlineContent{Run: refRun}

	if startPara == endPara && endIdx >= 0 {
		// Same paragraph: insert rangeStart before startIdx, rangeEnd+refRun after endIdx.
		p := startPara
		newContent := make([]wml.InlineContent, 0, len(p.Content)+3)
		for i, ic := range p.Content {
			if i == startIdx {
				newContent = append(newContent, rangeStart)
			}
			newContent = append(newContent, ic)
			if i == endIdx {
				newContent = append(newContent, rangeEnd, refRunIC)
			}
		}
		p.Content = newContent
		return
	}

	// Cross-paragraph: insert rangeStart in startPara, rangeEnd+refRun in endPara.
	newContent := make([]wml.InlineContent, 0, len(startPara.Content)+1)
	for i, ic := range startPara.Content {
		if i == startIdx {
			newContent = append(newContent, rangeStart)
		}
		newContent = append(newContent, ic)
	}
	startPara.Content = newContent

	if endPara != nil {
		newContent2 := make([]wml.InlineContent, 0, len(endPara.Content)+2)
		for j, ic2 := range endPara.Content {
			newContent2 = append(newContent2, ic2)
			if j == endIdx {
				newContent2 = append(newContent2, rangeEnd, refRunIC)
			}
		}
		endPara.Content = newContent2
	}
}

// removeCommentMarkers removes all commentRangeStart, commentRangeEnd, and commentReference
// elements with the given comment ID from all paragraphs, including those inside table cells.
func (d *Document) removeCommentMarkers(commentID int) {
	for _, p := range d.collectAllParagraphs() {
		filtered := make([]wml.InlineContent, 0, len(p.Content))
		for _, ic := range p.Content {
			// Skip commentRangeStart with matching ID.
			if ic.CommentRangeStart != nil && ic.CommentRangeStart.ID == commentID {
				continue
			}
			// Skip commentRangeEnd with matching ID.
			if ic.CommentRangeEnd != nil && ic.CommentRangeEnd.ID == commentID {
				continue
			}
			// Skip runs that only contain a commentReference with matching ID.
			if ic.Run != nil && isCommentRefRun(ic.Run, commentID) {
				continue
			}
			filtered = append(filtered, ic)
		}
		p.Content = filtered
	}
}

// isCommentRefRun returns true if the run contains only a single commentReference
// with the given ID (i.e. it's a pure comment reference marker run).
func isCommentRefRun(r *wml.CT_R, commentID int) bool {
	if len(r.Content) != 1 {
		return false
	}
	cr := r.Content[0].CommentReference
	return cr != nil && cr.ID == commentID
}

// ensureCommentsRelationship ensures that the document part has a relationship
// to comments.xml and the content type override is registered.
func (d *Document) ensureCommentsRelationship() {
	// Check if the document part already has a comments relationship.
	docPart, ok := d.pkg.Parts[d.docPath]
	if !ok {
		return
	}

	if opc.FindRelByType(docPart.Rels, opc.RelComments) != nil {
		return
	}

	// Add part-level relationship.
	rel := opc.Relationship{
		ID:     opc.NextRelID(docPart.Rels),
		Type:   opc.RelComments,
		Target: "comments.xml",
	}
	docPart.Rels = append(docPart.Rels, rel)

	// Ensure content type override.
	commentsPath := resolveRelPath(d.docPath, "comments.xml")
	d.pkg.ContentTypes.Overrides["/"+commentsPath] = ctComments
}
