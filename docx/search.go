package docx

import (
	"regexp"
	"strings"

	"github.com/ieshan/go-ooxml/docx/wml"
)

// SearchResult represents a text match in the document.
type SearchResult struct {
	Paragraph *Paragraph
	Runs      []*Run
	Start     int // Character offset within first run where the match begins
	End       int // Character offset within last run where the match ends
}

// Text returns the matched text by concatenating the relevant portions of all matching runs.
func (sr *SearchResult) Text() string {
	if len(sr.Runs) == 0 {
		return ""
	}
	if len(sr.Runs) == 1 {
		t := sr.Runs[0].el.Text()
		if sr.End > len(t) {
			sr.End = len(t)
		}
		return t[sr.Start:sr.End]
	}
	var sb strings.Builder
	for i, r := range sr.Runs {
		t := r.el.Text()
		switch {
		case i == 0:
			sb.WriteString(t[sr.Start:])
		case i == len(sr.Runs)-1:
			if sr.End > len(t) {
				sr.End = len(t)
			}
			sb.WriteString(t[:sr.End])
		default:
			sb.WriteString(t)
		}
	}
	return sb.String()
}

// Markdown returns the matched text formatted as Markdown, preserving
// bold, italic, strikethrough, and code formatting from the matched runs.
func (sr *SearchResult) Markdown() string {
	if len(sr.Runs) == 0 {
		return ""
	}
	var sb strings.Builder
	for i, r := range sr.Runs {
		text := r.el.Text()
		switch {
		case i == 0 && i == len(sr.Runs)-1:
			if sr.End > len(text) {
				sr.End = len(text)
			}
			text = text[sr.Start:sr.End]
		case i == 0:
			text = text[sr.Start:]
		case i == len(sr.Runs)-1:
			if sr.End > len(text) {
				sr.End = len(text)
			}
			text = text[:sr.End]
		}
		if text == "" {
			continue
		}
		// Build a temporary CT_R with the sliced text to reuse runToMarkdown.
		tmp := *r.el
		tmp.Content = nil
		tmp.AddText(text)
		sb.WriteString(runToMarkdown(&tmp))
	}
	return sb.String()
}

// SearchQuery provides a fluent API for search-and-operate.
type SearchQuery struct {
	doc           *Document
	text          string
	caseSensitive bool
	styleName     string
	first         bool
	nth           int // -1 = all, >=0 = specific index
	useRegex      bool
	pattern       *regexp.Regexp
}

// InStyle filters results to paragraphs with the given style name.
func (q *SearchQuery) InStyle(styleName string) *SearchQuery {
	q.styleName = styleName
	return q
}

// CaseSensitive sets whether the search is case-sensitive (default: false).
func (q *SearchQuery) CaseSensitive(v bool) *SearchQuery {
	q.caseSensitive = v
	return q
}

// First restricts results to only the first match.
func (q *SearchQuery) First() *SearchQuery {
	q.first = true
	q.nth = 0
	return q
}

// Nth restricts results to only the nth match (0-based index).
func (q *SearchQuery) Nth(n int) *SearchQuery {
	q.nth = n
	q.first = false
	return q
}

// Results executes the query and returns all matching SearchResults.
func (q *SearchQuery) Results() []SearchResult {
	q.doc.mu.RLock()
	defer q.doc.mu.RUnlock()
	return q.execute()
}

// execute performs the search without acquiring any locks.
// Caller must hold at least RLock.
func (q *SearchQuery) execute() []SearchResult {
	var results []SearchResult

	matchIndex := 0
	for _, para := range q.doc.collectAllParagraphs() {
		p := &Paragraph{doc: q.doc, el: para}

		// Filter by style if requested.
		if q.styleName != "" {
			var style string
			if p.el.PPr != nil && p.el.PPr.Style != nil {
				style = *p.el.PPr.Style
			}
			if style != q.styleName {
				continue
			}
		}

		// Collect runs and build paragraph text with run boundary tracking.
		type runSpan struct {
			run   *Run
			start int // start offset in paragraph text
			end   int // end offset (exclusive) in paragraph text
		}

		var spans []runSpan
		var sb strings.Builder
		for _, ic := range para.Content {
			if ic.Run == nil {
				continue
			}
			t := ic.Run.Text()
			start := sb.Len()
			sb.WriteString(t)
			spans = append(spans, runSpan{
				run:   &Run{doc: q.doc, el: ic.Run},
				start: start,
				end:   sb.Len(),
			})
		}

		paraText := sb.String()
		searchText := paraText
		needle := q.text

		if !q.caseSensitive && !q.useRegex {
			searchText = strings.ToLower(paraText)
			needle = strings.ToLower(q.text)
		}

		// Find all matches in this paragraph.
		var matches [][2]int // [matchStart, matchEnd) in paragraph text

		if q.useRegex {
			locs := q.pattern.FindAllStringIndex(paraText, -1)
			for _, loc := range locs {
				matches = append(matches, [2]int{loc[0], loc[1]})
			}
		} else {
			offset := 0
			for {
				idx := strings.Index(searchText[offset:], needle)
				if idx < 0 {
					break
				}
				start := offset + idx
				end := start + len(needle)
				matches = append(matches, [2]int{start, end})
				offset = start + 1
				if offset >= len(searchText) {
					break
				}
			}
		}

		for _, m := range matches {
			matchStart, matchEnd := m[0], m[1]

			// Map character offsets back to runs.
			var matchRuns []*Run
			var firstRunStart, lastRunEnd int

			for _, span := range spans {
				// Check if this run overlaps with the match range.
				if span.end <= matchStart || span.start >= matchEnd {
					continue
				}
				matchRuns = append(matchRuns, span.run)
				if len(matchRuns) == 1 {
					firstRunStart = matchStart - span.start
				}
				lastRunEnd = matchEnd - span.start
			}

			if len(matchRuns) == 0 {
				continue
			}

			// Apply nth/first filtering.
			if q.first {
				if matchIndex == 0 {
					results = append(results, SearchResult{
						Paragraph: p,
						Runs:      matchRuns,
						Start:     firstRunStart,
						End:       lastRunEnd,
					})
				}
				matchIndex++
				continue
			}
			if q.nth >= 0 {
				if matchIndex == q.nth {
					results = append(results, SearchResult{
						Paragraph: p,
						Runs:      matchRuns,
						Start:     firstRunStart,
						End:       lastRunEnd,
					})
				}
				matchIndex++
				continue
			}

			results = append(results, SearchResult{
				Paragraph: p,
				Runs:      matchRuns,
				Start:     firstRunStart,
				End:       lastRunEnd,
			})
		}
	}

	return results
}

// AddComment adds a comment to each matched range and returns the created comments.
func (q *SearchQuery) AddComment(text string, opts ...CommentOption) []*Comment {
	results := q.Results() // acquires and releases read lock
	var comments []*Comment
	for _, r := range results {
		// AddComment acquires write lock internally
		c := q.doc.AddComment(r.Runs[0], r.Runs[len(r.Runs)-1], text, opts...)
		comments = append(comments, c)
	}
	return comments
}

// AddRevision creates a tracked change revision for each matched range.
func (q *SearchQuery) AddRevision(newText string, opts ...RevisionOption) []*Revision {
	results := q.Results() // acquires and releases read lock
	var revisions []*Revision
	for _, r := range results {
		rev := q.doc.AddRevision(r.Runs[0], r.Runs[len(r.Runs)-1], newText, opts...)
		if rev != nil {
			revisions = append(revisions, rev)
		}
	}
	return revisions
}

// SetBold sets bold formatting on all runs in all matched ranges.
// Returns the number of runs modified.
func (q *SearchQuery) SetBold(v *bool) int {
	results := q.Results()
	count := 0
	for _, r := range results {
		for _, run := range r.Runs {
			run.SetBold(v)
			count++
		}
	}
	return count
}

// SetItalic sets italic formatting on all runs in all matched ranges.
// Returns the number of runs modified.
func (q *SearchQuery) SetItalic(v *bool) int {
	results := q.Results()
	count := 0
	for _, r := range results {
		for _, run := range r.Runs {
			run.SetItalic(v)
			count++
		}
	}
	return count
}

// SetStyle sets the paragraph style for all paragraphs with matches.
// Returns the number of paragraphs modified.
func (q *SearchQuery) SetStyle(name string) int {
	results := q.Results()
	count := 0
	seen := make(map[*Paragraph]bool)
	for _, r := range results {
		if !seen[r.Paragraph] {
			r.Paragraph.SetStyle(name)
			seen[r.Paragraph] = true
			count++
		}
	}
	return count
}

// ReplaceText replaces matched text with newText.
// For single-run matches, the run text is updated in place.
// For multi-run matches, the text is set on the first run and remaining runs are cleared.
// Returns the number of matches replaced.
func (q *SearchQuery) ReplaceText(newText string) int {
	results := q.Results()
	count := 0
	for _, r := range results {
		if len(r.Runs) == 0 {
			continue
		}
		if len(r.Runs) == 1 {
			// Replace just the matched portion within this run.
			run := r.Runs[0]
			t := run.el.Text()
			prefix := t[:r.Start]
			suffix := t[r.End:]
			run.SetText(prefix + newText + suffix)
		} else {
			// Multi-run match: update first run, clear middle runs, update last run.
			first := r.Runs[0]
			last := r.Runs[len(r.Runs)-1]

			firstText := first.el.Text()
			prefix := firstText[:r.Start]
			first.SetText(prefix + newText)

			// Clear middle runs.
			for _, mid := range r.Runs[1 : len(r.Runs)-1] {
				mid.SetText("")
			}

			// Trim the matched portion from the last run.
			lastText := last.el.Text()
			if r.End <= len(lastText) {
				last.SetText(lastText[r.End:])
			} else {
				last.SetText("")
			}
		}
		count++
	}
	return count
}

// ReplaceMarkdown replaces matched text with inline Markdown content while
// preserving the formatting of the matched runs as a base. The replacement
// inherits the union of all formatting from the matched runs, and Markdown
// formatting is applied additively on top.
//
// For example, if the matched text is italic and the Markdown is
// "**bold** point", the result is: "bold" as bold+italic, " point" as
// italic only.
//
// Supported inline Markdown: **bold**, *italic*, ~~strike~~, `code`,
// ***bold+italic***.
//
// Returns the number of matches replaced.
func (q *SearchQuery) ReplaceMarkdown(md string) int {
	results := q.Results()
	count := 0
	for _, r := range results {
		if len(r.Runs) == 0 {
			continue
		}

		merged := mergedRunProperties(r.Runs)
		mdRuns := parseInlineMarkdown(md)

		// Build WML runs with merged base formatting + markdown formatting.
		var newRuns []*wml.CT_R
		for _, ir := range mdRuns {
			wr := &wml.CT_R{XMLName: wmlRunName}
			wr.AddText(ir.text)

			// Start with a copy of the inherited formatting.
			var rpr wml.CT_RPr
			if merged != nil {
				rpr = *merged
			}

			// Additively apply markdown formatting.
			if ir.bold {
				v := true
				rpr.Bold = &v
			}
			if ir.italic {
				v := true
				rpr.Italic = &v
			}
			if ir.strike {
				v := true
				rpr.Strike = &v
			}
			if ir.code {
				fn := "Courier New"
				rpr.FontName = &fn
			}

			// Only set RPr if there's any formatting.
			if rpr.Bold != nil || rpr.Italic != nil || rpr.Strike != nil ||
				rpr.Underline != nil || rpr.FontName != nil || rpr.FontSize != nil ||
				rpr.Color != nil {
				wr.RPr = &rpr
			}
			newRuns = append(newRuns, wr)
		}

		if len(newRuns) == 0 {
			continue
		}

		// Replace runs in the paragraph.
		first := r.Runs[0]
		first.doc.mu.Lock()

		firstText := first.el.Text()
		prefix := firstText[:r.Start]

		if len(r.Runs) == 1 {
			suffix := firstText[r.End:]

			if prefix != "" {
				prefixRun := &wml.CT_R{XMLName: wmlRunName}
				prefixRun.RPr = first.el.RPr
				prefixRun.AddText(prefix)
				newRuns = append([]*wml.CT_R{prefixRun}, newRuns...)
			}
			if suffix != "" {
				suffixRun := &wml.CT_R{XMLName: wmlRunName}
				suffixRun.RPr = first.el.RPr
				suffixRun.AddText(suffix)
				newRuns = append(newRuns, suffixRun)
			}

			replaceRunInParagraph(first, newRuns)
		} else {
			last := r.Runs[len(r.Runs)-1]

			if prefix != "" {
				prefixRun := &wml.CT_R{XMLName: wmlRunName}
				prefixRun.RPr = first.el.RPr
				prefixRun.AddText(prefix)
				newRuns = append([]*wml.CT_R{prefixRun}, newRuns...)
			}

			lastText := last.el.Text()
			suffix := ""
			if r.End <= len(lastText) {
				suffix = lastText[r.End:]
			}
			if suffix != "" {
				suffixRun := &wml.CT_R{XMLName: wmlRunName}
				suffixRun.RPr = last.el.RPr
				suffixRun.AddText(suffix)
				newRuns = append(newRuns, suffixRun)
			}

			replaceRunsInParagraph(r.Runs, newRuns)
		}

		first.doc.mu.Unlock()
		count++
	}
	return count
}

// replaceRunInParagraph replaces a single run in its parent paragraph with
// a set of new WML runs. Caller must hold the write lock.
func replaceRunInParagraph(old *Run, newRuns []*wml.CT_R) {
	par := old.doc.findParagraphForRun(old.el)
	if par == nil {
		return
	}
	var content []wml.InlineContent
	for _, ic := range par.Content {
		if ic.Run == old.el {
			for _, nr := range newRuns {
				content = append(content, wml.InlineContent{Run: nr})
			}
		} else {
			content = append(content, ic)
		}
	}
	par.Content = content
}

// replaceRunsInParagraph replaces a contiguous set of runs in a paragraph
// with new WML runs. The first run determines the paragraph. Caller must
// hold the write lock.
func replaceRunsInParagraph(oldRuns []*Run, newRuns []*wml.CT_R) {
	if len(oldRuns) == 0 {
		return
	}
	par := oldRuns[0].doc.findParagraphForRun(oldRuns[0].el)
	if par == nil {
		return
	}

	oldSet := make(map[*wml.CT_R]bool, len(oldRuns))
	for _, r := range oldRuns {
		oldSet[r.el] = true
	}

	var content []wml.InlineContent
	inserted := false
	for _, ic := range par.Content {
		if ic.Run != nil && oldSet[ic.Run] {
			if !inserted {
				for _, nr := range newRuns {
					content = append(content, wml.InlineContent{Run: nr})
				}
				inserted = true
			}
			continue
		}
		content = append(content, ic)
	}
	par.Content = content
}

// ReplaceTextFormatted replaces matched text with newText while preserving
// the formatting of the matched runs. The replacement inherits the union of
// all formatting from the matched runs: if any matched run is bold, the
// replacement is bold; if any is italic, the replacement is italic; and so on.
//
// For single-run matches, the text is replaced in place and the run's
// properties are preserved. For multi-run matches, the merged formatting is
// applied to the first run (which receives the replacement text), and middle
// runs are cleared.
//
// This differs from [SearchQuery.ReplaceText], which does not preserve formatting.
//
// Returns the number of matches replaced.
func (q *SearchQuery) ReplaceTextFormatted(newText string) int {
	results := q.Results()
	count := 0
	for _, r := range results {
		if len(r.Runs) == 0 {
			continue
		}
		if len(r.Runs) == 1 {
			// Single-run: replace text in place; RPr stays on the element.
			run := r.Runs[0]
			t := run.el.Text()
			prefix := t[:r.Start]
			suffix := t[r.End:]
			run.SetText(prefix + newText + suffix)
		} else {
			merged := mergedRunProperties(r.Runs)
			first := r.Runs[0]
			last := r.Runs[len(r.Runs)-1]

			firstText := first.el.Text()
			prefix := firstText[:r.Start]
			first.SetText(prefix + newText)

			// Apply merged formatting to the first run.
			if merged != nil {
				first.doc.mu.Lock()
				first.el.RPr = merged
				first.doc.mu.Unlock()
			}

			for _, mid := range r.Runs[1 : len(r.Runs)-1] {
				mid.SetText("")
			}

			lastText := last.el.Text()
			if r.End <= len(lastText) {
				last.SetText(lastText[r.End:])
			} else {
				last.SetText("")
			}
		}
		count++
	}
	return count
}

// mergedRunProperties computes the union of formatting from all matched runs.
// Boolean toggles (bold, italic, strikethrough) use OR — if any run has it,
// the result has it. Value properties (font name, size, color, underline)
// use first-wins — the first run that sets the property determines the value.
// Returns nil if no runs have any formatting.
func mergedRunProperties(runs []*Run) *wml.CT_RPr {
	var merged wml.CT_RPr
	hasAny := false

	for _, r := range runs {
		rpr := r.el.RPr
		if rpr == nil {
			continue
		}
		if rpr.Bold != nil && *rpr.Bold {
			v := true
			merged.Bold = &v
			hasAny = true
		}
		if rpr.Italic != nil && *rpr.Italic {
			v := true
			merged.Italic = &v
			hasAny = true
		}
		if rpr.Strike != nil && *rpr.Strike {
			v := true
			merged.Strike = &v
			hasAny = true
		}
		if merged.Underline == nil && rpr.Underline != nil {
			s := *rpr.Underline
			merged.Underline = &s
			hasAny = true
		}
		if merged.FontName == nil && rpr.FontName != nil {
			s := *rpr.FontName
			merged.FontName = &s
			hasAny = true
		}
		if merged.FontSize == nil && rpr.FontSize != nil {
			s := *rpr.FontSize
			merged.FontSize = &s
			hasAny = true
		}
		if merged.Color == nil && rpr.Color != nil {
			s := *rpr.Color
			merged.Color = &s
			hasAny = true
		}
	}

	if !hasAny {
		return nil
	}
	return &merged
}

// Find searches the document for all occurrences of the given text string.
// The search is case-insensitive by default.
func (d *Document) Find(text string) []SearchResult {
	return d.Search(text).Results()
}

// FindRegex searches the document for all matches of the given regular expression.
func (d *Document) FindRegex(pattern *regexp.Regexp) []SearchResult {
	q := &SearchQuery{
		doc:      d,
		useRegex: true,
		pattern:  pattern,
		nth:      -1,
	}
	return q.Results()
}

// Search returns a SearchQuery for the given text, providing a fluent API
// for filtering and applying operations to matched results.
func (d *Document) Search(text string) *SearchQuery {
	return &SearchQuery{
		doc:  d,
		text: text,
		nth:  -1,
	}
}
