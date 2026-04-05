package docx

import (
	"strconv"
	"strings"
)

// PMNode represents a ProseMirror document node.
type PMNode struct {
	Type    string         `json:"type"`
	Attrs   map[string]any `json:"attrs,omitempty"`
	Content []PMNode       `json:"content,omitempty"`
	Marks   []PMMark       `json:"marks,omitempty"`
	Text    string         `json:"text,omitempty"`
}

// PMMark represents a ProseMirror mark (inline formatting).
type PMMark struct {
	Type  string         `json:"type"`
	Attrs map[string]any `json:"attrs,omitempty"`
}

// ProseMirrorOptions configures ProseMirror JSON export.
type ProseMirrorOptions struct {
	// IncludeHeaders includes header/footer content in doc node attrs.
	IncludeHeaders bool
	// UseTipTapNames uses TipTap naming convention (bold/italic/orderedList)
	// instead of vanilla ProseMirror (strong/em/ordered_list). Default: true.
	UseTipTapNames bool
}

func defaultPMOptions(opts *ProseMirrorOptions) *ProseMirrorOptions {
	if opts != nil {
		return opts
	}
	return &ProseMirrorOptions{UseTipTapNames: true}
}

// nodeNameAliases maps vanilla ProseMirror node/mark names to TipTap names.
var nodeNameAliases = map[string]string{
	"strong":          "bold",
	"em":              "italic",
	"ordered_list":    "orderedList",
	"bullet_list":     "bulletList",
	"list_item":       "listItem",
	"code_block":      "codeBlock",
	"hard_break":      "hardBreak",
	"horizontal_rule": "horizontalRule",
	"table_row":       "tableRow",
	"table_cell":      "tableCell",
	"table_header":    "tableHeader",
	"task_list":       "taskList",
	"task_item":       "taskItem",
	"page_break":      "pageBreak",
	"section_break":   "sectionBreak",
}

// tipTapToVanilla maps TipTap names to vanilla ProseMirror names for export.
var tipTapToVanilla = map[string]string{
	"bold":           "strong",
	"italic":         "em",
	"orderedList":    "ordered_list",
	"bulletList":     "bullet_list",
	"listItem":       "list_item",
	"codeBlock":      "code_block",
	"hardBreak":      "hard_break",
	"horizontalRule": "horizontal_rule",
	"tableRow":       "table_row",
	"tableCell":      "table_cell",
	"tableHeader":    "table_header",
	"taskList":       "task_list",
	"taskItem":       "task_item",
	"pageBreak":      "page_break",
	"sectionBreak":   "section_break",
}

// normalizeTypeName converts any vanilla ProseMirror alias to the canonical TipTap name.
func normalizeTypeName(t string) string {
	if canon, ok := nodeNameAliases[t]; ok {
		return canon
	}
	return t
}

// exportName returns the node/mark type name for the selected convention.
func exportName(name string, useTipTap bool) string {
	if !useTipTap {
		if vanilla, ok := tipTapToVanilla[name]; ok {
			return vanilla
		}
	}
	return name
}

// pmTwipsToPoints converts twips to points (1 pt = 20 twips).
func pmTwipsToPoints(twips int64) float64 {
	return float64(twips) / 20.0
}

// pmPointsToTwips converts points to twips.
func pmPointsToTwips(pts float64) int64 {
	return int64(pts * 20)
}

// pmHalfPointsToPtString converts half-points (w:sz value like "24") to "12pt".
func pmHalfPointsToPtString(hp string) string {
	n, err := strconv.ParseFloat(hp, 64)
	if err != nil {
		return hp
	}
	pts := n / 2
	if pts == float64(int(pts)) {
		return strconv.Itoa(int(pts)) + "pt"
	}
	return strconv.FormatFloat(pts, 'f', 1, 64) + "pt"
}

// pmPtStringToHalfPoints converts "12pt" or "12px" to half-points string "24".
func pmPtStringToHalfPoints(s string) string {
	s = strings.TrimSuffix(s, "pt")
	s = strings.TrimSuffix(s, "px")
	n, err := strconv.ParseFloat(strings.TrimSpace(s), 64)
	if err != nil {
		return s
	}
	return strconv.Itoa(int(n * 2))
}
