package docx

import (
	"encoding/xml"

	"github.com/ieshan/go-ooxml/docx/wml"
	"github.com/ieshan/go-ooxml/xmlutil"
)

// builtinStyleDef describes a built-in style's properties for the registry.
type builtinStyleDef struct {
	id          string
	name        string
	styleType   string
	isDefault   bool
	basedOn     string
	next        string
	fontName    string
	fontSize    string // half-points
	color       string // hex without #
	bold        bool
	italic      bool
	keepNext    bool
	keepLines   bool
	spaceBefore string // twips
	spaceAfter  string // twips
	outlineLvl  int    // -1 means not set
	indentLeft  string // twips
	indentRight string // twips
	extra       []xmlutil.RawXML
}

// builtinRegistry maps style IDs to their definitions.
var builtinRegistry = map[string]builtinStyleDef{
	"Normal": {
		id: "Normal", name: "Normal", styleType: "paragraph", isDefault: true,
		outlineLvl: -1,
		extra:      parseRawExtras(`<w:qFormat xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`),
	},
	"DefaultParagraphFont": {
		id: "DefaultParagraphFont", name: "Default Paragraph Font", styleType: "character", isDefault: true,
		outlineLvl: -1,
		extra: parseRawExtras(
			`<w:uiPriority xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="1"/>`,
			`<w:semiHidden xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`,
			`<w:unhideWhenUsed xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`,
		),
	},
	"TableNormal": {
		id: "TableNormal", name: "Normal Table", styleType: "table", isDefault: true,
		outlineLvl: -1,
		extra: parseRawExtras(
			`<w:uiPriority xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="99"/>`,
			`<w:semiHidden xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`,
			`<w:unhideWhenUsed xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`,
			`<w:tblPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr>`,
		),
	},
	"NoList": {
		id: "NoList", name: "No List", styleType: "numbering", isDefault: true,
		outlineLvl: -1,
		extra: parseRawExtras(
			`<w:uiPriority xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="99"/>`,
			`<w:semiHidden xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`,
			`<w:unhideWhenUsed xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`,
		),
	},
	"Heading1": headingDef("Heading1", "heading 1", "32", "2F5496", false, "240", 0),
	"Heading2": headingDef("Heading2", "heading 2", "26", "2F5496", false, "40", 1),
	"Heading3": headingDef("Heading3", "heading 3", "24", "1F3763", false, "40", 2),
	"Heading4": headingDef("Heading4", "heading 4", "22", "2F5496", true, "40", 3),
	"Heading5": headingDef("Heading5", "heading 5", "22", "2F5496", false, "40", 4),
	"Heading6": headingDef("Heading6", "heading 6", "22", "1F3763", true, "40", 5),
	"Title": {
		id: "Title", name: "Title", styleType: "paragraph", basedOn: "Normal", next: "Normal",
		fontName: "Calibri Light", fontSize: "56", outlineLvl: -1,
		spaceAfter: "0",
		extra: parseRawExtras(
			`<w:uiPriority xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="10"/>`,
			`<w:qFormat xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`,
		),
	},
	"Subtitle": {
		id: "Subtitle", name: "Subtitle", styleType: "paragraph", basedOn: "Normal", next: "Normal",
		color: "5A5A5A", outlineLvl: -1,
		extra: parseRawExtras(
			`<w:uiPriority xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="11"/>`,
			`<w:qFormat xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`,
		),
	},
	"ListParagraph": {
		id: "ListParagraph", name: "List Paragraph", styleType: "paragraph", basedOn: "Normal",
		indentLeft: "720", outlineLvl: -1,
		extra: parseRawExtras(
			`<w:uiPriority xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="34"/>`,
			`<w:qFormat xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`,
		),
	},
	"Quote": {
		id: "Quote", name: "Quote", styleType: "paragraph", basedOn: "Normal", next: "Normal",
		italic: true, color: "404040", indentLeft: "720", indentRight: "720",
		spaceBefore: "200", outlineLvl: -1,
		extra: parseRawExtras(
			`<w:uiPriority xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="29"/>`,
			`<w:qFormat xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`,
		),
	},
	"IntenseQuote": {
		id: "IntenseQuote", name: "Intense Quote", styleType: "paragraph", basedOn: "Normal", next: "Normal",
		bold: true, color: "2F5496", indentLeft: "720", indentRight: "720",
		spaceBefore: "200", outlineLvl: -1,
		extra: parseRawExtras(
			`<w:uiPriority xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="30"/>`,
			`<w:qFormat xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`,
		),
	},
	"NoSpacing": {
		id: "NoSpacing", name: "No Spacing", styleType: "paragraph", basedOn: "Normal",
		spaceAfter: "0", outlineLvl: -1,
		extra: parseRawExtras(
			`<w:uiPriority xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="1"/>`,
			`<w:qFormat xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`,
		),
	},
	"Strong": {
		id: "Strong", name: "Strong", styleType: "character",
		bold: true, outlineLvl: -1,
		extra: parseRawExtras(
			`<w:uiPriority xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="22"/>`,
			`<w:qFormat xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`,
		),
	},
	"Emphasis": {
		id: "Emphasis", name: "Emphasis", styleType: "character",
		italic: true, outlineLvl: -1,
		extra: parseRawExtras(
			`<w:uiPriority xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="20"/>`,
			`<w:qFormat xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`,
		),
	},
}

func headingDef(id, name, fontSize, color string, italic bool, spaceBefore string, outlineLvl int) builtinStyleDef {
	return builtinStyleDef{
		id: id, name: name, styleType: "paragraph",
		basedOn: "Normal", next: "Normal",
		fontName: "Calibri Light", fontSize: fontSize, color: color,
		italic: italic, keepNext: true, keepLines: true,
		spaceBefore: spaceBefore, spaceAfter: "0", outlineLvl: outlineLvl,
		extra: parseRawExtras(
			`<w:uiPriority xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="9"/>`,
			`<w:qFormat xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`,
		),
	}
}

// parseRawExtras parses XML strings into RawXML values.
func parseRawExtras(xmlStrs ...string) []xmlutil.RawXML {
	var extras []xmlutil.RawXML
	for _, s := range xmlStrs {
		var raw xmlutil.RawXML
		if err := xml.Unmarshal([]byte(s), &raw); err == nil {
			extras = append(extras, raw)
		}
	}
	return extras
}

// buildCTStyle creates a wml.CT_Style from a builtinStyleDef.
func buildCTStyle(def builtinStyleDef) *wml.CT_Style {
	s := &wml.CT_Style{
		XMLName: xml.Name{Space: wml.Ns, Local: "style"},
		Type:    def.styleType,
		StyleID: def.id,
		Name:    def.name,
		Extra:   def.extra,
	}
	if def.isDefault {
		v := true
		s.Default = &v
	}
	if def.basedOn != "" {
		s.BasedOn = &def.basedOn
	}
	if def.next != "" {
		s.Next = &def.next
	}

	// Build PPr
	needsPPr := def.keepNext || def.keepLines || def.spaceBefore != "" || def.spaceAfter != "" ||
		def.indentLeft != "" || def.indentRight != "" || def.outlineLvl >= 0
	if needsPPr {
		ppr := &wml.CT_PPr{XMLName: xml.Name{Space: wml.Ns, Local: "pPr"}}
		if def.keepNext {
			v := true
			ppr.KeepNext = &v
		}
		if def.keepLines {
			v := true
			ppr.KeepLines = &v
		}
		if def.spaceBefore != "" || def.spaceAfter != "" {
			sp := &wml.CT_Spacing{}
			if def.spaceBefore != "" {
				sp.Before = &def.spaceBefore
			}
			if def.spaceAfter != "" {
				sp.After = &def.spaceAfter
			}
			ppr.Spacing = sp
		}
		if def.indentLeft != "" || def.indentRight != "" {
			ind := &wml.CT_Ind{}
			if def.indentLeft != "" {
				ind.Left = &def.indentLeft
			}
			if def.indentRight != "" {
				ind.Right = &def.indentRight
			}
			ppr.Ind = ind
		}
		if def.outlineLvl >= 0 {
			ppr.OutlineLvl = &def.outlineLvl
		}
		s.PPr = ppr
	}

	// Build RPr
	needsRPr := def.fontName != "" || def.fontSize != "" || def.color != "" || def.bold || def.italic
	if needsRPr {
		rpr := &wml.CT_RPr{}
		if def.fontName != "" {
			rpr.FontName = &def.fontName
			rpr.FontHAnsi = &def.fontName
			rpr.FontEastAsia = &def.fontName
			rpr.FontCS = &def.fontName
		}
		if def.fontSize != "" {
			rpr.FontSize = &def.fontSize
			rpr.FontSizeCS = &def.fontSize
		}
		if def.color != "" {
			rpr.Color = &def.color
		}
		if def.bold {
			v := true
			rpr.Bold = &v
		}
		if def.italic {
			v := true
			rpr.Italic = &v
		}
		s.RPr = rpr
	}

	return s
}

// docDefaultsXML is the raw XML for <w:docDefaults>.
const docDefaultsXML = `<w:docDefaults xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:eastAsia="Calibri" w:hAnsi="Calibri" w:cs="Times New Roman"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-US" w:eastAsia="en-US" w:bidi="ar-SA"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>`

// latentStylesXML is the raw XML for <w:latentStyles>.
const latentStylesXML = `<w:latentStyles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:defLockedState="0" w:defUIPriority="99" w:defSemiHidden="1" w:defUnhideWhenUsed="1" w:defQFormat="0" w:count="276"><w:lsdException w:name="Normal" w:semiHidden="0" w:uiPriority="0" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="heading 1" w:semiHidden="0" w:uiPriority="9" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="heading 2" w:uiPriority="9" w:qFormat="1"/><w:lsdException w:name="heading 3" w:uiPriority="9" w:qFormat="1"/><w:lsdException w:name="heading 4" w:uiPriority="9" w:qFormat="1"/><w:lsdException w:name="heading 5" w:uiPriority="9" w:qFormat="1"/><w:lsdException w:name="heading 6" w:uiPriority="9" w:qFormat="1"/><w:lsdException w:name="heading 7" w:uiPriority="9" w:qFormat="1"/><w:lsdException w:name="heading 8" w:uiPriority="9" w:qFormat="1"/><w:lsdException w:name="heading 9" w:uiPriority="9" w:qFormat="1"/><w:lsdException w:name="toc 1" w:uiPriority="39"/><w:lsdException w:name="toc 2" w:uiPriority="39"/><w:lsdException w:name="toc 3" w:uiPriority="39"/><w:lsdException w:name="toc 4" w:uiPriority="39"/><w:lsdException w:name="toc 5" w:uiPriority="39"/><w:lsdException w:name="toc 6" w:uiPriority="39"/><w:lsdException w:name="toc 7" w:uiPriority="39"/><w:lsdException w:name="toc 8" w:uiPriority="39"/><w:lsdException w:name="toc 9" w:uiPriority="39"/><w:lsdException w:name="caption" w:uiPriority="35" w:qFormat="1"/><w:lsdException w:name="Title" w:semiHidden="0" w:uiPriority="10" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Default Paragraph Font" w:uiPriority="1"/><w:lsdException w:name="Subtitle" w:semiHidden="0" w:uiPriority="11" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Strong" w:semiHidden="0" w:uiPriority="22" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Emphasis" w:semiHidden="0" w:uiPriority="20" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Table Grid" w:semiHidden="0" w:uiPriority="59" w:unhideWhenUsed="0"/><w:lsdException w:name="No Spacing" w:semiHidden="0" w:uiPriority="1" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Light Shading" w:semiHidden="0" w:uiPriority="60" w:unhideWhenUsed="0"/><w:lsdException w:name="Light List" w:semiHidden="0" w:uiPriority="61" w:unhideWhenUsed="0"/><w:lsdException w:name="Light Grid" w:semiHidden="0" w:uiPriority="62" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 1" w:semiHidden="0" w:uiPriority="63" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 2" w:semiHidden="0" w:uiPriority="64" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium List 1" w:semiHidden="0" w:uiPriority="65" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium List 2" w:semiHidden="0" w:uiPriority="66" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 1" w:semiHidden="0" w:uiPriority="67" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 2" w:semiHidden="0" w:uiPriority="68" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 3" w:semiHidden="0" w:uiPriority="69" w:unhideWhenUsed="0"/><w:lsdException w:name="Dark List" w:semiHidden="0" w:uiPriority="70" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Shading" w:semiHidden="0" w:uiPriority="71" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful List" w:semiHidden="0" w:uiPriority="72" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Grid" w:semiHidden="0" w:uiPriority="73" w:unhideWhenUsed="0"/><w:lsdException w:name="Light Shading Accent 1" w:semiHidden="0" w:uiPriority="60" w:unhideWhenUsed="0"/><w:lsdException w:name="Light List Accent 1" w:semiHidden="0" w:uiPriority="61" w:unhideWhenUsed="0"/><w:lsdException w:name="Light Grid Accent 1" w:semiHidden="0" w:uiPriority="62" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 1 Accent 1" w:semiHidden="0" w:uiPriority="63" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 2 Accent 1" w:semiHidden="0" w:uiPriority="64" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium List 1 Accent 1" w:semiHidden="0" w:uiPriority="65" w:unhideWhenUsed="0"/><w:lsdException w:name="List Paragraph" w:semiHidden="0" w:uiPriority="34" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Quote" w:semiHidden="0" w:uiPriority="29" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Intense Quote" w:semiHidden="0" w:uiPriority="30" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Medium List 2 Accent 1" w:semiHidden="0" w:uiPriority="66" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 1 Accent 1" w:semiHidden="0" w:uiPriority="67" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 2 Accent 1" w:semiHidden="0" w:uiPriority="68" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 3 Accent 1" w:semiHidden="0" w:uiPriority="69" w:unhideWhenUsed="0"/><w:lsdException w:name="Dark List Accent 1" w:semiHidden="0" w:uiPriority="70" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Shading Accent 1" w:semiHidden="0" w:uiPriority="71" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful List Accent 1" w:semiHidden="0" w:uiPriority="72" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Grid Accent 1" w:semiHidden="0" w:uiPriority="73" w:unhideWhenUsed="0"/><w:lsdException w:name="TOC Heading" w:uiPriority="39" w:qFormat="1"/></w:latentStyles>`

// initDefaultStyles creates the default styles.xml content for a new document.
// Called by New() and by ensureStyleExists() when styles is nil.
func (d *Document) initDefaultStyles() {
	d.styles = &wml.CT_Styles{
		XMLName: xml.Name{Space: wml.Ns, Local: "styles"},
	}

	// Parse and add docDefaults and latentStyles as Extra (raw XML).
	var docDefaults xmlutil.RawXML
	if err := xml.Unmarshal([]byte(docDefaultsXML), &docDefaults); err == nil {
		d.styles.Extra = append(d.styles.Extra, docDefaults)
	}
	var latentStyles xmlutil.RawXML
	if err := xml.Unmarshal([]byte(latentStylesXML), &latentStyles); err == nil {
		d.styles.Extra = append(d.styles.Extra, latentStyles)
	}

	// Add all built-in style definitions in a stable order.
	order := []string{
		"Normal", "DefaultParagraphFont", "TableNormal", "NoList",
		"Heading1", "Heading2", "Heading3", "Heading4", "Heading5", "Heading6",
		"Title", "Subtitle", "ListParagraph", "Quote", "IntenseQuote", "NoSpacing",
		"Strong", "Emphasis",
	}
	for _, id := range order {
		if def, ok := builtinRegistry[id]; ok {
			d.styles.Styles = append(d.styles.Styles, buildCTStyle(def))
		}
	}
}

// ensureStyleExists checks if a style definition exists and creates it
// from the built-in registry if it doesn't. Does nothing for unknown styles.
// Caller must hold d.mu (Lock or RLock).
func (d *Document) ensureStyleExists(name string) {
	if d.styles == nil {
		d.initDefaultStyles()
	}

	for _, s := range d.styles.Styles {
		if s.StyleID == name {
			return
		}
	}

	def, ok := builtinRegistry[name]
	if !ok {
		return
	}

	d.styles.Styles = append(d.styles.Styles, buildCTStyle(def))
}
