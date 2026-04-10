package wml

import (
	"encoding/xml"
	"strings"
	"testing"

	"github.com/ieshan/go-ooxml/xmlutil"
)

func TestCT_R_Unmarshal_PlainText(t *testing.T) {
	input := `<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:t>Hello</w:t></w:r>`
	var r CT_R
	if err := xmlutil.Unmarshal([]byte(input), &r); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if r.Text() != "Hello" {
		t.Errorf("Text() = %q, want %q", r.Text(), "Hello")
	}
}

func TestCT_R_Unmarshal_WithBold(t *testing.T) {
	input := `<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:rPr><w:b/></w:rPr>
		<w:t>Bold</w:t>
	</w:r>`
	var r CT_R
	if err := xmlutil.Unmarshal([]byte(input), &r); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if r.RPr == nil {
		t.Fatal("RPr nil")
	}
	if r.RPr.Bold == nil || *r.RPr.Bold != true {
		t.Error("Bold should be true")
	}
}

func TestCT_R_Unmarshal_BoldFalse(t *testing.T) {
	input := `<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:rPr><w:b w:val="false"/></w:rPr>
		<w:t>Not Bold</w:t>
	</w:r>`
	var r CT_R
	if err := xmlutil.Unmarshal([]byte(input), &r); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if r.RPr.Bold == nil {
		t.Fatal("Bold nil, want explicit false")
	}
	if *r.RPr.Bold != false {
		t.Error("Bold should be false")
	}
}

func TestCT_R_Unmarshal_BoldItalic(t *testing.T) {
	input := `<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:rPr><w:b/><w:i/></w:rPr>
		<w:t>Both</w:t>
	</w:r>`
	var r CT_R
	if err := xmlutil.Unmarshal([]byte(input), &r); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if r.RPr.Bold == nil || *r.RPr.Bold != true {
		t.Error("Bold")
	}
	if r.RPr.Italic == nil || *r.RPr.Italic != true {
		t.Error("Italic")
	}
}

func TestCT_R_Marshal_RoundTrip(t *testing.T) {
	r := CT_R{XMLName: xml.Name{Space: Ns, Local: "r"}}
	r.AddText("Hello World")
	r.RPr = &CT_RPr{Bold: new(true)}

	data, err := xmlutil.Marshal(&r, xmlutil.OOXML)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}

	var r2 CT_R
	if err := xmlutil.Unmarshal(data, &r2); err != nil {
		t.Fatalf("re-Unmarshal: %v", err)
	}
	if r2.Text() != "Hello World" {
		t.Errorf("text = %q", r2.Text())
	}
	if r2.RPr == nil || r2.RPr.Bold == nil || *r2.RPr.Bold != true {
		t.Error("lost bold")
	}
}

func TestCT_R_Unmarshal_WithFontInfo(t *testing.T) {
	input := `<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:rPr>
			<w:rFonts w:ascii="Arial"/>
			<w:sz w:val="24"/>
			<w:color w:val="FF0000"/>
			<w:rStyle w:val="Strong"/>
		</w:rPr>
		<w:t>styled</w:t>
	</w:r>`
	var r CT_R
	if err := xmlutil.Unmarshal([]byte(input), &r); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if r.RPr.FontName == nil || *r.RPr.FontName != "Arial" {
		t.Error("FontName")
	}
	if r.RPr.FontSize == nil || *r.RPr.FontSize != "24" {
		t.Error("FontSize")
	}
	if r.RPr.Color == nil || *r.RPr.Color != "FF0000" {
		t.Error("Color")
	}
	if r.RPr.RunStyle == nil || *r.RPr.RunStyle != "Strong" {
		t.Error("RunStyle")
	}
}

func TestCT_R_Unmarshal_Tab_Break(t *testing.T) {
	input := `<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:t>before</w:t><w:tab/><w:t>after</w:t><w:br/>
	</w:r>`
	var r CT_R
	if err := xmlutil.Unmarshal([]byte(input), &r); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if len(r.Content) != 4 {
		t.Errorf("content count = %d, want 4", len(r.Content))
	}
	if r.Text() != "before\tafter\n" {
		t.Errorf("Text() = %q", r.Text())
	}
}

func TestCT_R_Unmarshal_PreservesUnknown(t *testing.T) {
	input := `<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:t>text</w:t><w:noProof/>
	</w:r>`
	var r CT_R
	if err := xmlutil.Unmarshal([]byte(input), &r); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if len(r.Content) != 2 {
		t.Errorf("content count = %d, want 2", len(r.Content))
	}
	if r.Content[1].Raw == nil {
		t.Error("unknown element should be preserved as RawXML")
	}
}

func TestCT_R_Unmarshal_AnnotationRef(t *testing.T) {
	input := `<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:t>text</w:t><w:annotationRef/>
	</w:r>`
	var r CT_R
	if err := xmlutil.Unmarshal([]byte(input), &r); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if len(r.Content) != 2 {
		t.Errorf("content count = %d, want 2", len(r.Content))
	}
	if r.Content[1].AnnotationRef == nil {
		t.Error("annotationRef should be parsed as CT_AnnotationRef")
	}
}

func TestCT_R_Marshal_WithUnderline(t *testing.T) {
	r := CT_R{XMLName: xml.Name{Space: Ns, Local: "r"}}
	r.RPr = &CT_RPr{Underline: new("single")}
	r.AddText("underlined")
	data, err := xmlutil.Marshal(&r, xmlutil.OOXML)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	var r2 CT_R
	if err := xmlutil.Unmarshal(data, &r2); err != nil {
		t.Fatalf("re-Unmarshal: %v", err)
	}
	if r2.RPr == nil || r2.RPr.Underline == nil || *r2.RPr.Underline != "single" {
		t.Error("Underline lost")
	}
}

func TestCT_R_Marshal_WithColorFontSizeFontNameStyle(t *testing.T) {
	r := CT_R{XMLName: xml.Name{Space: Ns, Local: "r"}}
	r.RPr = &CT_RPr{
		Color:    new("FF0000"),
		FontSize: new("28"),
		FontName: new("Times New Roman"),
		RunStyle: new("Emphasis"),
	}
	r.AddText("styled run")
	data, err := xmlutil.Marshal(&r, xmlutil.OOXML)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	var r2 CT_R
	if err := xmlutil.Unmarshal(data, &r2); err != nil {
		t.Fatalf("re-Unmarshal: %v", err)
	}
	if r2.RPr == nil {
		t.Fatal("RPr nil")
	}
	if r2.RPr.Color == nil || *r2.RPr.Color != "FF0000" {
		t.Error("Color")
	}
	if r2.RPr.FontSize == nil || *r2.RPr.FontSize != "28" {
		t.Error("FontSize")
	}
	if r2.RPr.FontName == nil || *r2.RPr.FontName != "Times New Roman" {
		t.Error("FontName")
	}
	if r2.RPr.RunStyle == nil || *r2.RPr.RunStyle != "Emphasis" {
		t.Error("RunStyle")
	}
}

func TestCT_R_Marshal_TabBreakText(t *testing.T) {
	r := CT_R{XMLName: xml.Name{Space: Ns, Local: "r"}}
	r.Content = append(r.Content,
		RunContent{Tab: &CT_Tab{}},
		RunContent{Text: &CT_Text{Value: "mid"}},
		RunContent{Break: &CT_Break{Type: "page"}},
	)
	data, err := xmlutil.Marshal(&r, xmlutil.OOXML)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	var r2 CT_R
	if err := xmlutil.Unmarshal(data, &r2); err != nil {
		t.Fatalf("re-Unmarshal: %v", err)
	}
	if len(r2.Content) != 3 {
		t.Fatalf("content = %d, want 3", len(r2.Content))
	}
	if r2.Content[0].Tab == nil {
		t.Error("tab not preserved")
	}
	if r2.Content[1].Text == nil || r2.Content[1].Text.Value != "mid" {
		t.Error("text not preserved")
	}
	if r2.Content[2].Break == nil {
		t.Error("break not preserved")
	}
}

func TestCT_R_Marshal_DelText(t *testing.T) {
	r := CT_R{XMLName: xml.Name{Space: Ns, Local: "r"}}
	r.Content = append(r.Content, RunContent{DelText: &CT_Text{Value: "deleted content"}})
	data, err := xmlutil.Marshal(&r, xmlutil.OOXML)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	var r2 CT_R
	if err := xmlutil.Unmarshal(data, &r2); err != nil {
		t.Fatalf("re-Unmarshal: %v", err)
	}
	if len(r2.Content) != 1 || r2.Content[0].DelText == nil {
		t.Fatal("DelText not preserved")
	}
	if r2.Content[0].DelText.Value != "deleted content" {
		t.Errorf("DelText = %q", r2.Content[0].DelText.Value)
	}
}

func TestCT_RPr_MultipleFmtCombined(t *testing.T) {
	r := CT_R{XMLName: xml.Name{Space: Ns, Local: "r"}}
	r.RPr = &CT_RPr{
		Bold:      new(true),
		Italic:    new(true),
		Strike:    new(true),
		Underline: new("double"),
		Color:     new("00FF00"),
		FontSize:  new("32"),
		FontName:  new("Arial"),
		RunStyle:  new("Strong"),
	}
	r.AddText("all formats")
	data, err := xmlutil.Marshal(&r, xmlutil.OOXML)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	var r2 CT_R
	if err := xmlutil.Unmarshal(data, &r2); err != nil {
		t.Fatalf("re-Unmarshal: %v", err)
	}
	rpr := r2.RPr
	if rpr == nil {
		t.Fatal("RPr nil")
	}
	if rpr.Bold == nil || !*rpr.Bold {
		t.Error("Bold")
	}
	if rpr.Italic == nil || !*rpr.Italic {
		t.Error("Italic")
	}
	if rpr.Strike == nil || !*rpr.Strike {
		t.Error("Strike")
	}
	if rpr.Underline == nil || *rpr.Underline != "double" {
		t.Error("Underline")
	}
	if rpr.Color == nil || *rpr.Color != "00FF00" {
		t.Error("Color")
	}
}

func TestCT_RPr_ExtraPreservation(t *testing.T) {
	input := `<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:rPr>
			<w:b/>
			<w:vertAlign w:val="superscript"/>
		</w:rPr>
		<w:t>sup</w:t>
	</w:r>`
	var r CT_R
	if err := xmlutil.Unmarshal([]byte(input), &r); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if r.RPr == nil {
		t.Fatal("RPr nil")
	}
	if len(r.RPr.Extra) != 1 {
		t.Errorf("Extra = %d, want 1", len(r.RPr.Extra))
	}
	// Marshal and check Extra preserved
	data, err := xmlutil.Marshal(&r, xmlutil.OOXML)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	var r2 CT_R
	if err := xmlutil.Unmarshal(data, &r2); err != nil {
		t.Fatalf("re-Unmarshal: %v", err)
	}
	if r2.RPr == nil || len(r2.RPr.Extra) != 1 {
		t.Errorf("Extra not preserved after round-trip, got %d", len(r2.RPr.Extra))
	}
}

func TestCT_R_Marshal_Raw_Content(t *testing.T) {
	// Test that raw run content is preserved through marshal
	input := `<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:t>text</w:t>
		<w:noProof/>
	</w:r>`
	var r CT_R
	if err := xmlutil.Unmarshal([]byte(input), &r); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	data, err := xmlutil.Marshal(&r, xmlutil.OOXML)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	var r2 CT_R
	if err := xmlutil.Unmarshal(data, &r2); err != nil {
		t.Fatalf("re-Unmarshal: %v", err)
	}
	if len(r2.Content) != 2 {
		t.Fatalf("content = %d, want 2", len(r2.Content))
	}
	if r2.Content[1].Raw == nil {
		t.Error("raw content lost")
	}
}

func TestCT_R_Marshal_AnnotationRef(t *testing.T) {
	// Test that annotationRef round-trips correctly
	input := `<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:t>text</w:t>
		<w:annotationRef/>
	</w:r>`
	var r CT_R
	if err := xmlutil.Unmarshal([]byte(input), &r); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	data, err := xmlutil.Marshal(&r, xmlutil.OOXML)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	var r2 CT_R
	if err := xmlutil.Unmarshal(data, &r2); err != nil {
		t.Fatalf("re-Unmarshal: %v", err)
	}
	if len(r2.Content) != 2 {
		t.Fatalf("content = %d, want 2", len(r2.Content))
	}
	if r2.Content[1].AnnotationRef == nil {
		t.Error("annotationRef lost during round-trip")
	}
}

func TestParseBoolVal_TrueVariants(t *testing.T) {
	// Test parseBoolVal with "true" value (hits the default case returning true)
	attrs := []xml.Attr{{Name: xml.Name{Local: "val"}, Value: "true"}}
	if !parseBoolVal(attrs) {
		t.Error("expected true for val='true'")
	}
	// With "1"
	attrs1 := []xml.Attr{{Name: xml.Name{Local: "val"}, Value: "1"}}
	if !parseBoolVal(attrs1) {
		t.Error("expected true for val='1'")
	}
}

func TestGetAttrVal_NoMatch(t *testing.T) {
	// Test getAttrVal when no val attr (returns "")
	attrs := []xml.Attr{{Name: xml.Name{Local: "other"}, Value: "x"}}
	v := getAttrVal(attrs)
	if v != "" {
		t.Errorf("getAttrVal = %q, want empty", v)
	}
}

func TestGetAttr_NoMatch(t *testing.T) {
	// Test getAttr when no matching attr (returns "")
	attrs := []xml.Attr{{Name: xml.Name{Local: "other"}, Value: "x"}}
	v := getAttr(attrs, "ascii")
	if v != "" {
		t.Errorf("getAttr = %q, want empty", v)
	}
}

func TestCT_RPr_Marshal_OnlyColor(t *testing.T) {
	// RPr with only color set — Bold/Italic/Strike are nil, exercises marshalBoolToggle(nil) path
	r := CT_R{XMLName: xml.Name{Space: Ns, Local: "r"}}
	r.RPr = &CT_RPr{Color: new("FF0000")}
	r.AddText("colored text")
	data, err := xmlutil.Marshal(&r, xmlutil.OOXML)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	var r2 CT_R
	if err := xmlutil.Unmarshal(data, &r2); err != nil {
		t.Fatalf("re-Unmarshal: %v", err)
	}
	if r2.RPr == nil || r2.RPr.Color == nil || *r2.RPr.Color != "FF0000" {
		t.Error("Color")
	}
	// Bold/Italic/Strike should be nil (not marshaled)
	if r2.RPr.Bold != nil {
		t.Error("Bold should be nil")
	}
}

func TestCT_RPr_rFonts_NoAscii(t *testing.T) {
	// rFonts with no ascii attr — FontName should remain nil
	input := `<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:rPr><w:rFonts w:hAnsi="Times New Roman"/></w:rPr>
		<w:t>text</w:t>
	</w:r>`
	var r CT_R
	if err := xmlutil.Unmarshal([]byte(input), &r); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if r.RPr == nil {
		t.Fatal("RPr nil")
	}
	if r.RPr.FontName != nil {
		t.Errorf("FontName should be nil when no ascii attr, got %q", *r.RPr.FontName)
	}
}

func TestCT_RPr_Underline_EmptyVal(t *testing.T) {
	// u element with no val attr — Underline should remain nil
	input := `<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:rPr><w:u/></w:rPr>
		<w:t>text</w:t>
	</w:r>`
	var r CT_R
	if err := xmlutil.Unmarshal([]byte(input), &r); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if r.RPr == nil {
		t.Fatal("RPr nil")
	}
	// u with no val has empty string — Underline should be nil (empty string check)
	if r.RPr.Underline != nil {
		t.Errorf("Underline should be nil for empty val, got %q", *r.RPr.Underline)
	}
}

func TestCT_R_Text_ExcludesDelText(t *testing.T) {
	// CT_R.Text() should NOT include delText — deleted text is not visible.
	r := CT_R{}
	r.Content = append(r.Content, RunContent{DelText: &CT_Text{Value: "deleted"}})
	if r.Text() != "" {
		t.Errorf("Text() = %q, want empty (delText should be excluded)", r.Text())
	}
}

func TestCT_R_DelText(t *testing.T) {
	// CT_R.DelText() should return delText content.
	r := CT_R{}
	r.Content = append(r.Content, RunContent{DelText: &CT_Text{Value: "deleted"}})
	if r.DelText() != "deleted" {
		t.Errorf("DelText() = %q, want %q", r.DelText(), "deleted")
	}
}

func TestCT_R_AddText_LeadingTrailingSpace(t *testing.T) {
	r := CT_R{XMLName: xml.Name{Space: Ns, Local: "r"}}
	r.AddText(" leading")
	r.AddText("trailing ")
	r.AddText("neither")
	if len(r.Content) != 3 {
		t.Fatalf("content = %d", len(r.Content))
	}
	if r.Content[0].Text.Space != "preserve" {
		t.Error("leading space not preserved")
	}
	if r.Content[1].Text.Space != "preserve" {
		t.Error("trailing space not preserved")
	}
	if r.Content[2].Text.Space != "" {
		t.Error("no-space should be empty")
	}
}

func TestCT_R_Marshal_StrikeBoldFalse(t *testing.T) {
	r := CT_R{XMLName: xml.Name{Space: Ns, Local: "r"}}
	r.RPr = &CT_RPr{Bold: new(false), Strike: new(true)}
	r.AddText("explicit false bold")
	data, err := xmlutil.Marshal(&r, xmlutil.OOXML)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	var r2 CT_R
	if err := xmlutil.Unmarshal(data, &r2); err != nil {
		t.Fatalf("re-Unmarshal: %v", err)
	}
	if r2.RPr == nil {
		t.Fatal("RPr nil")
	}
	if r2.RPr.Bold == nil || *r2.RPr.Bold != false {
		t.Error("Bold false not preserved")
	}
	if r2.RPr.Strike == nil || !*r2.RPr.Strike {
		t.Error("Strike not preserved")
	}
}

// ---------------------------------------------------------------------------
// OOXML Compliance Tests: rFonts and color round-trip
// ---------------------------------------------------------------------------

func TestCT_RPr_RFonts_AllAttrs_RoundTrip(t *testing.T) {
	input := `<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:rPr>
			<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="MS Mincho" w:cs="Arial"/>
		</w:rPr>
		<w:t>text</w:t>
	</w:r>`
	var r CT_R
	if err := xmlutil.Unmarshal([]byte(input), &r); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if r.RPr == nil {
		t.Fatal("RPr nil")
	}
	if r.RPr.FontName == nil || *r.RPr.FontName != "Arial" {
		t.Errorf("FontName = %v", r.RPr.FontName)
	}
	if r.RPr.FontHAnsi == nil || *r.RPr.FontHAnsi != "Arial" {
		t.Errorf("FontHAnsi = %v", r.RPr.FontHAnsi)
	}
	if r.RPr.FontEastAsia == nil || *r.RPr.FontEastAsia != "MS Mincho" {
		t.Errorf("FontEastAsia = %v", r.RPr.FontEastAsia)
	}
	if r.RPr.FontCS == nil || *r.RPr.FontCS != "Arial" {
		t.Errorf("FontCS = %v", r.RPr.FontCS)
	}

	// Round-trip.
	data, err := xmlutil.Marshal(&r, xmlutil.OOXML)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	var r2 CT_R
	if err := xmlutil.Unmarshal(data, &r2); err != nil {
		t.Fatalf("re-Unmarshal: %v", err)
	}
	if r2.RPr.FontHAnsi == nil || *r2.RPr.FontHAnsi != "Arial" {
		t.Error("FontHAnsi lost during round-trip")
	}
	if r2.RPr.FontEastAsia == nil || *r2.RPr.FontEastAsia != "MS Mincho" {
		t.Error("FontEastAsia lost during round-trip")
	}
	if r2.RPr.FontCS == nil || *r2.RPr.FontCS != "Arial" {
		t.Error("FontCS lost during round-trip")
	}
}

func TestCT_RPr_RFonts_AsciiOnly_AddsHAnsi(t *testing.T) {
	// When only ascii is set, marshal should also emit hAnsi (same value).
	r := CT_R{XMLName: xml.Name{Space: Ns, Local: "r"}}
	r.RPr = &CT_RPr{FontName: new("Courier New")}
	r.AddText("code")
	data, err := xmlutil.Marshal(&r, xmlutil.OOXML)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	var r2 CT_R
	if err := xmlutil.Unmarshal(data, &r2); err != nil {
		t.Fatalf("re-Unmarshal: %v", err)
	}
	if r2.RPr.FontHAnsi == nil || *r2.RPr.FontHAnsi != "Courier New" {
		t.Error("hAnsi should default to ascii value")
	}
}

func TestCT_RPr_Color_ThemeAttrs_RoundTrip(t *testing.T) {
	input := `<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:rPr>
			<w:color w:val="FF0000" w:themeColor="accent1" w:themeShade="BF" w:themeTint="99"/>
		</w:rPr>
		<w:t>text</w:t>
	</w:r>`
	var r CT_R
	if err := xmlutil.Unmarshal([]byte(input), &r); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if r.RPr == nil {
		t.Fatal("RPr nil")
	}
	if r.RPr.Color == nil || *r.RPr.Color != "FF0000" {
		t.Errorf("Color = %v", r.RPr.Color)
	}
	if r.RPr.ThemeColor == nil || *r.RPr.ThemeColor != "accent1" {
		t.Errorf("ThemeColor = %v", r.RPr.ThemeColor)
	}
	if r.RPr.ThemeShade == nil || *r.RPr.ThemeShade != "BF" {
		t.Errorf("ThemeShade = %v", r.RPr.ThemeShade)
	}
	if r.RPr.ThemeTint == nil || *r.RPr.ThemeTint != "99" {
		t.Errorf("ThemeTint = %v", r.RPr.ThemeTint)
	}

	// Round-trip.
	data, err := xmlutil.Marshal(&r, xmlutil.OOXML)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	var r2 CT_R
	if err := xmlutil.Unmarshal(data, &r2); err != nil {
		t.Fatalf("re-Unmarshal: %v", err)
	}
	if r2.RPr.ThemeColor == nil || *r2.RPr.ThemeColor != "accent1" {
		t.Error("ThemeColor lost during round-trip")
	}
	if r2.RPr.ThemeShade == nil || *r2.RPr.ThemeShade != "BF" {
		t.Error("ThemeShade lost during round-trip")
	}
	if r2.RPr.ThemeTint == nil || *r2.RPr.ThemeTint != "99" {
		t.Error("ThemeTint lost during round-trip")
	}
}

func TestCT_RPr_Marshal_ElementOrdering(t *testing.T) {
	// Per ECMA-376, rStyle must be the first child of rPr.
	// Full order: rStyle, rFonts, b, i, strike, u, color, sz
	r := CT_R{XMLName: xml.Name{Space: Ns, Local: "r"}}
	r.RPr = &CT_RPr{
		Bold:      new(true),
		Italic:    new(true),
		RunStyle:  new("Emphasis"),
		FontName:  new("Arial"),
		Color:     new("FF0000"),
		FontSize:  new("24"),
		Underline: new("single"),
	}
	r.AddText("test")

	data, err := xmlutil.Marshal(&r, xmlutil.OOXML)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	xml := string(data)

	// rStyle must appear before rFonts, b, i, strike, u, color, sz
	rStyleIdx := strings.Index(xml, "rStyle")
	rFontsIdx := strings.Index(xml, "rFonts")
	bIdx := strings.Index(xml, "<w:b")
	iIdx := strings.Index(xml, "<w:i")
	colorIdx := strings.Index(xml, "<w:color")
	szIdx := strings.Index(xml, "<w:sz")

	if rStyleIdx < 0 {
		t.Fatal("rStyle not found in output")
	}
	if rFontsIdx >= 0 && rStyleIdx > rFontsIdx {
		t.Error("rStyle must come before rFonts")
	}
	if bIdx >= 0 && rStyleIdx > bIdx {
		t.Error("rStyle must come before b")
	}
	if iIdx >= 0 && rStyleIdx > iIdx {
		t.Error("rStyle must come before i")
	}
	// rFonts before b
	if rFontsIdx >= 0 && bIdx >= 0 && rFontsIdx > bIdx {
		t.Error("rFonts must come before b")
	}
	// color before sz
	if colorIdx >= 0 && szIdx >= 0 && colorIdx > szIdx {
		t.Error("color must come before sz")
	}
}
