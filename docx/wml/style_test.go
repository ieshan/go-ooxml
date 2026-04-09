package wml

import (
	"encoding/xml"
	"strings"
	"testing"

	"github.com/ieshan/go-ooxml/xmlutil"
)

func TestCT_Styles_Unmarshal(t *testing.T) {
	input := `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:style w:type="paragraph" w:styleId="Heading1">
			<w:name w:val="heading 1"/>
			<w:basedOn w:val="Normal"/>
			<w:pPr><w:jc w:val="center"/></w:pPr>
			<w:rPr><w:b/></w:rPr>
		</w:style>
		<w:style w:type="character" w:styleId="Strong" w:default="1">
			<w:name w:val="Strong"/>
			<w:rPr><w:b/></w:rPr>
		</w:style>
	</w:styles>`
	var styles CT_Styles
	if err := xmlutil.Unmarshal([]byte(input), &styles); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if len(styles.Styles) != 2 {
		t.Fatalf("count = %d", len(styles.Styles))
	}
	s := styles.Styles[0]
	if s.Type != "paragraph" || s.StyleID != "Heading1" {
		t.Errorf("style 0: type=%q id=%q", s.Type, s.StyleID)
	}
	if s.Name != "heading 1" {
		t.Errorf("name = %q", s.Name)
	}
	if s.BasedOn == nil || *s.BasedOn != "Normal" {
		t.Error("basedOn")
	}
	if s.PPr == nil || s.PPr.Alignment == nil {
		t.Error("pPr")
	}
	if s.RPr == nil || s.RPr.Bold == nil {
		t.Error("rPr")
	}
	s2 := styles.Styles[1]
	if s2.Default == nil || *s2.Default != true {
		t.Error("default")
	}
}

func TestCT_Styles_RoundTrip(t *testing.T) {
	styles := CT_Styles{XMLName: xml.Name{Space: Ns, Local: "styles"}}
	bold := true
	s := &CT_Style{
		XMLName: xml.Name{Space: Ns, Local: "style"},
		Type:    "paragraph", StyleID: "Normal", Name: "Normal",
		RPr: &CT_RPr{Bold: &bold},
	}
	styles.Styles = append(styles.Styles, s)
	data, err := xmlutil.Marshal(&styles, xmlutil.OOXML)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	var styles2 CT_Styles
	if err := xmlutil.Unmarshal(data, &styles2); err != nil {
		t.Fatalf("Unmarshal round-trip: %v", err)
	}
	if len(styles2.Styles) != 1 || styles2.Styles[0].Name != "Normal" {
		t.Error("round-trip")
	}
}

func TestCT_Style_ExtraPreserved(t *testing.T) {
	input := `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:docDefaults><w:rPrDefault/></w:docDefaults>
		<w:style w:type="paragraph" w:styleId="Normal">
			<w:name w:val="Normal"/>
			<w:qFormat/>
		</w:style>
	</w:styles>`
	var styles CT_Styles
	if err := xmlutil.Unmarshal([]byte(input), &styles); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if len(styles.Extra) != 1 {
		t.Errorf("CT_Styles extra = %d, want 1", len(styles.Extra))
	}
	if len(styles.Styles) != 1 {
		t.Fatalf("styles count = %d", len(styles.Styles))
	}
	// qFormat should land in Extra on the style
	if len(styles.Styles[0].Extra) != 1 {
		t.Errorf("CT_Style extra = %d, want 1", len(styles.Styles[0].Extra))
	}
}

func TestCT_Style_DefaultAttribute(t *testing.T) {
	input := `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:style w:type="paragraph" w:styleId="Normal" w:default="1">
			<w:name w:val="Normal"/>
		</w:style>
	</w:styles>`
	var styles CT_Styles
	if err := xmlutil.Unmarshal([]byte(input), &styles); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if len(styles.Styles) != 1 {
		t.Fatal("styles count")
	}
	s := styles.Styles[0]
	if s.Default == nil || !*s.Default {
		t.Error("default attribute not parsed")
	}
	// Round-trip should preserve default=1
	data, err := xmlutil.Marshal(&styles, xmlutil.OOXML)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	var styles2 CT_Styles
	if err := xmlutil.Unmarshal(data, &styles2); err != nil {
		t.Fatalf("re-Unmarshal: %v", err)
	}
	if styles2.Styles[0].Default == nil || !*styles2.Styles[0].Default {
		t.Error("default lost in round-trip")
	}
}

func TestCT_Style_NoOptionalFields(t *testing.T) {
	input := `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:style w:type="character" w:styleId="DefaultFont">
			<w:name w:val="Default Font"/>
		</w:style>
	</w:styles>`
	var styles CT_Styles
	if err := xmlutil.Unmarshal([]byte(input), &styles); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if len(styles.Styles) != 1 {
		t.Fatal("styles count")
	}
	s := styles.Styles[0]
	if s.BasedOn != nil {
		t.Error("BasedOn should be nil")
	}
	if s.PPr != nil {
		t.Error("PPr should be nil")
	}
	if s.RPr != nil {
		t.Error("RPr should be nil")
	}
	if s.Default != nil {
		t.Error("Default should be nil")
	}
}

func TestCT_Styles_Empty(t *testing.T) {
	input := `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:styles>`
	var styles CT_Styles
	if err := xmlutil.Unmarshal([]byte(input), &styles); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if len(styles.Styles) != 0 {
		t.Errorf("styles = %d, want 0", len(styles.Styles))
	}
	// Marshal empty collection
	data, err := xmlutil.Marshal(&styles, xmlutil.OOXML)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	var styles2 CT_Styles
	if err := xmlutil.Unmarshal(data, &styles2); err != nil {
		t.Fatalf("re-Unmarshal: %v", err)
	}
	if len(styles2.Styles) != 0 {
		t.Errorf("styles after round-trip = %d, want 0", len(styles2.Styles))
	}
}

func TestCT_Styles_RoundTrip_PreservesExtra(t *testing.T) {
	input := `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:docDefaults><w:rPrDefault/></w:docDefaults>
		<w:style w:type="paragraph" w:styleId="Normal">
			<w:name w:val="Normal"/>
			<w:qFormat/>
		</w:style>
	</w:styles>`
	var styles CT_Styles
	if err := xmlutil.Unmarshal([]byte(input), &styles); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	data, err := xmlutil.Marshal(&styles, xmlutil.OOXML)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	var styles2 CT_Styles
	if err := xmlutil.Unmarshal(data, &styles2); err != nil {
		t.Fatalf("re-Unmarshal: %v", err)
	}
	if len(styles2.Extra) != 1 {
		t.Errorf("Extra = %d, want 1", len(styles2.Extra))
	}
	if len(styles2.Styles) != 1 {
		t.Errorf("Styles = %d, want 1", len(styles2.Styles))
	}
	if len(styles2.Styles[0].Extra) != 1 {
		t.Errorf("Style.Extra = %d, want 1", len(styles2.Styles[0].Extra))
	}
}

func TestCT_Style_Extra_RoundTrip_Marshal(t *testing.T) {
	// Style with Extra elements — exercise Extra marshal path
	input := `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:style w:type="paragraph" w:styleId="Heading1">
			<w:name w:val="heading 1"/>
			<w:basedOn w:val="Normal"/>
			<w:pPr><w:jc w:val="center"/></w:pPr>
			<w:rPr><w:b/></w:rPr>
			<w:qFormat/>
			<w:pPrDefault/>
		</w:style>
	</w:styles>`
	var styles CT_Styles
	if err := xmlutil.Unmarshal([]byte(input), &styles); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if len(styles.Styles[0].Extra) != 2 {
		t.Fatalf("Extra = %d, want 2", len(styles.Styles[0].Extra))
	}
	// Marshal to exercise Extra loop in CT_Style.MarshalXML
	data, err := xmlutil.Marshal(&styles, xmlutil.OOXML)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	var styles2 CT_Styles
	if err := xmlutil.Unmarshal(data, &styles2); err != nil {
		t.Fatalf("re-Unmarshal: %v", err)
	}
	if len(styles2.Styles[0].Extra) != 2 {
		t.Errorf("Extra = %d, want 2", len(styles2.Styles[0].Extra))
	}
}

func TestCT_Style_DefaultFalse(t *testing.T) {
	// Styles with default="0" should parse as false, not nil
	input := `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:style w:type="paragraph" w:styleId="Body" w:default="0">
			<w:name w:val="Body Text"/>
		</w:style>
	</w:styles>`
	var styles CT_Styles
	if err := xmlutil.Unmarshal([]byte(input), &styles); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	s := styles.Styles[0]
	if s.Default == nil {
		t.Fatal("Default should not be nil for explicit default=0")
	}
	if *s.Default {
		t.Error("Default should be false")
	}
}

func TestCT_Style_AllFields_RoundTrip(t *testing.T) {
	// Exercise all branches in CT_Style.MarshalXML: type, styleID, default, name, basedOn, pPr, rPr, extra
	input := `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:style w:type="paragraph" w:styleId="Heading1" w:default="1">
			<w:name w:val="heading 1"/>
			<w:basedOn w:val="Normal"/>
			<w:pPr><w:jc w:val="left"/></w:pPr>
			<w:rPr><w:b/><w:sz w:val="28"/></w:rPr>
			<w:qFormat/>
		</w:style>
	</w:styles>`
	var styles CT_Styles
	if err := xmlutil.Unmarshal([]byte(input), &styles); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	data, err := xmlutil.Marshal(&styles, xmlutil.OOXML)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	var styles2 CT_Styles
	if err := xmlutil.Unmarshal(data, &styles2); err != nil {
		t.Fatalf("re-Unmarshal: %v", err)
	}
	s := styles2.Styles[0]
	if s.Type != "paragraph" || s.StyleID != "Heading1" {
		t.Errorf("type=%q styleID=%q", s.Type, s.StyleID)
	}
	if s.Default == nil || !*s.Default {
		t.Error("Default")
	}
	if s.Name != "heading 1" {
		t.Errorf("Name = %q", s.Name)
	}
	if s.BasedOn == nil || *s.BasedOn != "Normal" {
		t.Error("BasedOn")
	}
	if s.PPr == nil || s.PPr.Alignment == nil {
		t.Error("PPr")
	}
	if s.RPr == nil || s.RPr.Bold == nil || s.RPr.FontSize == nil {
		t.Error("RPr")
	}
	if len(s.Extra) != 1 {
		t.Errorf("Extra = %d, want 1", len(s.Extra))
	}
}

func TestCT_Styles_MarshalXML_WithExtra(t *testing.T) {
	// CT_Styles.MarshalXML should emit Extra before styles
	input := `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:docDefaults>
			<w:rPrDefault><w:rPr><w:sz w:val="24"/></w:rPr></w:rPrDefault>
		</w:docDefaults>
		<w:latentStyles/>
		<w:style w:type="paragraph" w:styleId="Normal">
			<w:name w:val="Normal"/>
		</w:style>
	</w:styles>`
	var styles CT_Styles
	if err := xmlutil.Unmarshal([]byte(input), &styles); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if len(styles.Extra) != 2 {
		t.Fatalf("Extra = %d, want 2", len(styles.Extra))
	}
	data, err := xmlutil.Marshal(&styles, xmlutil.OOXML)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	var styles2 CT_Styles
	if err := xmlutil.Unmarshal(data, &styles2); err != nil {
		t.Fatalf("re-Unmarshal: %v", err)
	}
	if len(styles2.Extra) != 2 {
		t.Errorf("Extra = %d, want 2", len(styles2.Extra))
	}
	if len(styles2.Styles) != 1 {
		t.Errorf("Styles = %d, want 1", len(styles2.Styles))
	}
}

func TestCT_Style_Next_RoundTrip(t *testing.T) {
	next := "Normal"
	style := CT_Style{
		XMLName: xml.Name{Space: Ns, Local: "style"},
		Type:    "paragraph",
		StyleID: "Heading1",
		Name:    "heading 1",
		Next:    &next,
	}

	data, err := xmlutil.Marshal(&style, xmlutil.OOXML)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	if !strings.Contains(string(data), "next") {
		t.Errorf("XML should contain next element: %s", data)
	}

	var style2 CT_Style
	if err := xmlutil.Unmarshal(data, &style2); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if style2.Next == nil || *style2.Next != "Normal" {
		t.Errorf("Next = %v, want Normal", style2.Next)
	}
}
