package wml

import (
	"encoding/xml"
	"testing"

	"github.com/ieshan/go-ooxml/xmlutil"
)

func TestCT_RPrChange_Unmarshal(t *testing.T) {
	input := `<w:rPrChange xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:id="1" w:author="Joe" w:date="2024-01-01T00:00:00Z">
		<w:rPr><w:b/></w:rPr>
	</w:rPrChange>`
	var rpc CT_RPrChange
	if err := xmlutil.Unmarshal([]byte(input), &rpc); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if rpc.ID != 1 || rpc.Author != "Joe" {
		t.Errorf("attrs: %+v", rpc)
	}
	if rpc.RPr == nil || rpc.RPr.Bold == nil || *rpc.RPr.Bold != true {
		t.Error("rPr not parsed")
	}
}

func TestCT_RPrChange_RoundTrip(t *testing.T) {
	bold := true
	rpc := CT_RPrChange{
		XMLName: xml.Name{Space: Ns, Local: "rPrChange"},
		ID:      1,
		Author:  "Test",
		RPr:     &CT_RPr{Bold: &bold},
	}
	data, err := xmlutil.Marshal(&rpc, xmlutil.OOXML)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	var rpc2 CT_RPrChange
	if err := xmlutil.Unmarshal(data, &rpc2); err != nil {
		t.Fatalf("re-Unmarshal: %v", err)
	}
	if rpc2.RPr == nil || rpc2.RPr.Bold == nil || *rpc2.RPr.Bold != true {
		t.Error("lost rPr")
	}
}

func TestCT_PPrChange_Unmarshal(t *testing.T) {
	input := `<w:pPrChange xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:id="2" w:author="Jane">
		<w:pPr><w:jc w:val="center"/></w:pPr>
	</w:pPrChange>`
	var ppc CT_PPrChange
	if err := xmlutil.Unmarshal([]byte(input), &ppc); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if ppc.ID != 2 || ppc.Author != "Jane" {
		t.Error("attrs")
	}
	if ppc.PPr == nil || ppc.PPr.Alignment == nil || *ppc.PPr.Alignment != "center" {
		t.Error("pPr")
	}
}

func TestCT_PPrChange_MultipleProperties(t *testing.T) {
	input := `<w:pPrChange xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:id="10" w:author="Editor" w:date="2024-03-01T00:00:00Z">
		<w:pPr>
			<w:pStyle w:val="Heading1"/>
			<w:jc w:val="right"/>
		</w:pPr>
	</w:pPrChange>`
	var ppc CT_PPrChange
	if err := xmlutil.Unmarshal([]byte(input), &ppc); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if ppc.ID != 10 || ppc.Author != "Editor" || ppc.Date != "2024-03-01T00:00:00Z" {
		t.Errorf("attrs: id=%d author=%q date=%q", ppc.ID, ppc.Author, ppc.Date)
	}
	if ppc.PPr == nil {
		t.Fatal("PPr nil")
	}
	if ppc.PPr.Style == nil || *ppc.PPr.Style != "Heading1" {
		t.Error("Style")
	}
	if ppc.PPr.Alignment == nil || *ppc.PPr.Alignment != "right" {
		t.Error("Alignment")
	}
}

func TestCT_PPrChange_RoundTrip(t *testing.T) {
	style := "Normal"
	align := "left"
	ppc := CT_PPrChange{
		XMLName: xml.Name{Space: Ns, Local: "pPrChange"},
		ID:      3,
		Author:  "Writer",
		Date:    "2024-05-01T00:00:00Z",
		PPr:     &CT_PPr{Style: &style, Alignment: &align},
	}
	data, err := xmlutil.Marshal(&ppc, xmlutil.OOXML)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	var ppc2 CT_PPrChange
	if err := xmlutil.Unmarshal(data, &ppc2); err != nil {
		t.Fatalf("re-Unmarshal: %v", err)
	}
	if ppc2.ID != 3 || ppc2.Author != "Writer" {
		t.Errorf("attrs: id=%d author=%q", ppc2.ID, ppc2.Author)
	}
	if ppc2.PPr == nil || ppc2.PPr.Style == nil || *ppc2.PPr.Style != "Normal" {
		t.Error("PPr.Style lost")
	}
	if ppc2.PPr.Alignment == nil || *ppc2.PPr.Alignment != "left" {
		t.Error("PPr.Alignment lost")
	}
}

func TestCT_RPrChange_MultipleFmtFields(t *testing.T) {
	input := `<w:rPrChange xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:id="20" w:author="Rev" w:date="2024-04-01T00:00:00Z">
		<w:rPr>
			<w:b/>
			<w:i/>
			<w:color w:val="FF0000"/>
		</w:rPr>
	</w:rPrChange>`
	var rpc CT_RPrChange
	if err := xmlutil.Unmarshal([]byte(input), &rpc); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if rpc.RPr == nil {
		t.Fatal("RPr nil")
	}
	if rpc.RPr.Bold == nil || !*rpc.RPr.Bold {
		t.Error("Bold")
	}
	if rpc.RPr.Italic == nil || !*rpc.RPr.Italic {
		t.Error("Italic")
	}
	if rpc.RPr.Color == nil || *rpc.RPr.Color != "FF0000" {
		t.Error("Color")
	}
}

func TestCT_RPrChange_RoundTrip_WithDate(t *testing.T) {
	bold := true
	italic := true
	color := "0000FF"
	rpc := CT_RPrChange{
		XMLName: xml.Name{Space: Ns, Local: "rPrChange"},
		ID:      99,
		Author:  "Reviewer",
		Date:    "2024-07-04T00:00:00Z",
		RPr:     &CT_RPr{Bold: &bold, Italic: &italic, Color: &color},
	}
	data, err := xmlutil.Marshal(&rpc, xmlutil.OOXML)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	var rpc2 CT_RPrChange
	if err := xmlutil.Unmarshal(data, &rpc2); err != nil {
		t.Fatalf("re-Unmarshal: %v", err)
	}
	if rpc2.Date != "2024-07-04T00:00:00Z" {
		t.Errorf("Date = %q", rpc2.Date)
	}
	if rpc2.RPr == nil || rpc2.RPr.Color == nil || *rpc2.RPr.Color != "0000FF" {
		t.Error("Color lost")
	}
}

func TestCT_PPrChange_NoPPr(t *testing.T) {
	// PPrChange with no pPr child (unknown child → skipped)
	input := `<w:pPrChange xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:id="1" w:author="X">
		<w:unknownChild/>
	</w:pPrChange>`
	var ppc CT_PPrChange
	if err := xmlutil.Unmarshal([]byte(input), &ppc); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if ppc.PPr != nil {
		t.Error("PPr should be nil when no pPr child")
	}
}

func TestCT_RPrChange_UnknownChild(t *testing.T) {
	// RPrChange with unknown child should skip it
	input := `<w:rPrChange xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:id="5" w:author="X">
		<w:rPr><w:b/></w:rPr>
		<w:unknownChild/>
	</w:rPrChange>`
	var rpc CT_RPrChange
	if err := xmlutil.Unmarshal([]byte(input), &rpc); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if rpc.RPr == nil || rpc.RPr.Bold == nil || !*rpc.RPr.Bold {
		t.Error("RPr.Bold")
	}
}

func TestCT_PPrChange_UnknownChild(t *testing.T) {
	// PPrChange with unknown child should skip it
	input := `<w:pPrChange xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:id="6" w:author="Y">
		<w:pPr><w:jc w:val="center"/></w:pPr>
		<w:unknownPPrChangeChild/>
	</w:pPrChange>`
	var ppc CT_PPrChange
	if err := xmlutil.Unmarshal([]byte(input), &ppc); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if ppc.PPr == nil || ppc.PPr.Alignment == nil || *ppc.PPr.Alignment != "center" {
		t.Error("PPr.Alignment")
	}
}

func TestCT_RPrChange_NilRPr(t *testing.T) {
	// RPrChange round-trip with nil RPr
	rpc := CT_RPrChange{
		XMLName: xml.Name{Space: Ns, Local: "rPrChange"},
		ID:      5,
		Author:  "NoRpr",
	}
	data, err := xmlutil.Marshal(&rpc, xmlutil.OOXML)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	var rpc2 CT_RPrChange
	if err := xmlutil.Unmarshal(data, &rpc2); err != nil {
		t.Fatalf("re-Unmarshal: %v", err)
	}
	if rpc2.RPr != nil {
		t.Error("RPr should be nil")
	}
}
