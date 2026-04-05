package wml

import (
	"encoding/xml"
	"strings"
	"testing"

	"github.com/ieshan/go-ooxml/xmlutil"
)

func TestCT_SectPr_Unmarshal(t *testing.T) {
	input := `<w:sectPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:pgSz w:w="12240" w:h="15840"/>
		<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
		<w:cols w:space="720"/>
	</w:sectPr>`
	var sp CT_SectPr
	if err := xmlutil.Unmarshal([]byte(input), &sp); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if sp.PgSz == nil || sp.PgSz.W != "12240" {
		t.Error("pgSz")
	}
	if sp.PgMar == nil || sp.PgMar.Top != "1440" {
		t.Error("pgMar")
	}
	if sp.Cols == nil {
		t.Error("cols nil")
	}
}

func TestCT_Columns_MultiColumn(t *testing.T) {
	input := `<w:cols xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:num="2" w:space="720" w:equalWidth="0">
		<w:col w:w="4000" w:space="720"/>
		<w:col w:w="4000"/>
	</w:cols>`
	var cols CT_Columns
	if err := xmlutil.Unmarshal([]byte(input), &cols); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if cols.Num == nil || *cols.Num != 2 {
		t.Error("num")
	}
	if len(cols.Cols) != 2 {
		t.Errorf("cols = %d", len(cols.Cols))
	}
}

func TestCT_SectPr_RoundTrip(t *testing.T) {
	sp := CT_SectPr{
		XMLName: xml.Name{Space: Ns, Local: "sectPr"},
		PgSz:    &CT_PgSz{W: "12240", H: "15840"},
		PgMar:   &CT_PgMar{Top: "1440", Right: "1440", Bottom: "1440", Left: "1440"},
	}
	data, err := xmlutil.Marshal(&sp, xmlutil.OOXML)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	var sp2 CT_SectPr
	if err := xmlutil.Unmarshal(data, &sp2); err != nil {
		t.Fatalf("Unmarshal round-trip: %v", err)
	}
	if sp2.PgSz == nil || sp2.PgSz.W != "12240" {
		t.Error("round-trip pgSz")
	}
}

func TestCT_SectPr_TypeAndExtra(t *testing.T) {
	input := `<w:sectPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:type w:val="continuous"/>
		<w:headerReference w:type="default" r:id="rId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
	</w:sectPr>`
	var sp CT_SectPr
	if err := xmlutil.Unmarshal([]byte(input), &sp); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if sp.Type == nil || *sp.Type != "continuous" {
		t.Errorf("type = %v", sp.Type)
	}
	if len(sp.HeaderRefs) != 1 {
		t.Errorf("HeaderRefs = %d, want 1", len(sp.HeaderRefs))
	}
}

func TestCT_Columns_Sep(t *testing.T) {
	input := `<w:cols xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:num="2" w:sep="1" w:equalWidth="1"/>`
	var cols CT_Columns
	if err := xmlutil.Unmarshal([]byte(input), &cols); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if cols.Sep == nil || !*cols.Sep {
		t.Error("sep")
	}
	if cols.EqualWidth == nil || !*cols.EqualWidth {
		t.Error("equalWidth")
	}
}

func TestCT_SectPr_AllFields(t *testing.T) {
	input := `<w:sectPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:type w:val="continuous"/>
		<w:pgSz w:w="12240" w:h="15840" w:orient="portrait"/>
		<w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="360" w:footer="360" w:gutter="0"/>
		<w:cols w:num="2" w:space="720"/>
	</w:sectPr>`
	var sp CT_SectPr
	if err := xmlutil.Unmarshal([]byte(input), &sp); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if sp.Type == nil || *sp.Type != "continuous" {
		t.Error("Type")
	}
	if sp.PgSz == nil || sp.PgSz.Orient != "portrait" {
		t.Error("PgSz.Orient")
	}
	if sp.PgMar == nil || sp.PgMar.Header != "360" || sp.PgMar.Footer != "360" || sp.PgMar.Gutter != "0" {
		t.Error("PgMar fields")
	}
	if sp.Cols == nil || sp.Cols.Num == nil || *sp.Cols.Num != 2 {
		t.Error("Cols")
	}
}

func TestCT_SectPr_Empty(t *testing.T) {
	input := `<w:sectPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`
	var sp CT_SectPr
	if err := xmlutil.Unmarshal([]byte(input), &sp); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if sp.PgSz != nil {
		t.Error("PgSz should be nil")
	}
	if sp.PgMar != nil {
		t.Error("PgMar should be nil")
	}
	if sp.Cols != nil {
		t.Error("Cols should be nil")
	}
	if sp.Type != nil {
		t.Error("Type should be nil")
	}
}

func TestCT_Columns_IndividualCols_RoundTrip(t *testing.T) {
	num := 3
	space := "360"
	eq := false
	cols := CT_Columns{
		Num:        &num,
		Space:      &space,
		EqualWidth: &eq,
		Cols: []*CT_Column{
			{W: "2000", Space: "360"},
			{W: "3000", Space: "360"},
			{W: "2000"},
		},
	}
	sp := CT_SectPr{
		XMLName: xml.Name{Space: Ns, Local: "sectPr"},
		Cols:    &cols,
	}
	data, err := xmlutil.Marshal(&sp, xmlutil.OOXML)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	var sp2 CT_SectPr
	if err := xmlutil.Unmarshal(data, &sp2); err != nil {
		t.Fatalf("re-Unmarshal: %v", err)
	}
	if sp2.Cols == nil {
		t.Fatal("Cols nil")
	}
	if sp2.Cols.Num == nil || *sp2.Cols.Num != 3 {
		t.Error("Num")
	}
	if len(sp2.Cols.Cols) != 3 {
		t.Errorf("col count = %d, want 3", len(sp2.Cols.Cols))
	}
	if sp2.Cols.Cols[0].W != "2000" {
		t.Errorf("col[0].W = %q", sp2.Cols.Cols[0].W)
	}
	if sp2.Cols.EqualWidth == nil || *sp2.Cols.EqualWidth {
		t.Error("EqualWidth should be false")
	}
}

func TestCT_SectPr_RoundTrip_WithType(t *testing.T) {
	contType := "nextPage"
	sp := CT_SectPr{
		XMLName: xml.Name{Space: Ns, Local: "sectPr"},
		Type:    &contType,
		PgSz:    &CT_PgSz{W: "15840", H: "12240", Orient: "landscape"},
		PgMar:   &CT_PgMar{Top: "720", Bottom: "720", Left: "1440", Right: "1440"},
	}
	data, err := xmlutil.Marshal(&sp, xmlutil.OOXML)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	var sp2 CT_SectPr
	if err := xmlutil.Unmarshal(data, &sp2); err != nil {
		t.Fatalf("re-Unmarshal: %v", err)
	}
	if sp2.Type == nil || *sp2.Type != "nextPage" {
		t.Error("Type lost")
	}
	if sp2.PgSz == nil || sp2.PgSz.Orient != "landscape" {
		t.Error("PgSz.Orient lost")
	}
}

func TestCT_Columns_UnknownChild(t *testing.T) {
	// Columns with unknown child element should skip gracefully
	input := `<w:sectPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:cols w:num="2">
			<w:col w:w="3000"/>
			<w:unknownColChild/>
			<w:col w:w="3000"/>
		</w:cols>
	</w:sectPr>`
	var sp CT_SectPr
	if err := xmlutil.Unmarshal([]byte(input), &sp); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if sp.Cols == nil || len(sp.Cols.Cols) != 2 {
		t.Errorf("Cols = %v", sp.Cols)
	}
}

func TestCT_SectPr_MarshalXML_AllFields(t *testing.T) {
	// Exercise all branches in CT_SectPr.MarshalXML: type, pgSz, pgMar, cols, extra
	// Use a round-trip through unmarshal to get the Extra field populated
	input := `<w:sectPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
		<w:type w:val="continuous"/>
		<w:pgSz w:w="12240" w:h="15840"/>
		<w:pgMar w:top="720" w:bottom="720" w:left="720" w:right="720"/>
		<w:cols w:num="1"/>
		<w:headerReference w:type="default" r:id="rId2"/>
	</w:sectPr>`
	var sp CT_SectPr
	if err := xmlutil.Unmarshal([]byte(input), &sp); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	data, err := xmlutil.Marshal(&sp, xmlutil.OOXML)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	var sp2 CT_SectPr
	if err := xmlutil.Unmarshal(data, &sp2); err != nil {
		t.Fatalf("re-Unmarshal: %v", err)
	}
	if sp2.Type == nil || *sp2.Type != "continuous" {
		t.Error("Type")
	}
	if sp2.PgSz == nil {
		t.Error("PgSz")
	}
	if sp2.PgMar == nil {
		t.Error("PgMar")
	}
	if sp2.Cols == nil {
		t.Error("Cols")
	}
	if len(sp2.HeaderRefs) != 1 {
		t.Errorf("HeaderRefs = %d, want 1", len(sp2.HeaderRefs))
	}
}

func TestCT_Columns_Sep_True_Marshal(t *testing.T) {
	// Test CT_Columns.MarshalXML with Sep=true and EqualWidth=true
	num := 2
	space := "720"
	sep := true
	eq := true
	cols := CT_Columns{
		Num:        &num,
		Space:      &space,
		Sep:        &sep,
		EqualWidth: &eq,
	}
	sp := CT_SectPr{
		XMLName: xml.Name{Space: Ns, Local: "sectPr"},
		Cols:    &cols,
	}
	data, err := xmlutil.Marshal(&sp, xmlutil.OOXML)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	var sp2 CT_SectPr
	if err := xmlutil.Unmarshal(data, &sp2); err != nil {
		t.Fatalf("re-Unmarshal: %v", err)
	}
	if sp2.Cols == nil {
		t.Fatal("Cols nil")
	}
	if sp2.Cols.Num == nil || *sp2.Cols.Num != 2 {
		t.Error("Num")
	}
	if sp2.Cols.Space == nil || *sp2.Cols.Space != "720" {
		t.Error("Space")
	}
	if sp2.Cols.Sep == nil || !*sp2.Cols.Sep {
		t.Error("Sep should be true")
	}
	if sp2.Cols.EqualWidth == nil || !*sp2.Cols.EqualWidth {
		t.Error("EqualWidth should be true")
	}
}

func TestCT_SectPr_HdrFtrRef_RoundTrip(t *testing.T) {
	input := `<w:sectPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
		<w:type w:val="nextPage"/>
		<w:pgSz w:w="12240" w:h="15840"/>
		<w:headerReference w:type="default" r:id="rId2"/>
		<w:footerReference w:type="default" r:id="rId3"/>
	</w:sectPr>`
	var sp CT_SectPr
	if err := xmlutil.Unmarshal([]byte(input), &sp); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if len(sp.HeaderRefs) != 1 {
		t.Fatalf("HeaderRefs = %d, want 1", len(sp.HeaderRefs))
	}
	if sp.HeaderRefs[0].Type != "default" || sp.HeaderRefs[0].ID != "rId2" {
		t.Errorf("HeaderRef = %+v", sp.HeaderRefs[0])
	}
	if len(sp.FooterRefs) != 1 {
		t.Fatalf("FooterRefs = %d, want 1", len(sp.FooterRefs))
	}
	if sp.FooterRefs[0].Type != "default" || sp.FooterRefs[0].ID != "rId3" {
		t.Errorf("FooterRef = %+v", sp.FooterRefs[0])
	}
	// Round-trip through marshal
	data, err := xmlutil.Marshal(&sp, xmlutil.OOXML)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	var sp2 CT_SectPr
	if err := xmlutil.Unmarshal(data, &sp2); err != nil {
		t.Fatalf("re-Unmarshal: %v", err)
	}
	if len(sp2.HeaderRefs) != 1 {
		t.Errorf("HeaderRefs after round-trip = %d, want 1", len(sp2.HeaderRefs))
	}
	if len(sp2.FooterRefs) != 1 {
		t.Errorf("FooterRefs after round-trip = %d, want 1", len(sp2.FooterRefs))
	}
}

func TestCT_Columns_Sep_False_RoundTrip(t *testing.T) {
	sep := false
	eq := true
	cols := CT_Columns{Sep: &sep, EqualWidth: &eq}
	sp := CT_SectPr{
		XMLName: xml.Name{Space: Ns, Local: "sectPr"},
		Cols:    &cols,
	}
	data, err := xmlutil.Marshal(&sp, xmlutil.OOXML)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	var sp2 CT_SectPr
	if err := xmlutil.Unmarshal(data, &sp2); err != nil {
		t.Fatalf("re-Unmarshal: %v", err)
	}
	if sp2.Cols == nil {
		t.Fatal("Cols nil")
	}
	if sp2.Cols.Sep == nil || *sp2.Cols.Sep {
		t.Error("Sep should be false")
	}
	if sp2.Cols.EqualWidth == nil || !*sp2.Cols.EqualWidth {
		t.Error("EqualWidth should be true")
	}
}

func TestCT_SectPr_Marshal_NoCorruptedNamespaces(t *testing.T) {
	// Go's encoding/xml generates synthetic _xmlns attributes for namespace-
	// qualified struct tag attributes. Verify our marshal path avoids this.
	sp := &CT_SectPr{
		XMLName: xml.Name{Space: Ns, Local: "sectPr"},
		PgSz:    &CT_PgSz{W: "12240", H: "15840", Orient: "landscape"},
		PgMar:   &CT_PgMar{Top: "1440", Right: "1440", Bottom: "1440", Left: "1440"},
	}
	num := 2
	space := "720"
	eq := true
	sep := false
	sp.Cols = &CT_Columns{
		Num: &num, Space: &space, EqualWidth: &eq, Sep: &sep,
		Cols: []*CT_Column{{W: "4320", Space: "720"}},
	}

	data, err := xmlutil.Marshal(sp, xmlutil.OOXML)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	out := string(data)

	if strings.Contains(out, "_xmlns") {
		t.Errorf("output contains _xmlns corruption: %s", out)
	}
	if strings.Contains(out, "main:") {
		t.Errorf("output contains main: prefix instead of w: %s", out)
	}
	if !strings.Contains(out, "w:w=") {
		t.Errorf("output missing w:w attribute: %s", out)
	}
	if !strings.Contains(out, "w:orient=") {
		t.Errorf("output missing w:orient attribute: %s", out)
	}
}
