package wml

import (
	"encoding/xml"
	"testing"

	"github.com/ieshan/go-ooxml/xmlutil"
)

func TestCT_Tbl_Unmarshal(t *testing.T) {
	input := `<w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:tblPr><w:tblStyle w:val="TableGrid"/></w:tblPr>
		<w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>
		<w:tr>
			<w:tc><w:p><w:r><w:t>Cell 1</w:t></w:r></w:p></w:tc>
			<w:tc><w:p><w:r><w:t>Cell 2</w:t></w:r></w:p></w:tc>
		</w:tr>
	</w:tbl>`
	var tbl CT_Tbl
	if err := xmlutil.Unmarshal([]byte(input), &tbl); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if len(tbl.Rows) != 1 {
		t.Errorf("rows = %d", len(tbl.Rows))
	}
	if len(tbl.Rows[0].Cells) != 2 {
		t.Errorf("cells = %d", len(tbl.Rows[0].Cells))
	}
}

func TestCT_Tc_BlockContent(t *testing.T) {
	input := `<w:tc xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:p><w:r><w:t>para 1</w:t></w:r></w:p>
		<w:p><w:r><w:t>para 2</w:t></w:r></w:p>
	</w:tc>`
	var tc CT_Tc
	xmlutil.Unmarshal([]byte(input), &tc)
	paraCount := 0
	for _, c := range tc.Content {
		if c.Paragraph != nil {
			paraCount++
		}
	}
	if paraCount != 2 {
		t.Errorf("paragraphs = %d, want 2", paraCount)
	}
}

func TestCT_Tbl_RoundTrip(t *testing.T) {
	// Build a table, marshal, unmarshal, verify
	tbl := CT_Tbl{XMLName: xml.Name{Space: Ns, Local: "tbl"}}
	row := &CT_Row{XMLName: xml.Name{Space: Ns, Local: "tr"}}
	cell := &CT_Tc{XMLName: xml.Name{Space: Ns, Local: "tc"}}
	p := &CT_P{XMLName: xml.Name{Space: Ns, Local: "p"}}
	p.AddRun("test")
	cell.Content = append(cell.Content, BlockLevelContent{Paragraph: p})
	row.Cells = append(row.Cells, cell)
	tbl.Rows = append(tbl.Rows, row)

	data, err := xmlutil.Marshal(&tbl, xmlutil.OOXML)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	var tbl2 CT_Tbl
	if err := xmlutil.Unmarshal(data, &tbl2); err != nil {
		t.Fatalf("re-Unmarshal: %v", err)
	}
	if len(tbl2.Rows) != 1 || len(tbl2.Rows[0].Cells) != 1 {
		t.Error("structure lost")
	}
}

func TestCT_TcPr_CellMerge(t *testing.T) {
	input := `<w:tc xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:tcPr>
			<w:vMerge w:val="restart"/>
			<w:gridSpan w:val="2"/>
		</w:tcPr>
		<w:p/>
	</w:tc>`
	var tc CT_Tc
	xmlutil.Unmarshal([]byte(input), &tc)
	if tc.TcPr == nil {
		t.Fatal("TcPr nil")
	}
	if tc.TcPr.VMerge == nil || tc.TcPr.VMerge.Val != "restart" {
		t.Error("VMerge")
	}
	if tc.TcPr.GridSpan == nil || *tc.TcPr.GridSpan != 2 {
		t.Error("GridSpan")
	}
}

func TestCT_Tbl_EmptyTable(t *testing.T) {
	input := `<w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:tbl>`
	var tbl CT_Tbl
	if err := xmlutil.Unmarshal([]byte(input), &tbl); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if len(tbl.Rows) != 0 {
		t.Errorf("rows = %d, want 0", len(tbl.Rows))
	}
	if tbl.TblPr != nil {
		t.Error("TblPr should be nil")
	}
}

func TestCT_Tbl_WithTblPrStyleAndWidth(t *testing.T) {
	input := `<w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:tblPr>
			<w:tblStyle w:val="TableGrid"/>
			<w:tblW w:w="5000" w:type="pct"/>
			<w:jc w:val="center"/>
		</w:tblPr>
	</w:tbl>`
	var tbl CT_Tbl
	if err := xmlutil.Unmarshal([]byte(input), &tbl); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if tbl.TblPr == nil {
		t.Fatal("TblPr nil")
	}
	if tbl.TblPr.Style == nil || *tbl.TblPr.Style != "TableGrid" {
		t.Error("Style")
	}
	if tbl.TblPr.Width == nil || tbl.TblPr.Width.W != "5000" || tbl.TblPr.Width.Type != "pct" {
		t.Error("Width")
	}
	if tbl.TblPr.Jc == nil || *tbl.TblPr.Jc != "center" {
		t.Error("Jc")
	}
}

func TestCT_TblPr_MarshalXML(t *testing.T) {
	pr := CT_TblPr{
		Style: new("TableGrid"),
		Width: &CT_TblWidth{W: "5000", Type: "pct"},
		Jc:    new("center"),
	}
	// Wrap in a table to exercise TblPr marshal round-trip
	tbl := CT_Tbl{XMLName: xml.Name{Space: Ns, Local: "tbl"}, TblPr: &pr}
	data, err := xmlutil.Marshal(&tbl, xmlutil.OOXML)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	var tbl2 CT_Tbl
	if err := xmlutil.Unmarshal(data, &tbl2); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	pr2 := tbl2.TblPr
	if pr2 == nil {
		t.Fatal("TblPr nil")
	}
	if pr2.Style == nil || *pr2.Style != "TableGrid" {
		t.Error("Style")
	}
	if pr2.Width == nil || pr2.Width.W != "5000" {
		t.Error("Width")
	}
	if pr2.Jc == nil || *pr2.Jc != "center" {
		t.Error("Jc")
	}
}

func TestCT_TblGrid_MarshalXML(t *testing.T) {
	grid := CT_TblGrid{
		Cols: []*CT_TblGridCol{
			{W: "4000"},
			{W: "3000"},
		},
	}
	tbl := CT_Tbl{XMLName: xml.Name{Space: Ns, Local: "tbl"}, TblGrid: &grid}
	data, err := xmlutil.Marshal(&tbl, xmlutil.OOXML)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	var tbl2 CT_Tbl
	if err := xmlutil.Unmarshal(data, &tbl2); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if tbl2.TblGrid == nil {
		t.Fatal("TblGrid nil")
	}
	if len(tbl2.TblGrid.Cols) != 2 {
		t.Errorf("cols = %d, want 2", len(tbl2.TblGrid.Cols))
	}
	if tbl2.TblGrid.Cols[0].W != "4000" {
		t.Errorf("col[0].W = %q", tbl2.TblGrid.Cols[0].W)
	}
}

func TestCT_TcPr_Shading(t *testing.T) {
	input := `<w:tc xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:tcPr>
			<w:shd w:val="clear" w:color="auto" w:fill="FF0000"/>
		</w:tcPr>
		<w:p/>
	</w:tc>`
	var tc CT_Tc
	if err := xmlutil.Unmarshal([]byte(input), &tc); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if tc.TcPr == nil || tc.TcPr.Shading == nil {
		t.Fatal("Shading nil")
	}
	if tc.TcPr.Shading.Val != "clear" {
		t.Error("Val")
	}
	if tc.TcPr.Shading.Color != "auto" {
		t.Error("Color")
	}
	if tc.TcPr.Shading.Fill != "FF0000" {
		t.Error("Fill")
	}
}

func TestCT_TcPr_MarshalXML(t *testing.T) {
	pr := CT_TcPr{
		Width:    &CT_TblWidth{W: "2000", Type: "dxa"},
		GridSpan: new(2),
		VMerge:   &CT_VMerge{Val: "restart"},
		Shading:  &CT_Shd{Val: "clear", Color: "auto", Fill: "FF0000"},
	}
	// Wrap in a cell inside a table to round-trip TcPr
	tc := CT_Tc{XMLName: xml.Name{Space: Ns, Local: "tc"}, TcPr: &pr}
	p := &CT_P{XMLName: xml.Name{Space: Ns, Local: "p"}}
	tc.Content = append(tc.Content, BlockLevelContent{Paragraph: p})
	row := CT_Row{XMLName: xml.Name{Space: Ns, Local: "tr"}, Cells: []*CT_Tc{&tc}}
	tbl := CT_Tbl{XMLName: xml.Name{Space: Ns, Local: "tbl"}, Rows: []*CT_Row{&row}}
	data, err := xmlutil.Marshal(&tbl, xmlutil.OOXML)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	var tbl2 CT_Tbl
	if err := xmlutil.Unmarshal(data, &tbl2); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	pr2 := tbl2.Rows[0].Cells[0].TcPr
	if pr2 == nil {
		t.Fatal("TcPr nil after round-trip")
	}
	if pr2.Width == nil || pr2.Width.W != "2000" {
		t.Error("Width")
	}
	if pr2.GridSpan == nil || *pr2.GridSpan != 2 {
		t.Error("GridSpan")
	}
	if pr2.VMerge == nil || pr2.VMerge.Val != "restart" {
		t.Error("VMerge")
	}
	if pr2.Shading == nil || pr2.Shading.Fill != "FF0000" {
		t.Error("Shading")
	}
}

func TestCT_TcPr_VMerge_NoVal(t *testing.T) {
	// vMerge with no val attribute = continuation merge
	input := `<w:tc xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:tcPr><w:vMerge/></w:tcPr>
		<w:p/>
	</w:tc>`
	var tc CT_Tc
	if err := xmlutil.Unmarshal([]byte(input), &tc); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if tc.TcPr == nil || tc.TcPr.VMerge == nil {
		t.Fatal("VMerge nil")
	}
	if tc.TcPr.VMerge.Val != "" {
		t.Errorf("VMerge.Val = %q, want empty", tc.TcPr.VMerge.Val)
	}
}

func TestCT_Tc_NestedTable(t *testing.T) {
	input := `<w:tc xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:p><w:r><w:t>before</w:t></w:r></w:p>
		<w:tbl>
			<w:tr><w:tc><w:p><w:r><w:t>nested</w:t></w:r></w:p></w:tc></w:tr>
		</w:tbl>
		<w:p><w:r><w:t>after</w:t></w:r></w:p>
	</w:tc>`
	var tc CT_Tc
	if err := xmlutil.Unmarshal([]byte(input), &tc); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if len(tc.Content) != 3 {
		t.Fatalf("content = %d, want 3", len(tc.Content))
	}
	if tc.Content[1].Table == nil {
		t.Error("nested table not parsed")
	}
}

func TestCT_Tbl_RoundTrip_WithTblPr(t *testing.T) {
	tbl := CT_Tbl{XMLName: xml.Name{Space: Ns, Local: "tbl"}}
	tbl.TblPr = &CT_TblPr{
		Style: new("TableGrid"),
		Width: &CT_TblWidth{W: "5000", Type: "pct"},
	}
	tbl.TblGrid = &CT_TblGrid{
		Cols: []*CT_TblGridCol{{W: "5000"}},
	}
	row := &CT_Row{XMLName: xml.Name{Space: Ns, Local: "tr"}}
	cell := &CT_Tc{XMLName: xml.Name{Space: Ns, Local: "tc"}}
	p := &CT_P{XMLName: xml.Name{Space: Ns, Local: "p"}}
	p.AddRun("content")
	cell.Content = append(cell.Content, BlockLevelContent{Paragraph: p})
	row.Cells = append(row.Cells, cell)
	tbl.Rows = append(tbl.Rows, row)

	data, err := xmlutil.Marshal(&tbl, xmlutil.OOXML)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	var tbl2 CT_Tbl
	if err := xmlutil.Unmarshal(data, &tbl2); err != nil {
		t.Fatalf("re-Unmarshal: %v", err)
	}
	if tbl2.TblPr == nil || tbl2.TblPr.Style == nil || *tbl2.TblPr.Style != "TableGrid" {
		t.Error("TblPr/Style not preserved")
	}
	if tbl2.TblGrid == nil || len(tbl2.TblGrid.Cols) != 1 {
		t.Error("TblGrid not preserved")
	}
}

func TestCT_Row_WithTrPr(t *testing.T) {
	input := `<w:tr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:trPr><w:cantSplit/></w:trPr>
		<w:tc><w:p/></w:tc>
	</w:tr>`
	var row CT_Row
	if err := xmlutil.Unmarshal([]byte(input), &row); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if row.TrPr == nil {
		t.Error("TrPr should be preserved")
	}
	if len(row.Cells) != 1 {
		t.Errorf("cells = %d, want 1", len(row.Cells))
	}
}

func TestCT_TblPr_ExtraPreserved(t *testing.T) {
	input := `<w:tblPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:tblStyle w:val="Normal"/>
		<w:tblBorders><w:top w:val="single"/></w:tblBorders>
	</w:tblPr>`
	var pr CT_TblPr
	if err := xmlutil.Unmarshal([]byte(input), &pr); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if len(pr.Extra) != 1 {
		t.Errorf("Extra = %d, want 1", len(pr.Extra))
	}
}

func TestCT_TblPr_Extra_RoundTrip(t *testing.T) {
	// Table with Extra in TblPr should be preserved through marshal
	input := `<w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:tblPr>
			<w:tblStyle w:val="Grid"/>
			<w:tblBorders><w:top w:val="single"/></w:tblBorders>
		</w:tblPr>
		<w:tr><w:tc><w:p/></w:tc></w:tr>
	</w:tbl>`
	var tbl CT_Tbl
	if err := xmlutil.Unmarshal([]byte(input), &tbl); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if tbl.TblPr == nil || len(tbl.TblPr.Extra) != 1 {
		t.Fatalf("TblPr.Extra = %d, want 1", len(tbl.TblPr.Extra))
	}
	// Marshal and check Extra preserved
	data, err := xmlutil.Marshal(&tbl, xmlutil.OOXML)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	var tbl2 CT_Tbl
	if err := xmlutil.Unmarshal(data, &tbl2); err != nil {
		t.Fatalf("re-Unmarshal: %v", err)
	}
	if tbl2.TblPr == nil || len(tbl2.TblPr.Extra) != 1 {
		t.Errorf("TblPr.Extra lost, got %d", len(tbl2.TblPr.Extra))
	}
}

func TestCT_Tbl_Raw_RoundTrip(t *testing.T) {
	// Table with unknown element preserved in Raw
	input := `<w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:tr><w:tc><w:p/></w:tc></w:tr>
		<w:unknownElement/>
	</w:tbl>`
	var tbl CT_Tbl
	if err := xmlutil.Unmarshal([]byte(input), &tbl); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if len(tbl.Raw) != 1 {
		t.Fatalf("Raw = %d, want 1", len(tbl.Raw))
	}
	data, err := xmlutil.Marshal(&tbl, xmlutil.OOXML)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	var tbl2 CT_Tbl
	if err := xmlutil.Unmarshal(data, &tbl2); err != nil {
		t.Fatalf("re-Unmarshal: %v", err)
	}
	if len(tbl2.Raw) != 1 {
		t.Errorf("Raw lost, got %d", len(tbl2.Raw))
	}
}

func TestCT_Row_Raw_RoundTrip(t *testing.T) {
	// Row with raw (non-tc) elements preserved
	input := `<w:tr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:tc><w:p/></w:tc>
		<w:unknownRowChild/>
	</w:tr>`
	var row CT_Row
	if err := xmlutil.Unmarshal([]byte(input), &row); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if len(row.Raw) != 1 {
		t.Fatalf("Raw = %d, want 1", len(row.Raw))
	}
	if len(row.Cells) != 1 {
		t.Fatalf("Cells = %d, want 1", len(row.Cells))
	}
	// Marshal to exercise Raw path in CT_Row.MarshalXML
	tbl := CT_Tbl{XMLName: xml.Name{Space: Ns, Local: "tbl"}, Rows: []*CT_Row{&row}}
	data, err := xmlutil.Marshal(&tbl, xmlutil.OOXML)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	var tbl2 CT_Tbl
	if err := xmlutil.Unmarshal(data, &tbl2); err != nil {
		t.Fatalf("re-Unmarshal: %v", err)
	}
	if len(tbl2.Rows[0].Raw) != 1 {
		t.Errorf("Row.Raw lost, got %d", len(tbl2.Rows[0].Raw))
	}
}

func TestCT_Tc_Raw_Content_RoundTrip(t *testing.T) {
	// Cell with unknown block content preserved as Raw
	input := `<w:tc xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:p><w:r><w:t>text</w:t></w:r></w:p>
		<w:sdt><w:sdtContent><w:p/></w:sdtContent></w:sdt>
	</w:tc>`
	var tc CT_Tc
	if err := xmlutil.Unmarshal([]byte(input), &tc); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if len(tc.Content) != 2 {
		t.Fatalf("content = %d, want 2", len(tc.Content))
	}
	if tc.Content[1].Raw == nil {
		t.Fatal("Raw content nil")
	}
	// Marshal to exercise Raw path in CT_Tc.MarshalXML
	row := CT_Row{XMLName: xml.Name{Space: Ns, Local: "tr"}, Cells: []*CT_Tc{&tc}}
	tbl := CT_Tbl{XMLName: xml.Name{Space: Ns, Local: "tbl"}, Rows: []*CT_Row{&row}}
	data, err := xmlutil.Marshal(&tbl, xmlutil.OOXML)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	var tbl2 CT_Tbl
	if err := xmlutil.Unmarshal(data, &tbl2); err != nil {
		t.Fatalf("re-Unmarshal: %v", err)
	}
	cell2 := tbl2.Rows[0].Cells[0]
	if len(cell2.Content) != 2 {
		t.Fatalf("content = %d, want 2", len(cell2.Content))
	}
	if cell2.Content[1].Raw == nil {
		t.Error("Raw content lost")
	}
}

func TestCT_TcPr_Extra_RoundTrip(t *testing.T) {
	// Cell with TcPr containing extra elements
	input := `<w:tc xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:tcPr>
			<w:tcW w:w="2000" w:type="dxa"/>
			<w:noWrap/>
		</w:tcPr>
		<w:p/>
	</w:tc>`
	var tc CT_Tc
	if err := xmlutil.Unmarshal([]byte(input), &tc); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if tc.TcPr == nil || len(tc.TcPr.Extra) != 1 {
		t.Fatalf("TcPr.Extra = %d, want 1", len(tc.TcPr.Extra))
	}
	// Round-trip to exercise Extra path in CT_TcPr.MarshalXML
	row := CT_Row{XMLName: xml.Name{Space: Ns, Local: "tr"}, Cells: []*CT_Tc{&tc}}
	tbl := CT_Tbl{XMLName: xml.Name{Space: Ns, Local: "tbl"}, Rows: []*CT_Row{&row}}
	data, err := xmlutil.Marshal(&tbl, xmlutil.OOXML)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	var tbl2 CT_Tbl
	if err := xmlutil.Unmarshal(data, &tbl2); err != nil {
		t.Fatalf("re-Unmarshal: %v", err)
	}
	tcPr2 := tbl2.Rows[0].Cells[0].TcPr
	if tcPr2 == nil || len(tcPr2.Extra) != 1 {
		t.Errorf("TcPr.Extra lost, got %d", len(tcPr2.Extra))
	}
}

func TestCT_Row_Marshal_WithTrPr(t *testing.T) {
	// Test that TrPr is preserved through marshal in CT_Row.MarshalXML
	input := `<w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:tr>
			<w:trPr><w:cantSplit/><w:trHeight w:val="400"/></w:trPr>
			<w:tc><w:p><w:r><w:t>data</w:t></w:r></w:p></w:tc>
		</w:tr>
	</w:tbl>`
	var tbl CT_Tbl
	if err := xmlutil.Unmarshal([]byte(input), &tbl); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if tbl.Rows[0].TrPr == nil {
		t.Fatal("TrPr nil")
	}
	// Marshal to exercise TrPr branch in CT_Row.MarshalXML
	data, err := xmlutil.Marshal(&tbl, xmlutil.OOXML)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	var tbl2 CT_Tbl
	if err := xmlutil.Unmarshal(data, &tbl2); err != nil {
		t.Fatalf("re-Unmarshal: %v", err)
	}
	if tbl2.Rows[0].TrPr == nil {
		t.Error("TrPr lost after round-trip")
	}
}

func TestCT_TblGrid_UnknownChild(t *testing.T) {
	// TblGrid with unknown child should skip it
	input := `<w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:tblGrid>
			<w:gridCol w:w="3000"/>
			<w:unknownGridChild/>
			<w:gridCol w:w="3000"/>
		</w:tblGrid>
		<w:tr><w:tc><w:p/></w:tc></w:tr>
	</w:tbl>`
	var tbl CT_Tbl
	if err := xmlutil.Unmarshal([]byte(input), &tbl); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if tbl.TblGrid == nil {
		t.Fatal("TblGrid nil")
	}
	if len(tbl.TblGrid.Cols) != 2 {
		t.Errorf("cols = %d, want 2", len(tbl.TblGrid.Cols))
	}
}

func TestCT_Tc_NestedTable_Marshal(t *testing.T) {
	// Cell containing nested table — exercise table branch in CT_Tc.MarshalXML
	tc := CT_Tc{XMLName: xml.Name{Space: Ns, Local: "tc"}}
	nested := &CT_Tbl{XMLName: xml.Name{Space: Ns, Local: "tbl"}}
	nestedRow := &CT_Row{XMLName: xml.Name{Space: Ns, Local: "tr"}}
	nestedCell := &CT_Tc{XMLName: xml.Name{Space: Ns, Local: "tc"}}
	np := &CT_P{XMLName: xml.Name{Space: Ns, Local: "p"}}
	np.AddRun("nested text")
	nestedCell.Content = append(nestedCell.Content, BlockLevelContent{Paragraph: np})
	nestedRow.Cells = append(nestedRow.Cells, nestedCell)
	nested.Rows = append(nested.Rows, nestedRow)
	tc.Content = append(tc.Content, BlockLevelContent{Table: nested})

	row := CT_Row{XMLName: xml.Name{Space: Ns, Local: "tr"}, Cells: []*CT_Tc{&tc}}
	tbl := CT_Tbl{XMLName: xml.Name{Space: Ns, Local: "tbl"}, Rows: []*CT_Row{&row}}
	data, err := xmlutil.Marshal(&tbl, xmlutil.OOXML)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	var tbl2 CT_Tbl
	if err := xmlutil.Unmarshal(data, &tbl2); err != nil {
		t.Fatalf("re-Unmarshal: %v", err)
	}
	c2 := tbl2.Rows[0].Cells[0]
	if len(c2.Content) != 1 || c2.Content[0].Table == nil {
		t.Error("nested table lost")
	}
}
