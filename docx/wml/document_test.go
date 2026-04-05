package wml

import (
	"encoding/xml"
	"testing"

	"github.com/ieshan/go-ooxml/xmlutil"
)

func TestCT_Document_Unmarshal(t *testing.T) {
	input := `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:body>
			<w:p><w:r><w:t>Hello</w:t></w:r></w:p>
			<w:p><w:r><w:t>World</w:t></w:r></w:p>
		</w:body>
	</w:document>`
	var doc CT_Document
	if err := xmlutil.Unmarshal([]byte(input), &doc); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if doc.Body == nil {
		t.Fatal("Body nil")
	}
	paraCount := 0
	for _, c := range doc.Body.Content {
		if c.Paragraph != nil {
			paraCount++
		}
	}
	if paraCount != 2 {
		t.Errorf("paragraph count = %d, want 2", paraCount)
	}
}

func TestCT_Document_Marshal_RoundTrip(t *testing.T) {
	doc := CT_Document{
		XMLName: xml.Name{Space: Ns, Local: "document"},
		Body:    &CT_Body{XMLName: xml.Name{Space: Ns, Local: "body"}},
	}
	p := &CT_P{XMLName: xml.Name{Space: Ns, Local: "p"}}
	p.AddRun("Test")
	doc.Body.Content = append(doc.Body.Content, BlockLevelContent{Paragraph: p})

	data, err := xmlutil.Marshal(&doc, xmlutil.OOXML)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	var doc2 CT_Document
	if err := xmlutil.Unmarshal(data, &doc2); err != nil {
		t.Fatalf("re-Unmarshal: %v", err)
	}
	if doc2.Body == nil {
		t.Fatal("body nil")
	}
	if len(doc2.Body.Content) != 1 {
		t.Errorf("content = %d", len(doc2.Body.Content))
	}
}

func TestCT_Body_PreservesUnknown(t *testing.T) {
	input := `<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:p><w:r><w:t>text</w:t></w:r></w:p>
		<w:customXml><data/></w:customXml>
	</w:body>`
	var body CT_Body
	if err := xmlutil.Unmarshal([]byte(input), &body); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if len(body.Content) != 2 {
		t.Errorf("content = %d, want 2", len(body.Content))
	}
}

func TestCT_Body_PreservesSectPr(t *testing.T) {
	input := `<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:p><w:r><w:t>text</w:t></w:r></w:p>
		<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
	</w:body>`
	var body CT_Body
	if err := xmlutil.Unmarshal([]byte(input), &body); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if len(body.Content) != 1 {
		t.Errorf("content = %d, want 1 (sectPr should not be in Content)", len(body.Content))
	}
	if body.SectPr == nil {
		t.Error("sectPr not preserved")
	}
}

func TestCT_Body_WithTable(t *testing.T) {
	input := `<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:p><w:r><w:t>before</w:t></w:r></w:p>
		<w:tbl><w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
		<w:p><w:r><w:t>after</w:t></w:r></w:p>
	</w:body>`
	var body CT_Body
	if err := xmlutil.Unmarshal([]byte(input), &body); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if len(body.Content) != 3 {
		t.Errorf("content = %d, want 3", len(body.Content))
	}
	if body.Content[0].Paragraph == nil {
		t.Error("first should be paragraph")
	}
	if body.Content[1].Table == nil {
		t.Error("second should be table")
	}
	if body.Content[2].Paragraph == nil {
		t.Error("third should be paragraph")
	}
}

func TestCT_Document_Paragraphs(t *testing.T) {
	input := `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:body>
			<w:p><w:r><w:t>one</w:t></w:r></w:p>
			<w:p><w:r><w:t>two</w:t></w:r></w:p>
		</w:body>
	</w:document>`
	var doc CT_Document
	if err := xmlutil.Unmarshal([]byte(input), &doc); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	paras := doc.Body.Paragraphs()
	if len(paras) != 2 {
		t.Errorf("Paragraphs() = %d, want 2", len(paras))
	}
	if paras[0].Text() != "one" {
		t.Errorf("first = %q", paras[0].Text())
	}
}

func TestCT_Body_MixedContent_Marshal(t *testing.T) {
	// Build body with paragraph + table + unknown
	body := CT_Body{XMLName: xml.Name{Space: Ns, Local: "body"}}
	p := &CT_P{XMLName: xml.Name{Space: Ns, Local: "p"}}
	p.AddRun("hello")
	body.Content = append(body.Content, BlockLevelContent{Paragraph: p})

	tbl := &CT_Tbl{XMLName: xml.Name{Space: Ns, Local: "tbl"}}
	row := &CT_Row{XMLName: xml.Name{Space: Ns, Local: "tr"}}
	cell := &CT_Tc{XMLName: xml.Name{Space: Ns, Local: "tc"}}
	cp := &CT_P{XMLName: xml.Name{Space: Ns, Local: "p"}}
	cp.AddRun("cell text")
	cell.Content = append(cell.Content, BlockLevelContent{Paragraph: cp})
	row.Cells = append(row.Cells, cell)
	tbl.Rows = append(tbl.Rows, row)
	body.Content = append(body.Content, BlockLevelContent{Table: tbl})

	data, err := xmlutil.Marshal(&body, xmlutil.OOXML)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	var body2 CT_Body
	if err := xmlutil.Unmarshal(data, &body2); err != nil {
		t.Fatalf("re-Unmarshal: %v", err)
	}
	if len(body2.Content) != 2 {
		t.Fatalf("content = %d, want 2", len(body2.Content))
	}
	if body2.Content[0].Paragraph == nil {
		t.Error("first should be paragraph")
	}
	if body2.Content[1].Table == nil {
		t.Error("second should be table")
	}
}

func TestCT_Body_Marshal_PreservesSectPr(t *testing.T) {
	input := `<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:p><w:r><w:t>content</w:t></w:r></w:p>
		<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
	</w:body>`
	var body CT_Body
	if err := xmlutil.Unmarshal([]byte(input), &body); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if body.SectPr == nil {
		t.Fatal("SectPr nil")
	}
	// Marshal and re-unmarshal preserves sectPr
	data, err := xmlutil.Marshal(&body, xmlutil.OOXML)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	var body2 CT_Body
	if err := xmlutil.Unmarshal(data, &body2); err != nil {
		t.Fatalf("re-Unmarshal: %v", err)
	}
	if body2.SectPr == nil {
		t.Error("SectPr lost after round-trip")
	}
	if len(body2.Content) != 1 {
		t.Errorf("content = %d, want 1", len(body2.Content))
	}
}

func TestCT_Body_Raw_MarshalXML(t *testing.T) {
	input := `<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:customXml><data/></w:customXml>
		<w:p><w:r><w:t>text</w:t></w:r></w:p>
	</w:body>`
	var body CT_Body
	if err := xmlutil.Unmarshal([]byte(input), &body); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	data, err := xmlutil.Marshal(&body, xmlutil.OOXML)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	var body2 CT_Body
	if err := xmlutil.Unmarshal(data, &body2); err != nil {
		t.Fatalf("re-Unmarshal: %v", err)
	}
	if len(body2.Content) != 2 {
		t.Errorf("content = %d, want 2", len(body2.Content))
	}
	if body2.Content[0].Raw == nil {
		t.Error("first should be raw")
	}
}

func TestCT_Document_FullRoundTrip(t *testing.T) {
	// Comprehensive full document round-trip
	input := `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
	                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
		<w:body>
			<w:p>
				<w:pPr>
					<w:pStyle w:val="Heading1"/>
					<w:jc w:val="center"/>
					<w:spacing w:after="200"/>
				</w:pPr>
				<w:r><w:rPr><w:b/><w:sz w:val="28"/></w:rPr><w:t>Title</w:t></w:r>
			</w:p>
			<w:tbl>
				<w:tblPr>
					<w:tblStyle w:val="TableGrid"/>
					<w:tblW w:w="5000" w:type="pct"/>
				</w:tblPr>
				<w:tblGrid><w:gridCol w:w="2500"/><w:gridCol w:w="2500"/></w:tblGrid>
				<w:tr>
					<w:tc>
						<w:tcPr><w:tcW w:w="2500" w:type="dxa"/></w:tcPr>
						<w:p><w:r><w:t>Cell A1</w:t></w:r></w:p>
					</w:tc>
					<w:tc>
						<w:p><w:r><w:t>Cell A2</w:t></w:r></w:p>
					</w:tc>
				</w:tr>
			</w:tbl>
			<w:p>
				<w:r>
					<w:rPr>
						<w:color w:val="333333"/>
						<w:sz w:val="24"/>
					</w:rPr>
					<w:t>Body text </w:t>
				</w:r>
				<w:hyperlink r:id="rId1">
					<w:r><w:t>link</w:t></w:r>
				</w:hyperlink>
			</w:p>
			<w:sectPr>
				<w:pgSz w:w="12240" w:h="15840"/>
				<w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800"/>
			</w:sectPr>
		</w:body>
	</w:document>`
	var doc CT_Document
	if err := xmlutil.Unmarshal([]byte(input), &doc); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if doc.Body == nil {
		t.Fatal("Body nil")
	}
	// Marshal
	data, err := xmlutil.Marshal(&doc, xmlutil.OOXML)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	var doc2 CT_Document
	if err := xmlutil.Unmarshal(data, &doc2); err != nil {
		t.Fatalf("re-Unmarshal: %v", err)
	}
	paras := doc2.Body.Paragraphs()
	if len(paras) < 2 {
		t.Errorf("paragraphs = %d, want >= 2", len(paras))
	}
	if doc2.Body.SectPr == nil {
		t.Error("SectPr lost")
	}
	// Verify content preserved
	found := false
	for _, c := range doc2.Body.Content {
		if c.Table != nil {
			found = true
		}
	}
	if !found {
		t.Error("table not preserved")
	}
}

func TestCT_Document_NoBody(t *testing.T) {
	input := `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:unknownChild/>
	</w:document>`
	var doc CT_Document
	if err := xmlutil.Unmarshal([]byte(input), &doc); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if doc.Body != nil {
		t.Error("Body should be nil when no body element")
	}
	// Also test marshal with nil body
	doc2 := CT_Document{XMLName: xml.Name{Space: Ns, Local: "document"}}
	data, err := xmlutil.Marshal(&doc2, xmlutil.OOXML)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	var doc3 CT_Document
	if err := xmlutil.Unmarshal(data, &doc3); err != nil {
		t.Fatalf("re-Unmarshal: %v", err)
	}
	if doc3.Body != nil {
		t.Error("Body should remain nil")
	}
}
