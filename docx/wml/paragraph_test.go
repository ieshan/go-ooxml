package wml

import (
	"encoding/xml"
	"testing"

	"github.com/ieshan/go-ooxml/xmlutil"
)

func TestCT_P_Unmarshal_SimpleText(t *testing.T) {
	input := `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:r><w:t>Hello World</w:t></w:r>
	</w:p>`
	var p CT_P
	if err := xmlutil.Unmarshal([]byte(input), &p); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if p.Text() != "Hello World" {
		t.Errorf("Text() = %q", p.Text())
	}
}

func TestCT_P_Unmarshal_MultipleRuns(t *testing.T) {
	input := `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:r><w:t>Hello </w:t></w:r>
		<w:r><w:t>World</w:t></w:r>
	</w:p>`
	var p CT_P
	if err := xmlutil.Unmarshal([]byte(input), &p); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if p.Text() != "Hello World" {
		t.Errorf("Text() = %q", p.Text())
	}
	runCount := 0
	for _, c := range p.Content {
		if c.Run != nil {
			runCount++
		}
	}
	if runCount != 2 {
		t.Errorf("run count = %d, want 2", runCount)
	}
}

func TestCT_P_Unmarshal_WithStyle(t *testing.T) {
	input := `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
		<w:r><w:t>Title</w:t></w:r>
	</w:p>`
	var p CT_P
	if err := xmlutil.Unmarshal([]byte(input), &p); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if p.PPr == nil || p.PPr.Style == nil || *p.PPr.Style != "Heading1" {
		t.Error("style not parsed")
	}
}

func TestCT_P_Unmarshal_WithAlignment(t *testing.T) {
	input := `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:pPr><w:jc w:val="center"/></w:pPr>
		<w:r><w:t>centered</w:t></w:r>
	</w:p>`
	var p CT_P
	if err := xmlutil.Unmarshal([]byte(input), &p); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if p.PPr == nil || p.PPr.Alignment == nil || *p.PPr.Alignment != "center" {
		t.Error("alignment not parsed")
	}
}

func TestCT_P_Unmarshal_WithNumPr(t *testing.T) {
	input := `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:pPr>
			<w:numPr>
				<w:ilvl w:val="0"/>
				<w:numId w:val="1"/>
			</w:numPr>
		</w:pPr>
		<w:r><w:t>list item</w:t></w:r>
	</w:p>`
	var p CT_P
	if err := xmlutil.Unmarshal([]byte(input), &p); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if p.PPr == nil || p.PPr.NumPr == nil {
		t.Fatal("numPr not parsed")
	}
	if p.PPr.NumPr.ILvl == nil || *p.PPr.NumPr.ILvl != 0 {
		t.Error("ilvl not parsed")
	}
	if p.PPr.NumPr.NumID == nil || *p.PPr.NumPr.NumID != 1 {
		t.Error("numId not parsed")
	}
}

func TestCT_P_Unmarshal_WithIns(t *testing.T) {
	input := `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:ins w:id="1" w:author="Joe" w:date="2024-01-01T00:00:00Z">
			<w:r><w:t>inserted</w:t></w:r>
		</w:ins>
	</w:p>`
	var p CT_P
	if err := xmlutil.Unmarshal([]byte(input), &p); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if len(p.Content) != 1 || p.Content[0].Ins == nil {
		t.Fatal("ins not parsed")
	}
	if p.Content[0].Ins.Author != "Joe" {
		t.Error("author")
	}
	if len(p.Content[0].Ins.Runs) != 1 {
		t.Error("runs in ins")
	}
}

func TestCT_P_Unmarshal_WithDel(t *testing.T) {
	input := `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:del w:id="2" w:author="Jane">
			<w:r><w:t>deleted</w:t></w:r>
		</w:del>
	</w:p>`
	var p CT_P
	if err := xmlutil.Unmarshal([]byte(input), &p); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if len(p.Content) != 1 || p.Content[0].Del == nil {
		t.Fatal("del not parsed")
	}
	if p.Content[0].Del.Author != "Jane" {
		t.Error("author")
	}
}

func TestCT_P_Unmarshal_WithHyperlink(t *testing.T) {
	input := `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
	             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
		<w:hyperlink r:id="rId1">
			<w:r><w:t>click here</w:t></w:r>
		</w:hyperlink>
	</w:p>`
	var p CT_P
	if err := xmlutil.Unmarshal([]byte(input), &p); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if len(p.Content) != 1 || p.Content[0].Hyperlink == nil {
		t.Fatal("hyperlink not parsed")
	}
	if p.Content[0].Hyperlink.ID != "rId1" {
		t.Errorf("hyperlink ID = %q, want rId1", p.Content[0].Hyperlink.ID)
	}
	if len(p.Content[0].Hyperlink.Runs) != 1 {
		t.Error("runs in hyperlink")
	}
}

func TestCT_P_Unmarshal_WithCommentRange(t *testing.T) {
	input := `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:commentRangeStart w:id="0"/>
		<w:r><w:t>commented</w:t></w:r>
		<w:commentRangeEnd w:id="0"/>
	</w:p>`
	var p CT_P
	if err := xmlutil.Unmarshal([]byte(input), &p); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if len(p.Content) != 3 {
		t.Fatalf("content = %d, want 3", len(p.Content))
	}
	if p.Content[0].CommentRangeStart == nil {
		t.Error("commentRangeStart not parsed")
	}
	if p.Content[2].CommentRangeEnd == nil {
		t.Error("commentRangeEnd not parsed")
	}
}

func TestCT_P_AddRun(t *testing.T) {
	p := CT_P{XMLName: xml.Name{Space: Ns, Local: "p"}}
	r := p.AddRun("test")
	if r == nil {
		t.Fatal("nil")
	}
	if p.Text() != "test" {
		t.Errorf("Text() = %q", p.Text())
	}
}

func TestCT_P_PreservesUnknown(t *testing.T) {
	input := `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:r><w:t>text</w:t></w:r>
		<w:bookmarkStart w:id="0" w:name="test"/>
	</w:p>`
	var p CT_P
	if err := xmlutil.Unmarshal([]byte(input), &p); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if len(p.Content) != 2 {
		t.Errorf("content = %d, want 2", len(p.Content))
	}
	if p.Content[1].Raw == nil {
		t.Error("unknown should be RawXML")
	}
}

func TestCT_P_Marshal_RoundTrip(t *testing.T) {
	p := CT_P{XMLName: xml.Name{Space: Ns, Local: "p"}}
	style := "Heading1"
	p.PPr = &CT_PPr{Style: &style}
	p.AddRun("Hello")
	data, err := xmlutil.Marshal(&p, xmlutil.OOXML)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	var p2 CT_P
	if err := xmlutil.Unmarshal(data, &p2); err != nil {
		t.Fatalf("re-Unmarshal: %v", err)
	}
	if p2.Text() != "Hello" {
		t.Error("text lost")
	}
	if p2.PPr == nil || p2.PPr.Style == nil || *p2.PPr.Style != "Heading1" {
		t.Error("style lost")
	}
}

func TestCT_P_Text_IncludesHyperlinkRuns(t *testing.T) {
	input := `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
	             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
		<w:hyperlink r:id="rId1">
			<w:r><w:t>link </w:t></w:r>
			<w:r><w:t>text</w:t></w:r>
		</w:hyperlink>
	</w:p>`
	var p CT_P
	if err := xmlutil.Unmarshal([]byte(input), &p); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if p.Text() != "link text" {
		t.Errorf("Text() = %q, want %q", p.Text(), "link text")
	}
}

func TestCT_P_Text_IncludesInsDelRuns(t *testing.T) {
	input := `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:r><w:t>before </w:t></w:r>
		<w:ins w:id="1" w:author="Joe">
			<w:r><w:t>inserted</w:t></w:r>
		</w:ins>
		<w:del w:id="2" w:author="Joe">
			<w:r><w:t> deleted</w:t></w:r>
		</w:del>
	</w:p>`
	var p CT_P
	if err := xmlutil.Unmarshal([]byte(input), &p); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	text := p.Text()
	// Deleted text inside <w:del> is excluded from visible text.
	if text != "before inserted" {
		t.Errorf("Text() = %q, want %q", text, "before inserted")
	}
}

func TestCT_PPr_ExtraPreserved(t *testing.T) {
	input := `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:pPr>
			<w:pStyle w:val="Normal"/>
			<w:spacing w:after="200"/>
		</w:pPr>
		<w:r><w:t>text</w:t></w:r>
	</w:p>`
	var p CT_P
	if err := xmlutil.Unmarshal([]byte(input), &p); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if p.PPr == nil {
		t.Fatal("pPr nil")
	}
	if len(p.PPr.Extra) != 1 {
		t.Errorf("Extra = %d, want 1 (spacing)", len(p.PPr.Extra))
	}
}

func TestCT_P_Marshal_WithHyperlink(t *testing.T) {
	p := CT_P{XMLName: xml.Name{Space: Ns, Local: "p"}}
	h := &CT_Hyperlink{XMLName: xml.Name{Space: Ns, Local: "hyperlink"}, ID: "rId5"}
	r1 := &CT_R{XMLName: xml.Name{Space: Ns, Local: "r"}}
	r1.AddText("link text 1")
	r2 := &CT_R{XMLName: xml.Name{Space: Ns, Local: "r"}}
	r2.AddText(" link text 2")
	h.Runs = append(h.Runs, r1, r2)
	p.Content = append(p.Content, InlineContent{Hyperlink: h})

	data, err := xmlutil.Marshal(&p, xmlutil.OOXML)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	var p2 CT_P
	if err := xmlutil.Unmarshal(data, &p2); err != nil {
		t.Fatalf("re-Unmarshal: %v", err)
	}
	if len(p2.Content) != 1 || p2.Content[0].Hyperlink == nil {
		t.Fatal("hyperlink not preserved")
	}
	if len(p2.Content[0].Hyperlink.Runs) != 2 {
		t.Errorf("runs = %d, want 2", len(p2.Content[0].Hyperlink.Runs))
	}
}

func TestCT_P_Marshal_WithCommentRange(t *testing.T) {
	p := CT_P{XMLName: xml.Name{Space: Ns, Local: "p"}}
	p.Content = append(p.Content, InlineContent{CommentRangeStart: &CT_MarkupRange{ID: 3}})
	r := &CT_R{XMLName: xml.Name{Space: Ns, Local: "r"}}
	r.AddText("commented text")
	p.Content = append(p.Content, InlineContent{Run: r})
	p.Content = append(p.Content, InlineContent{CommentRangeEnd: &CT_MarkupRange{ID: 3}})

	data, err := xmlutil.Marshal(&p, xmlutil.OOXML)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	var p2 CT_P
	if err := xmlutil.Unmarshal(data, &p2); err != nil {
		t.Fatalf("re-Unmarshal: %v", err)
	}
	if len(p2.Content) != 3 {
		t.Fatalf("content = %d, want 3", len(p2.Content))
	}
	if p2.Content[0].CommentRangeStart == nil || p2.Content[0].CommentRangeStart.ID != 3 {
		t.Error("CommentRangeStart")
	}
	if p2.Content[2].CommentRangeEnd == nil || p2.Content[2].CommentRangeEnd.ID != 3 {
		t.Error("CommentRangeEnd")
	}
}

func TestCT_P_Marshal_WithDel(t *testing.T) {
	p := CT_P{XMLName: xml.Name{Space: Ns, Local: "p"}}
	del := &CT_RunTrackChange{
		XMLName: xml.Name{Space: Ns, Local: "del"},
		ID:      5, Author: "Editor", Date: "2024-01-01T00:00:00Z",
	}
	r := &CT_R{XMLName: xml.Name{Space: Ns, Local: "r"}}
	r.Content = append(r.Content, RunContent{DelText: &CT_Text{Value: "old text"}})
	del.Runs = append(del.Runs, r)
	p.Content = append(p.Content, InlineContent{Del: del})

	data, err := xmlutil.Marshal(&p, xmlutil.OOXML)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	var p2 CT_P
	if err := xmlutil.Unmarshal(data, &p2); err != nil {
		t.Fatalf("re-Unmarshal: %v", err)
	}
	if len(p2.Content) != 1 || p2.Content[0].Del == nil {
		t.Fatal("del not preserved")
	}
	if p2.Content[0].Del.Author != "Editor" {
		t.Errorf("Author = %q", p2.Content[0].Del.Author)
	}
}

func TestCT_NumPr_UnknownChild(t *testing.T) {
	// NumPr with unknown child should skip it gracefully
	input := `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:pPr>
			<w:numPr>
				<w:ilvl w:val="0"/>
				<w:numId w:val="1"/>
				<w:unknownNumChild/>
			</w:numPr>
		</w:pPr>
		<w:r><w:t>item</w:t></w:r>
	</w:p>`
	var p CT_P
	if err := xmlutil.Unmarshal([]byte(input), &p); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if p.PPr == nil || p.PPr.NumPr == nil {
		t.Fatal("NumPr nil")
	}
	if p.PPr.NumPr.ILvl == nil || *p.PPr.NumPr.ILvl != 0 {
		t.Error("ILvl")
	}
}

func TestCT_PPr_NumPr_RoundTrip(t *testing.T) {
	ilvl := 1
	numID := 3
	p := CT_P{XMLName: xml.Name{Space: Ns, Local: "p"}}
	p.PPr = &CT_PPr{
		NumPr: &CT_NumPr{ILvl: &ilvl, NumID: &numID},
	}
	p.AddRun("list item")

	data, err := xmlutil.Marshal(&p, xmlutil.OOXML)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	var p2 CT_P
	if err := xmlutil.Unmarshal(data, &p2); err != nil {
		t.Fatalf("re-Unmarshal: %v", err)
	}
	if p2.PPr == nil || p2.PPr.NumPr == nil {
		t.Fatal("NumPr lost")
	}
	if p2.PPr.NumPr.ILvl == nil || *p2.PPr.NumPr.ILvl != 1 {
		t.Error("ILvl")
	}
	if p2.PPr.NumPr.NumID == nil || *p2.PPr.NumPr.NumID != 3 {
		t.Error("NumID")
	}
}

func TestCT_P_Marshal_RoundTrip_PreservesUnknownInline(t *testing.T) {
	input := `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:r><w:t>text</w:t></w:r>
		<w:bookmarkStart w:id="0" w:name="anchor"/>
	</w:p>`
	var p CT_P
	if err := xmlutil.Unmarshal([]byte(input), &p); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	data, err := xmlutil.Marshal(&p, xmlutil.OOXML)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	var p2 CT_P
	if err := xmlutil.Unmarshal(data, &p2); err != nil {
		t.Fatalf("re-Unmarshal: %v", err)
	}
	if len(p2.Content) != 2 {
		t.Fatalf("content = %d, want 2", len(p2.Content))
	}
	if p2.Content[1].Raw == nil {
		t.Error("unknown inline element should be preserved as Raw")
	}
}

func TestCT_P_Marshal_WithIns(t *testing.T) {
	p := CT_P{XMLName: xml.Name{Space: Ns, Local: "p"}}
	ins := &CT_RunTrackChange{
		XMLName: xml.Name{Space: Ns, Local: "ins"},
		ID:      1, Author: "Writer", Date: "2024-01-01T00:00:00Z",
	}
	r := &CT_R{XMLName: xml.Name{Space: Ns, Local: "r"}}
	r.AddText("new text")
	ins.Runs = append(ins.Runs, r)
	p.Content = append(p.Content, InlineContent{Ins: ins})

	data, err := xmlutil.Marshal(&p, xmlutil.OOXML)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	var p2 CT_P
	if err := xmlutil.Unmarshal(data, &p2); err != nil {
		t.Fatalf("re-Unmarshal: %v", err)
	}
	if len(p2.Content) != 1 || p2.Content[0].Ins == nil {
		t.Fatal("ins not preserved")
	}
	if p2.Content[0].Ins.Author != "Writer" {
		t.Errorf("Author = %q", p2.Content[0].Ins.Author)
	}
}

func TestCT_Hyperlink_AnchorRoundTrip(t *testing.T) {
	p := CT_P{XMLName: xml.Name{Space: Ns, Local: "p"}}
	h := &CT_Hyperlink{
		XMLName: xml.Name{Space: Ns, Local: "hyperlink"},
		Anchor:  "section1",
	}
	r := &CT_R{XMLName: xml.Name{Space: Ns, Local: "r"}}
	r.AddText("go to section")
	h.Runs = append(h.Runs, r)
	p.Content = append(p.Content, InlineContent{Hyperlink: h})

	data, err := xmlutil.Marshal(&p, xmlutil.OOXML)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	var p2 CT_P
	if err := xmlutil.Unmarshal(data, &p2); err != nil {
		t.Fatalf("re-Unmarshal: %v", err)
	}
	if len(p2.Content) != 1 || p2.Content[0].Hyperlink == nil {
		t.Fatal("hyperlink not preserved")
	}
	if p2.Content[0].Hyperlink.Anchor != "section1" {
		t.Errorf("Anchor = %q", p2.Content[0].Hyperlink.Anchor)
	}
}

func TestCT_PPr_Marshal_WithAlignment(t *testing.T) {
	align := "right"
	p := CT_P{XMLName: xml.Name{Space: Ns, Local: "p"}}
	p.PPr = &CT_PPr{Alignment: &align}
	p.AddRun("right-aligned")
	data, err := xmlutil.Marshal(&p, xmlutil.OOXML)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	var p2 CT_P
	if err := xmlutil.Unmarshal(data, &p2); err != nil {
		t.Fatalf("re-Unmarshal: %v", err)
	}
	if p2.PPr == nil || p2.PPr.Alignment == nil || *p2.PPr.Alignment != "right" {
		t.Error("Alignment lost")
	}
}

func TestCT_PPr_Marshal_WithExtra(t *testing.T) {
	input := `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:pPr>
			<w:pStyle w:val="Normal"/>
			<w:jc w:val="center"/>
			<w:spacing w:after="200"/>
		</w:pPr>
		<w:r><w:t>text</w:t></w:r>
	</w:p>`
	var p CT_P
	if err := xmlutil.Unmarshal([]byte(input), &p); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	// Marshal to exercise the Extra path in CT_PPr.MarshalXML
	data, err := xmlutil.Marshal(&p, xmlutil.OOXML)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	var p2 CT_P
	if err := xmlutil.Unmarshal(data, &p2); err != nil {
		t.Fatalf("re-Unmarshal: %v", err)
	}
	if p2.PPr == nil {
		t.Fatal("PPr nil")
	}
	if p2.PPr.Style == nil || *p2.PPr.Style != "Normal" {
		t.Error("Style")
	}
	if p2.PPr.Alignment == nil || *p2.PPr.Alignment != "center" {
		t.Error("Alignment")
	}
	if len(p2.PPr.Extra) != 1 {
		t.Errorf("Extra = %d, want 1", len(p2.PPr.Extra))
	}
}

func TestCT_Hyperlink_Raw_RoundTrip(t *testing.T) {
	// Test Hyperlink with Raw content (non-run children)
	input := `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
	             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
		<w:hyperlink r:id="rId1">
			<w:r><w:t>text</w:t></w:r>
			<w:proofErr w:type="spellEnd"/>
		</w:hyperlink>
	</w:p>`
	var p CT_P
	if err := xmlutil.Unmarshal([]byte(input), &p); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	hl := p.Content[0].Hyperlink
	if hl == nil {
		t.Fatal("Hyperlink nil")
	}
	if len(hl.Raw) != 1 {
		t.Errorf("Raw = %d, want 1", len(hl.Raw))
	}
	// Marshal to exercise Raw path in CT_Hyperlink.MarshalXML
	data, err := xmlutil.Marshal(&p, xmlutil.OOXML)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	var p2 CT_P
	if err := xmlutil.Unmarshal(data, &p2); err != nil {
		t.Fatalf("re-Unmarshal: %v", err)
	}
	if len(p2.Content) != 1 || p2.Content[0].Hyperlink == nil {
		t.Fatal("Hyperlink lost")
	}
	if len(p2.Content[0].Hyperlink.Raw) != 1 {
		t.Errorf("Hyperlink.Raw = %d, want 1", len(p2.Content[0].Hyperlink.Raw))
	}
}

func TestCT_RunTrackChange_Raw_RoundTrip(t *testing.T) {
	// Test del with non-run child preserved as Raw
	input := `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
		<w:del w:id="3" w:author="X">
			<w:r><w:delText>deleted</w:delText></w:r>
			<w:bookmarkEnd w:id="0"/>
		</w:del>
	</w:p>`
	var p CT_P
	if err := xmlutil.Unmarshal([]byte(input), &p); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	del := p.Content[0].Del
	if del == nil {
		t.Fatal("del nil")
	}
	if len(del.Raw) != 1 {
		t.Errorf("Raw = %d, want 1", len(del.Raw))
	}
	// Marshal to exercise Raw path in CT_RunTrackChange.MarshalXML
	data, err := xmlutil.Marshal(&p, xmlutil.OOXML)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	var p2 CT_P
	if err := xmlutil.Unmarshal(data, &p2); err != nil {
		t.Fatalf("re-Unmarshal: %v", err)
	}
	if p2.Content[0].Del == nil {
		t.Fatal("del lost")
	}
	if len(p2.Content[0].Del.Raw) != 1 {
		t.Errorf("del.Raw = %d, want 1", len(p2.Content[0].Del.Raw))
	}
}
