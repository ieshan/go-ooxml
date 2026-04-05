package docx

import (
	"bytes"
	"testing"
)

func TestSection_AddHeader_Text(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	for section := range doc.Sections() {
		hdr := section.AddHeader(HdrFtrDefault)
		hdr.AddParagraph().AddRun("Company Name")
		if hdr.Text() != "Company Name" {
			t.Errorf("text = %q", hdr.Text())
		}
		if hdr.Type() != HdrFtrDefault {
			t.Errorf("type = %d", hdr.Type())
		}
	}
}

func TestSection_AddFooter_Text(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	for section := range doc.Sections() {
		ftr := section.AddFooter(HdrFtrDefault)
		ftr.AddParagraph().AddRun("Page 1")
		if ftr.Text() != "Page 1" {
			t.Errorf("text = %q", ftr.Text())
		}
	}
}

func TestSection_AddHeader_FirstPage(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	for section := range doc.Sections() {
		hdr := section.AddHeader(HdrFtrFirst)
		hdr.AddParagraph().AddRun("Title Page Header")
		if hdr.Type() != HdrFtrFirst {
			t.Errorf("type = %d, want HdrFtrFirst", hdr.Type())
		}
	}
}

func TestSection_Headers_Iterator(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	for section := range doc.Sections() {
		section.AddHeader(HdrFtrDefault)
		section.AddHeader(HdrFtrFirst)
		count := 0
		for range section.Headers() {
			count++
		}
		if count != 2 {
			t.Errorf("header count = %d, want 2", count)
		}
	}
}

func TestSection_Footers_Iterator(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	for section := range doc.Sections() {
		section.AddFooter(HdrFtrDefault)
		count := 0
		for range section.Footers() {
			count++
		}
		if count != 1 {
			t.Errorf("footer count = %d, want 1", count)
		}
	}
}

func TestHeaderFooter_Paragraphs(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	for section := range doc.Sections() {
		hdr := section.AddHeader(HdrFtrDefault)
		hdr.AddParagraph().AddRun("Para 1")
		hdr.AddParagraph().AddRun("Para 2")
		count := 0
		for range hdr.Paragraphs() {
			count++
		}
		if count != 2 {
			t.Errorf("paragraphs = %d, want 2", count)
		}
	}
}

func TestHeaderFooter_Markdown(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	for section := range doc.Sections() {
		hdr := section.AddHeader(HdrFtrDefault)
		p := hdr.AddParagraph()
		b := true
		r := p.AddRun("Bold Header")
		r.SetBold(&b)
		md := hdr.Markdown()
		if md != "**Bold Header**" {
			t.Errorf("markdown = %q", md)
		}
	}
}

func TestHeaderFooter_EmptyText(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	for section := range doc.Sections() {
		hdr := section.AddHeader(HdrFtrDefault)
		if hdr.Text() != "" {
			t.Errorf("empty header text = %q", hdr.Text())
		}
		if hdr.Markdown() != "" {
			t.Errorf("empty header markdown = %q", hdr.Markdown())
		}
	}
}

func TestHeaderFooter_RoundTrip(t *testing.T) {
	doc, _ := New(nil)
	for section := range doc.Sections() {
		hdr := section.AddHeader(HdrFtrDefault)
		hdr.AddParagraph().AddRun("Company Inc.")
		ftr := section.AddFooter(HdrFtrDefault)
		ftr.AddParagraph().AddRun("Confidential")
	}

	var buf bytes.Buffer
	doc.WriteTo(&buf)
	doc.Close()

	doc2, err := OpenReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()), nil)
	if err != nil {
		t.Fatalf("OpenReader: %v", err)
	}
	defer doc2.Close()

	for section := range doc2.Sections() {
		hdrFound := false
		for hdr := range section.Headers() {
			if hdr.Text() == "Company Inc." {
				hdrFound = true
			}
		}
		if !hdrFound {
			t.Error("header text not found after round-trip")
		}

		ftrFound := false
		for ftr := range section.Footers() {
			if ftr.Text() == "Confidential" {
				ftrFound = true
			}
		}
		if !ftrFound {
			t.Error("footer text not found after round-trip")
		}
	}
}

func ExampleSection_AddHeader() {
	doc, _ := New(nil)
	defer doc.Close()
	for section := range doc.Sections() {
		hdr := section.AddHeader(HdrFtrDefault)
		hdr.AddParagraph().AddRun("Company Name")
	}
}

func ExampleSection_AddFooter() {
	doc, _ := New(nil)
	defer doc.Close()
	for section := range doc.Sections() {
		ftr := section.AddFooter(HdrFtrDefault)
		ftr.AddParagraph().AddRun("Page 1 of 10")
	}
}

func ExampleHeaderFooter_Text() {
	doc, _ := New(nil)
	defer doc.Close()
	for section := range doc.Sections() {
		hdr := section.AddHeader(HdrFtrDefault)
		hdr.AddParagraph().AddRun("Header Text")
		_ = hdr.Text() // "Header Text"
	}
}
