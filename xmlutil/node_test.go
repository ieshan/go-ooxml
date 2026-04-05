package xmlutil

import (
	"encoding/xml"
	"strings"
	"testing"
)

func TestRawXML_UnmarshalMarshal_RoundTrip(t *testing.T) {
	// Test that unknown XML elements can be captured and re-emitted
	type wrapper struct {
		XMLName xml.Name `xml:"root"`
		Extra   []RawXML `xml:",any"`
	}
	input := `<root><unknown attr="val">text</unknown><other>data</other></root>`

	var w wrapper
	err := xml.Unmarshal([]byte(input), &w)
	if err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if len(w.Extra) != 2 {
		t.Fatalf("Extra count = %d, want 2", len(w.Extra))
	}

	out, err := xml.Marshal(&w)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	// Verify both unknown elements survived round-trip
	s := string(out)
	if !containsAll(s, "unknown", "attr", "val", "text", "other", "data") {
		t.Errorf("round-trip lost content: %s", s)
	}
}

func TestRawXML_Empty(t *testing.T) {
	var raw RawXML
	out, err := xml.Marshal(&raw)
	if err != nil {
		t.Fatalf("Marshal empty: %v", err)
	}
	if len(out) != 0 {
		t.Errorf("Marshal empty RawXML = %q, want empty", out)
	}
}

func TestRawXML_InContainerStruct(t *testing.T) {
	type container struct {
		XMLName xml.Name `xml:"root"`
		Name    string   `xml:"name"`
		Extra   []RawXML `xml:",any"`
	}
	input := `<root><name>hello</name><unknown1>data1</unknown1><unknown2 a="b"/></root>`

	var c container
	err := xml.Unmarshal([]byte(input), &c)
	if err != nil {
		t.Fatalf("Unmarshal container: %v", err)
	}
	if c.Name != "hello" {
		t.Errorf("Name = %q, want %q", c.Name, "hello")
	}
	if len(c.Extra) != 2 {
		t.Fatalf("Extra count = %d, want 2", len(c.Extra))
	}
}

func TestRawXML_PreservesAttributes(t *testing.T) {
	type wrapper struct {
		XMLName xml.Name `xml:"root"`
		Extra   []RawXML `xml:",any"`
	}
	input := `<root><elem foo="bar" baz="qux">inner</elem></root>`

	var w wrapper
	if err := xml.Unmarshal([]byte(input), &w); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}

	out, err := xml.Marshal(&w)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	s := string(out)
	if !containsAll(s, "foo", "bar", "baz", "qux", "inner") {
		t.Errorf("lost attributes or content: %s", s)
	}
}

func ExampleRawXML() {
	type Doc struct {
		XMLName xml.Name `xml:"doc"`
		Title   string   `xml:"title"`
		Extra   []RawXML `xml:",any"` // Captures unknown elements
	}

	input := `<doc><title>Test</title><custom>preserved</custom></doc>`
	var d Doc
	_ = xml.Unmarshal([]byte(input), &d)

	out, _ := xml.Marshal(&d)
	_ = string(out) // Contains both <title> and <custom>
}

func containsAll(s string, substrs ...string) bool {
	for _, sub := range substrs {
		if !strings.Contains(s, sub) {
			return false
		}
	}
	return true
}
