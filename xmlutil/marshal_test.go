package xmlutil

import (
	"encoding/xml"
	"errors"
	"strings"
	"testing"
)

type testElement struct {
	XMLName xml.Name `xml:"element"`
	Text    string   `xml:"text"`
}

func TestUnmarshal_Valid(t *testing.T) {
	input := `<element><text>hello</text></element>`
	var el testElement
	err := Unmarshal([]byte(input), &el)
	if err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if el.Text != "hello" {
		t.Errorf("Text = %q, want %q", el.Text, "hello")
	}
}

func TestUnmarshal_RejectsDTD(t *testing.T) {
	input := `<?xml version="1.0"?><!DOCTYPE foo [<!ENTITY xxe "bad">]><element><text>&xxe;</text></element>`
	var el testElement
	err := Unmarshal([]byte(input), &el)
	if err == nil {
		t.Fatal("expected error for DTD, got nil")
	}
	if !errors.Is(err, ErrXXE) {
		t.Errorf("error = %v, want ErrXXE", err)
	}
}

func TestUnmarshal_RejectsDeepNesting(t *testing.T) {
	var b strings.Builder
	for range 300 {
		b.WriteString("<a>")
	}
	b.WriteString("x")
	for range 300 {
		b.WriteString("</a>")
	}
	var result struct {
		XMLName xml.Name `xml:"a"`
	}
	err := Unmarshal([]byte(b.String()), &result)
	if err == nil {
		t.Fatal("expected error for deep nesting, got nil")
	}
	if !errors.Is(err, ErrNestingDepth) {
		t.Errorf("error = %v, want ErrNestingDepth", err)
	}
}

func TestUnmarshal_WithOptions(t *testing.T) {
	input := `<element><text>ok</text></element>`
	var el testElement
	err := Unmarshal([]byte(input), &el, WithMaxNestingDepth(10))
	if err != nil {
		t.Fatalf("Unmarshal with options: %v", err)
	}
}

func TestUnmarshal_RejectsTokenBomb(t *testing.T) {
	// Many sibling elements to exceed token count
	var b strings.Builder
	b.WriteString("<root>")
	for range 100 {
		b.WriteString("<x/>")
	}
	b.WriteString("</root>")
	var result struct {
		XMLName xml.Name `xml:"root"`
	}
	// Set very low token limit
	err := Unmarshal([]byte(b.String()), &result, WithMaxTokenCount(50))
	if err == nil {
		t.Fatal("expected error for token bomb")
	}
	if !errors.Is(err, ErrTokenLimit) {
		t.Errorf("error = %v, want ErrTokenLimit", err)
	}
}

func TestMarshal_Basic(t *testing.T) {
	el := testElement{Text: "hello"}
	out, err := Marshal(&el, nil)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	if !strings.Contains(string(out), "hello") {
		t.Errorf("output missing text: %s", out)
	}
}

func TestMarshal_WithNamespaceRewrite(t *testing.T) {
	type nsElement struct {
		XMLName xml.Name `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main document"`
		Body    string   `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main body"`
	}
	el := nsElement{Body: "content"}

	out, err := Marshal(&el, OOXML)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	s := string(out)
	// Should contain prefixed element names
	if !strings.Contains(s, "w:document") {
		t.Errorf("expected w:document prefix in output: %s", s)
	}
	if !strings.Contains(s, "w:body") {
		t.Errorf("expected w:body prefix in output: %s", s)
	}
	if !strings.Contains(s, "xmlns:w=") {
		t.Errorf("expected xmlns:w declaration in output: %s", s)
	}
}

func TestMarshalIndent(t *testing.T) {
	el := testElement{Text: "hello"}
	out, err := MarshalIndent(&el, nil, "  ")
	if err != nil {
		t.Fatalf("MarshalIndent: %v", err)
	}
	if !strings.Contains(string(out), "hello") {
		t.Errorf("output missing text: %s", out)
	}
	if !strings.Contains(string(out), "\n") {
		t.Errorf("expected indented output with newlines: %s", out)
	}
}

func TestAddXMLHeader(t *testing.T) {
	data := []byte("<root/>")
	result := AddXMLHeader(data)
	if !strings.HasPrefix(string(result), "<?xml") {
		t.Errorf("missing XML header: %s", result)
	}
	// Already has header — don't duplicate
	result2 := AddXMLHeader(result)
	if strings.Count(string(result2), "<?xml") != 1 {
		t.Error("duplicated XML header")
	}
}

func ExampleUnmarshal() {
	type Item struct {
		XMLName xml.Name `xml:"item"`
		Name    string   `xml:"name"`
	}
	data := []byte(`<item><name>Widget</name></item>`)
	var item Item
	_ = Unmarshal(data, &item)
	// item.Name == "Widget"
}
