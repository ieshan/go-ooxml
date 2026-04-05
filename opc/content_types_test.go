package opc

import "testing"

func TestContentTypes_Parse(t *testing.T) {
	data := []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`)
	ct, err := ParseContentTypes(data)
	if err != nil {
		t.Fatalf("ParseContentTypes: %v", err)
	}
	if got := ct.Defaults["rels"]; got != "application/vnd.openxmlformats-package.relationships+xml" {
		t.Errorf("Defaults[rels] = %q", got)
	}
	if got := ct.Defaults["xml"]; got != "application/xml" {
		t.Errorf("Defaults[xml] = %q", got)
	}
	if got := ct.Overrides["/word/document.xml"]; got != "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml" {
		t.Errorf("Overrides = %q", got)
	}
}

func TestContentTypes_ContentTypeFor(t *testing.T) {
	ct := &ContentTypes{
		Defaults:  map[string]string{"xml": "application/xml", "png": "image/png"},
		Overrides: map[string]string{"/word/document.xml": "application/vnd.custom+xml"},
	}
	tests := []struct{ path, want string }{
		{"/word/document.xml", "application/vnd.custom+xml"},
		{"/word/styles.xml", "application/xml"},
		{"/word/media/image1.png", "image/png"},
		{"/unknown.xyz", ""},
	}
	for _, tt := range tests {
		got := ct.ContentTypeFor(tt.path)
		if got != tt.want {
			t.Errorf("ContentTypeFor(%q) = %q, want %q", tt.path, got, tt.want)
		}
	}
}

func TestContentTypes_Marshal_RoundTrip(t *testing.T) {
	ct := &ContentTypes{
		Defaults:  map[string]string{"xml": "application/xml"},
		Overrides: map[string]string{"/word/document.xml": "application/vnd.custom+xml"},
	}
	data, err := ct.Marshal()
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	if len(data) == 0 {
		t.Fatal("empty")
	}
	ct2, err := ParseContentTypes(data)
	if err != nil {
		t.Fatalf("round-trip: %v", err)
	}
	if ct2.Defaults["xml"] != "application/xml" {
		t.Error("round-trip lost default")
	}
	if ct2.Overrides["/word/document.xml"] != "application/vnd.custom+xml" {
		t.Error("round-trip lost override")
	}
}

func ExampleContentTypes_ContentTypeFor() {
	ct := &ContentTypes{
		Defaults:  map[string]string{"xml": "application/xml"},
		Overrides: map[string]string{"/word/document.xml": "application/vnd.custom+xml"},
	}
	_ = ct.ContentTypeFor("/word/document.xml") // "application/vnd.custom+xml"
	_ = ct.ContentTypeFor("/word/styles.xml")   // "application/xml"
}
