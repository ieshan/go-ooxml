package opc

import (
	"archive/zip"
	"bytes"
	"errors"
	"fmt"
	"testing"
)

func createTestZip(t *testing.T, files map[string][]byte) []byte {
	t.Helper()
	var buf bytes.Buffer
	w := zip.NewWriter(&buf)
	for name, data := range files {
		fw, err := w.Create(name)
		if err != nil {
			t.Fatalf("create %q: %v", name, err)
		}
		if _, err := fw.Write(data); err != nil {
			t.Fatalf("write %q: %v", name, err)
		}
	}
	if err := w.Close(); err != nil {
		t.Fatalf("close: %v", err)
	}
	return buf.Bytes()
}

func TestOpen_MinimalPackage(t *testing.T) {
	data := createTestZip(t, map[string][]byte{
		"[Content_Types].xml": []byte(`<?xml version="1.0"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
</Types>`),
		"_rels/.rels": []byte(`<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`),
		"word/document.xml": []byte(`<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body/></w:document>`),
	})
	pkg, err := Open(bytes.NewReader(data), int64(len(data)), nil)
	if err != nil {
		t.Fatalf("Open: %v", err)
	}
	if pkg.Parts["word/document.xml"] == nil {
		t.Error("missing document part")
	}
	if len(pkg.Rels) != 1 {
		t.Errorf("package rels = %d, want 1", len(pkg.Rels))
	}
}

func TestOpen_PathTraversal(t *testing.T) {
	data := createTestZip(t, map[string][]byte{
		"[Content_Types].xml": []byte(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`),
		"../evil.txt":         []byte("pwned"),
	})
	_, err := Open(bytes.NewReader(data), int64(len(data)), nil)
	if !errors.Is(err, ErrPathTraversal) {
		t.Errorf("err = %v, want ErrPathTraversal", err)
	}
}

func TestOpen_TooManyFiles(t *testing.T) {
	files := map[string][]byte{
		"[Content_Types].xml": []byte(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`),
	}
	for i := range 20 {
		files[fmt.Sprintf("file%d.xml", i)] = []byte("<x/>")
	}
	data := createTestZip(t, files)
	limits := DefaultSecurityLimits()
	limits.MaxFileCount = 10
	_, err := Open(bytes.NewReader(data), int64(len(data)), &OpenOptions{Security: limits})
	if !errors.Is(err, ErrTooManyFiles) {
		t.Errorf("err = %v, want ErrTooManyFiles", err)
	}
}

func TestPackage_Save_RoundTrip(t *testing.T) {
	original := createTestZip(t, map[string][]byte{
		"[Content_Types].xml": []byte(`<?xml version="1.0"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
</Types>`),
		"_rels/.rels": []byte(`<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`),
		"word/document.xml": []byte(`<doc>content</doc>`),
	})
	pkg, err := Open(bytes.NewReader(original), int64(len(original)), nil)
	if err != nil {
		t.Fatalf("Open: %v", err)
	}

	var buf bytes.Buffer
	if err := pkg.Save(&buf); err != nil {
		t.Fatalf("Save: %v", err)
	}

	pkg2, err := Open(bytes.NewReader(buf.Bytes()), int64(buf.Len()), nil)
	if err != nil {
		t.Fatalf("re-open: %v", err)
	}
	part := pkg2.Parts["word/document.xml"]
	if part == nil {
		t.Fatal("missing part after round-trip")
	}
	data, _ := part.Data()
	if !bytes.Contains(data, []byte("content")) {
		t.Errorf("lost content: %s", data)
	}
}

func TestPackage_AddPart(t *testing.T) {
	pkg := NewPackage()
	pkg.AddPart("word/test.xml", "application/xml", []byte("<test/>"))
	if pkg.Parts["word/test.xml"] == nil {
		t.Error("AddPart failed")
	}
}

func TestPackage_AddRelationship(t *testing.T) {
	pkg := NewPackage()
	rel := pkg.AddRelationship(RelOfficeDocument, "word/document.xml")
	if rel.ID == "" {
		t.Error("empty ID")
	}
	if len(pkg.Rels) != 1 {
		t.Error("not added")
	}
}

func TestOpen_ParsesPartRels(t *testing.T) {
	data := createTestZip(t, map[string][]byte{
		"[Content_Types].xml": []byte(`<?xml version="1.0"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
</Types>`),
		"_rels/.rels": []byte(`<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`),
		"word/document.xml": []byte(`<doc/>`),
		"word/_rels/document.xml.rels": []byte(`<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`),
		"word/styles.xml": []byte(`<styles/>`),
	})
	pkg, err := Open(bytes.NewReader(data), int64(len(data)), nil)
	if err != nil {
		t.Fatalf("Open: %v", err)
	}
	docPart := pkg.Parts["word/document.xml"]
	if docPart == nil {
		t.Fatal("missing document part")
	}
	if len(docPart.Rels) != 1 {
		t.Errorf("document rels = %d, want 1", len(docPart.Rels))
	}
	if docPart.Rels[0].Target != "styles.xml" {
		t.Errorf("rel target = %q", docPart.Rels[0].Target)
	}
}

func TestHelpers_isRelsPath(t *testing.T) {
	tests := []struct {
		path string
		want bool
	}{
		{"_rels/.rels", true},
		{"word/_rels/document.xml.rels", true},
		{"word/document.xml", false},
		{"_rels/other.rels", true},
	}
	for _, tt := range tests {
		if got := isRelsPath(tt.path); got != tt.want {
			t.Errorf("isRelsPath(%q) = %v, want %v", tt.path, got, tt.want)
		}
	}
}

func TestHelpers_relsToSourcePath(t *testing.T) {
	tests := []struct {
		rels, want string
	}{
		{"word/_rels/document.xml.rels", "word/document.xml"},
		{"_rels/.rels", ""},
		{"ppt/_rels/presentation.xml.rels", "ppt/presentation.xml"},
	}
	for _, tt := range tests {
		if got := relsToSourcePath(tt.rels); got != tt.want {
			t.Errorf("relsToSourcePath(%q) = %q, want %q", tt.rels, got, tt.want)
		}
	}
}

func TestHelpers_sourceToRelsPath(t *testing.T) {
	tests := []struct {
		source, want string
	}{
		{"word/document.xml", "word/_rels/document.xml.rels"},
		{"ppt/presentation.xml", "ppt/_rels/presentation.xml.rels"},
	}
	for _, tt := range tests {
		if got := sourceToRelsPath(tt.source); got != tt.want {
			t.Errorf("sourceToRelsPath(%q) = %q, want %q", tt.source, got, tt.want)
		}
	}
}

func ExampleOpen() {
	// data := readDocxFile("report.docx")
	// pkg, err := Open(bytes.NewReader(data), int64(len(data)), nil)
	// docPart := pkg.Parts["word/document.xml"]
	_ = func() {}
}

func TestPackage_WriteTo(t *testing.T) {
	pkg := NewPackage()
	pkg.ContentTypes.Defaults["xml"] = "application/xml"
	pkg.AddPart("test.xml", "application/xml", []byte("<root/>"))

	var buf bytes.Buffer
	n, err := pkg.WriteTo(&buf)
	if err != nil {
		t.Fatalf("WriteTo: %v", err)
	}
	if n <= 0 {
		t.Errorf("WriteTo returned %d bytes, want > 0", n)
	}
	if int64(buf.Len()) != n {
		t.Errorf("buf.Len() = %d, WriteTo returned %d", buf.Len(), n)
	}

	// Verify round-trip.
	pkg2, err := Open(bytes.NewReader(buf.Bytes()), int64(buf.Len()), nil)
	if err != nil {
		t.Fatalf("Open after WriteTo: %v", err)
	}
	if _, ok := pkg2.Parts["test.xml"]; !ok {
		t.Error("part not found after round-trip")
	}
}
