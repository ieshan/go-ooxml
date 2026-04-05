package opc

import (
	"archive/zip"
	"bytes"
	"errors"
	"fmt"
	"testing"
)

func TestSecurity_PathTraversal_DotDot(t *testing.T) {
	data := createTestZip(t, map[string][]byte{
		"[Content_Types].xml": []byte(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`),
		"../../../etc/passwd": []byte("root:x:0:0"),
	})
	_, err := Open(bytes.NewReader(data), int64(len(data)), nil)
	if !errors.Is(err, ErrPathTraversal) {
		t.Errorf("expected ErrPathTraversal, got %v", err)
	}
}

func TestSecurity_PathTraversal_Backslash(t *testing.T) {
	data := createTestZip(t, map[string][]byte{
		"[Content_Types].xml":    []byte(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`),
		"word\\..\\..\\evil.txt": []byte("evil"),
	})
	_, err := Open(bytes.NewReader(data), int64(len(data)), nil)
	if !errors.Is(err, ErrPathTraversal) {
		t.Errorf("expected ErrPathTraversal, got %v", err)
	}
}

func TestSecurity_PathTraversal_AbsolutePath(t *testing.T) {
	data := createTestZip(t, map[string][]byte{
		"[Content_Types].xml": []byte(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`),
		"/etc/passwd":         []byte("root:x:0:0"),
	})
	_, err := Open(bytes.NewReader(data), int64(len(data)), nil)
	if !errors.Is(err, ErrPathTraversal) {
		t.Errorf("expected ErrPathTraversal, got %v", err)
	}
}

func TestSecurity_PathTraversal_NullByte(t *testing.T) {
	data := createTestZip(t, map[string][]byte{
		"[Content_Types].xml": []byte(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`),
		"word/doc\x00.xml":    []byte("<doc/>"),
	})
	_, err := Open(bytes.NewReader(data), int64(len(data)), nil)
	if !errors.Is(err, ErrPathTraversal) {
		t.Errorf("expected ErrPathTraversal, got %v", err)
	}
}

func TestSecurity_TooManyFiles(t *testing.T) {
	files := map[string][]byte{
		"[Content_Types].xml": []byte(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`),
	}
	for i := range 50 {
		files[fmt.Sprintf("part%d.xml", i)] = []byte("<x/>")
	}
	data := createTestZip(t, files)

	limits := DefaultSecurityLimits()
	limits.MaxFileCount = 20
	_, err := Open(bytes.NewReader(data), int64(len(data)), &OpenOptions{Security: limits})
	if !errors.Is(err, ErrTooManyFiles) {
		t.Errorf("expected ErrTooManyFiles, got %v", err)
	}
}

func TestSecurity_FileTooLarge(t *testing.T) {
	largeContent := make([]byte, 1024*1024+1) // 1MB+1
	for i := range largeContent {
		largeContent[i] = 'A'
	}

	data := createTestZip(t, map[string][]byte{
		"[Content_Types].xml": []byte(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`),
		"large.xml":           largeContent,
	})

	limits := DefaultSecurityLimits()
	limits.MaxPartSize = 1024 * 1024    // 1MB limit
	limits.MaxCompressionRatio = 100000 // allow high compression so per-part check fires
	_, err := Open(bytes.NewReader(data), int64(len(data)), &OpenOptions{Security: limits})
	if !errors.Is(err, ErrFileTooLarge) {
		t.Errorf("expected ErrFileTooLarge, got %v", err)
	}
}

func TestSecurity_DuplicateEntries(t *testing.T) {
	var buf bytes.Buffer
	w := zip.NewWriter(&buf)

	ct := []byte(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	fw, _ := w.Create("[Content_Types].xml")
	fw.Write(ct)

	fw, _ = w.Create("Word/Document.xml")
	fw.Write([]byte("<doc/>"))

	fw, _ = w.Create("word/document.xml") // duplicate (case-insensitive)
	fw.Write([]byte("<doc2/>"))

	w.Close()

	_, err := Open(bytes.NewReader(buf.Bytes()), int64(buf.Len()), nil)
	if !errors.Is(err, ErrDuplicateEntry) {
		t.Errorf("expected ErrDuplicateEntry, got %v", err)
	}
}

func TestSecurity_XXE_DOCTYPE(t *testing.T) {
	data := createTestZip(t, map[string][]byte{
		"[Content_Types].xml": []byte(`<?xml version="1.0"?><!DOCTYPE foo [<!ENTITY xxe "bad">]><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`),
	})
	_, err := Open(bytes.NewReader(data), int64(len(data)), nil)
	if err == nil {
		t.Error("expected error for XXE in content types")
	}
}

func TestSecurity_DecompressedSizeLimit(t *testing.T) {
	files := map[string][]byte{
		"[Content_Types].xml": []byte(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`),
	}
	content := make([]byte, 512*1024) // 512KB each
	for i := range content {
		content[i] = byte('A' + (i % 26))
	}
	for i := range 5 {
		files[fmt.Sprintf("part%d.bin", i)] = content
	}
	data := createTestZip(t, files)

	limits := DefaultSecurityLimits()
	limits.MaxDecompressedSize = 1024 * 1024 // 1MB total limit
	_, err := Open(bytes.NewReader(data), int64(len(data)), &OpenOptions{Security: limits})
	if !errors.Is(err, ErrZipBomb) {
		t.Errorf("expected ErrZipBomb, got %v", err)
	}
}

func TestSecurity_SecurityError_Type(t *testing.T) {
	data := createTestZip(t, map[string][]byte{
		"[Content_Types].xml": []byte(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`),
		"../evil.txt":         []byte("pwned"),
	})
	_, err := Open(bytes.NewReader(data), int64(len(data)), nil)

	var se *SecurityError
	if !errors.As(err, &se) {
		t.Fatalf("expected SecurityError, got %T: %v", err, err)
	}
	if se.Check != "path_traversal" {
		t.Errorf("check = %q, want path_traversal", se.Check)
	}
}

func TestSecurity_ValidPackage_Passes(t *testing.T) {
	data := createTestZip(t, map[string][]byte{
		"[Content_Types].xml": []byte(`<?xml version="1.0"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
</Types>`),
		"_rels/.rels": []byte(`<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`),
		"word/document.xml": []byte(`<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body/></w:document>`),
	})
	pkg, err := Open(bytes.NewReader(data), int64(len(data)), nil)
	if err != nil {
		t.Fatalf("valid package should open: %v", err)
	}
	if pkg == nil {
		t.Fatal("nil package")
	}
}
