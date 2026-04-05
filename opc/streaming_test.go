package opc

import (
	"bytes"
	"testing"
)

func TestOpenStreaming_XMLPartsLoaded(t *testing.T) {
	data := createTestZip(t, map[string][]byte{
		"[Content_Types].xml": []byte(`<?xml version="1.0"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="png" ContentType="image/png"/>
</Types>`),
		"_rels/.rels": []byte(`<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`),
		"word/document.xml":     []byte(`<doc>content</doc>`),
		"word/media/image1.png": []byte("FAKE_PNG_DATA"),
	})

	pkg, err := OpenStreaming(bytes.NewReader(data), int64(len(data)), nil)
	if err != nil {
		t.Fatalf("OpenStreaming: %v", err)
	}

	// XML part should be loaded eagerly.
	docPart := pkg.Parts["word/document.xml"]
	if docPart == nil {
		t.Fatal("missing document part")
	}
	if !docPart.IsLoaded() {
		t.Error("XML part should be loaded")
	}

	// Binary part should be lazy.
	imgPart := pkg.Parts["word/media/image1.png"]
	if imgPart == nil {
		t.Fatal("missing image part")
	}
	if imgPart.IsLoaded() {
		t.Error("image part should not be loaded yet")
	}

	// Accessing data should load it.
	imgData, err := imgPart.Data()
	if err != nil {
		t.Fatalf("Data: %v", err)
	}
	if string(imgData) != "FAKE_PNG_DATA" {
		t.Errorf("data = %q", imgData)
	}
	if !imgPart.IsLoaded() {
		t.Error("image part should be loaded after Data()")
	}
}

func TestOpenStreaming_SecurityChecks(t *testing.T) {
	data := createTestZip(t, map[string][]byte{
		"[Content_Types].xml": []byte(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`),
		"../evil.txt":         []byte("pwned"),
	})
	_, err := OpenStreaming(bytes.NewReader(data), int64(len(data)), nil)
	if err == nil {
		t.Error("should reject path traversal")
	}
}

func TestOpenStreaming_SaveRoundTrip(t *testing.T) {
	data := createTestZip(t, map[string][]byte{
		"[Content_Types].xml": []byte(`<?xml version="1.0"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="png" ContentType="image/png"/>
</Types>`),
		"_rels/.rels": []byte(`<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`),
		"word/document.xml":     []byte(`<doc>content</doc>`),
		"word/media/image1.png": []byte("PNG_DATA_HERE"),
	})

	pkg, err := OpenStreaming(bytes.NewReader(data), int64(len(data)), nil)
	if err != nil {
		t.Fatalf("OpenStreaming: %v", err)
	}

	// Save without accessing the image.
	var buf bytes.Buffer
	if err := pkg.Save(&buf); err != nil {
		t.Fatalf("Save: %v", err)
	}

	// Re-open and verify all parts survive.
	pkg2, err := Open(bytes.NewReader(buf.Bytes()), int64(buf.Len()), nil)
	if err != nil {
		t.Fatalf("re-open: %v", err)
	}
	imgPart := pkg2.Parts["word/media/image1.png"]
	if imgPart == nil {
		t.Fatal("image lost")
	}
	imgData, _ := imgPart.Data()
	if string(imgData) != "PNG_DATA_HERE" {
		t.Errorf("image data = %q", imgData)
	}
}
