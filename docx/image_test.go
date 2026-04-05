package docx

import (
	"testing"
)

func TestRun_Images_Empty(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	r := doc.Body().AddParagraph().AddRun("no images")
	count := 0
	for range r.Images() {
		count++
	}
	if count != 0 {
		t.Error("should be empty")
	}
}

func TestImage_ContentType(t *testing.T) {
	img := &Image{contentType: "image/png"}
	if img.ContentType() != "image/png" {
		t.Errorf("ContentType = %q, want %q", img.ContentType(), "image/png")
	}
}
