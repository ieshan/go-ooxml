package docx

import (
	"bytes"
	"testing"
)

func TestDocument_Properties(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	props := doc.Properties()
	if props == nil {
		t.Fatal("nil properties")
	}
}

func TestProperties_Title(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	props := doc.Properties()
	props.SetTitle("Test Document")
	if props.Title() != "Test Document" {
		t.Errorf("title = %q", props.Title())
	}
}

func TestProperties_Author(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	props := doc.Properties()
	props.SetAuthor("Jane Doe")
	if props.Author() != "Jane Doe" {
		t.Errorf("author = %q", props.Author())
	}
}

func TestProperties_Description(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	props := doc.Properties()
	props.SetDescription("A test")
	if props.Description() != "A test" {
		t.Errorf("desc = %q", props.Description())
	}
}

func TestProperties_CreatedModified(t *testing.T) {
	doc, _ := New(nil)
	defer doc.Close()
	props := doc.Properties()
	// Created and Modified on a fresh document should be zero values (no timestamp set).
	// Just ensure they don't panic.
	_ = props.Created()
	_ = props.Modified()
}

func TestProperties_RoundTrip(t *testing.T) {
	doc, _ := New(nil)
	doc.Properties().SetTitle("My Doc")
	doc.Properties().SetAuthor("Author")

	var buf bytes.Buffer
	doc.Write(&buf)
	doc2, _ := OpenReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()), nil)
	defer doc2.Close()
	if doc2.Properties().Title() != "My Doc" {
		t.Error("title lost")
	}
	if doc2.Properties().Author() != "Author" {
		t.Error("author lost")
	}
}

func ExampleProperties_SetTitle() {
	doc, _ := New(nil)
	defer doc.Close()
	doc.Properties().SetTitle("Quarterly Report")
	doc.Properties().SetAuthor("Finance Team")
}
