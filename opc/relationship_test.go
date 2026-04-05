package opc

import "testing"

func TestParseRelationships(t *testing.T) {
	data := []byte(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
</Relationships>`)
	rels, err := ParseRelationships(data)
	if err != nil {
		t.Fatalf("ParseRelationships: %v", err)
	}
	if len(rels) != 2 {
		t.Fatalf("len = %d, want 2", len(rels))
	}
	if rels[0].ID != "rId1" || rels[0].Target != "word/document.xml" {
		t.Errorf("rels[0] = %+v", rels[0])
	}
	if rels[1].ID != "rId2" || rels[1].Target != "docProps/core.xml" {
		t.Errorf("rels[1] = %+v", rels[1])
	}
}

func TestParseRelationships_External(t *testing.T) {
	data := []byte(`<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>
</Relationships>`)
	rels, _ := ParseRelationships(data)
	if !rels[0].External {
		t.Error("expected External=true")
	}
}

func TestMarshalRelationships_RoundTrip(t *testing.T) {
	rels := []Relationship{{ID: "rId1", Type: RelOfficeDocument, Target: "word/document.xml"}}
	data, err := MarshalRelationships(rels)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	rels2, err := ParseRelationships(data)
	if err != nil {
		t.Fatalf("round-trip: %v", err)
	}
	if len(rels2) != 1 || rels2[0].ID != "rId1" {
		t.Errorf("round-trip: %+v", rels2)
	}
}

func TestFindRelByType(t *testing.T) {
	rels := []Relationship{
		{ID: "rId1", Type: RelOfficeDocument, Target: "word/document.xml"},
		{ID: "rId2", Type: RelStyles, Target: "styles.xml"},
	}
	found := FindRelByType(rels, RelStyles)
	if found == nil || found.ID != "rId2" {
		t.Errorf("FindRelByType = %v", found)
	}
	if FindRelByType(rels, "nonexistent") != nil {
		t.Error("should return nil")
	}
}

func ExampleParseRelationships() {
	data := []byte(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`)
	rels, _ := ParseRelationships(data)
	_ = rels[0].Target // "word/document.xml"
}
