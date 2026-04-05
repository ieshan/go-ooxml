package xmlutil

import "testing"

func TestNewRegistry(t *testing.T) {
	r := NewRegistry(
		Ns("w", "http://example.com/w"),
		Ns("r", "http://example.com/r"),
	)
	if r == nil {
		t.Fatal("NewRegistry returned nil")
	}
}

func TestRegistryPrefixToURI(t *testing.T) {
	r := NewRegistry(Ns("w", "http://example.com/w"))
	uri, ok := r.URI("w")
	if !ok || uri != "http://example.com/w" {
		t.Errorf("URI(w) = %q, %v; want %q, true", uri, ok, "http://example.com/w")
	}
}

func TestRegistryURIToPrefix(t *testing.T) {
	r := NewRegistry(Ns("w", "http://example.com/w"))
	prefix, ok := r.Prefix("http://example.com/w")
	if !ok || prefix != "w" {
		t.Errorf("Prefix() = %q, %v; want %q, true", prefix, ok, "w")
	}
}

func TestRegistryMiss(t *testing.T) {
	r := NewRegistry(Ns("w", "http://example.com/w"))
	_, ok := r.URI("x")
	if ok {
		t.Error("URI(x) should return false for unknown prefix")
	}
	_, ok = r.Prefix("http://example.com/unknown")
	if ok {
		t.Error("Prefix() should return false for unknown URI")
	}
}

func TestRegistryPrefixes(t *testing.T) {
	r := NewRegistry(
		Ns("w", "http://example.com/w"),
		Ns("r", "http://example.com/r"),
	)
	count := 0
	for range r.Prefixes() {
		count++
	}
	if count != 2 {
		t.Errorf("Prefixes() yielded %d, want 2", count)
	}
}

func TestOOXMLRegistry(t *testing.T) {
	uri, ok := OOXML.URI("w")
	if !ok {
		t.Fatal("OOXML registry missing 'w' prefix")
	}
	if uri != NsWordprocessingML {
		t.Errorf("OOXML w = %q, want %q", uri, NsWordprocessingML)
	}
	uri, ok = OOXML.URI("r")
	if !ok {
		t.Fatal("OOXML registry missing 'r' prefix")
	}
	if uri != NsRelationships {
		t.Errorf("OOXML r = %q, want %q", uri, NsRelationships)
	}
}

func ExampleNewRegistry() {
	reg := NewRegistry(
		Ns("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main"),
	)
	uri, _ := reg.URI("w")
	_ = uri
}
