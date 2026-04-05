package opc

import (
	"bytes"
	"io"
	"testing"
)

func TestPart_Data(t *testing.T) {
	p := &Part{Path: "test.xml", data: []byte("<test/>"), loaded: true}
	data, err := p.Data()
	if err != nil {
		t.Fatalf("Data: %v", err)
	}
	if string(data) != "<test/>" {
		t.Errorf("Data = %q", data)
	}
}

func TestPart_Data_LazyLoad(t *testing.T) {
	content := []byte("<lazy/>")
	p := &Part{
		Path:   "lazy.xml",
		loaded: false,
		reader: func() (io.ReadCloser, error) {
			return io.NopCloser(bytes.NewReader(content)), nil
		},
	}
	data, err := p.Data()
	if err != nil {
		t.Fatalf("Data: %v", err)
	}
	if string(data) != "<lazy/>" {
		t.Errorf("Data = %q, want %q", data, "<lazy/>")
	}
	if !p.loaded {
		t.Error("expected loaded to be true after Data()")
	}
}

func TestPart_Reader(t *testing.T) {
	p := &Part{Path: "test.xml", data: []byte("<test/>"), loaded: true}
	rc, err := p.Reader()
	if err != nil {
		t.Fatalf("Reader: %v", err)
	}
	defer rc.Close()
	got, err := io.ReadAll(rc)
	if err != nil {
		t.Fatalf("ReadAll: %v", err)
	}
	if string(got) != "<test/>" {
		t.Errorf("Reader content = %q", got)
	}
}

func TestPart_Reader_LazyLoad(t *testing.T) {
	content := []byte("<lazy/>")
	p := &Part{
		Path:   "lazy.xml",
		loaded: false,
		reader: func() (io.ReadCloser, error) {
			return io.NopCloser(bytes.NewReader(content)), nil
		},
	}
	rc, err := p.Reader()
	if err != nil {
		t.Fatalf("Reader: %v", err)
	}
	defer rc.Close()
	got, err := io.ReadAll(rc)
	if err != nil {
		t.Fatalf("ReadAll: %v", err)
	}
	if string(got) != "<lazy/>" {
		t.Errorf("Reader content = %q", got)
	}
}
