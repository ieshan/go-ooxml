package opc

import (
	"bytes"
	"io"
)

// Part represents a single part (file) within an OPC package.
type Part struct {
	Path        string
	ContentType string
	Rels        []Relationship // relationships for this part
	data        []byte
	reader      func() (io.ReadCloser, error) // lazy accessor for streaming mode
	loaded      bool
}

// Data returns the part's content as bytes. Lazy-loads in streaming mode.
func (p *Part) Data() ([]byte, error) {
	if p.loaded {
		return p.data, nil
	}
	if p.reader == nil {
		return nil, &PartError{Path: p.Path, Op: "read", Err: ErrMissingPart}
	}
	rc, err := p.reader()
	if err != nil {
		return nil, &PartError{Path: p.Path, Op: "read", Err: err}
	}
	defer rc.Close()
	data, err := io.ReadAll(rc)
	if err != nil {
		return nil, &PartError{Path: p.Path, Op: "read", Err: err}
	}
	p.data = data
	p.loaded = true
	return p.data, nil
}

// IsLoaded reports whether the part's data has been loaded into memory.
func (p *Part) IsLoaded() bool { return p.loaded }

// SetData replaces the part's content with the given bytes.
func (p *Part) SetData(data []byte) {
	p.data = data
	p.loaded = true
}

// Reader returns an io.ReadCloser for the part's content.
func (p *Part) Reader() (io.ReadCloser, error) {
	if p.loaded {
		return io.NopCloser(bytes.NewReader(p.data)), nil
	}
	if p.reader != nil {
		return p.reader()
	}
	return nil, &PartError{Path: p.Path, Op: "read", Err: ErrMissingPart}
}
