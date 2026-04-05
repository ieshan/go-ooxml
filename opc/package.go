package opc

import (
	"archive/zip"
	"fmt"
	"io"
	"os"
	"path"
	"strings"
	"sync"
)

// OpenOptions configures how a package is opened.
type OpenOptions struct {
	Security *SecurityLimits // nil = DefaultSecurityLimits()
}

// Package represents an OPC (Open Packaging Conventions) package.
type Package struct {
	mu           sync.RWMutex
	Parts        map[string]*Part
	ContentTypes *ContentTypes
	Rels         []Relationship // package-level rels (_rels/.rels)
}

// NewPackage creates a new empty Package with default content types.
func NewPackage() *Package {
	return &Package{
		Parts: make(map[string]*Part),
		ContentTypes: &ContentTypes{
			Defaults:  make(map[string]string),
			Overrides: make(map[string]string),
		},
	}
}

// Open reads a ZIP archive from r and returns a Package.
func Open(r io.ReaderAt, size int64, opts *OpenOptions) (*Package, error) {
	limits := DefaultSecurityLimits()
	if opts != nil && opts.Security != nil {
		limits = opts.Security
	}

	zr, err := zip.NewReader(r, size)
	if err != nil {
		return nil, &SecurityError{Check: "zip_read", Detail: err.Error(), Err: ErrCorruptZip}
	}

	// Validate file count.
	if len(zr.File) > limits.MaxFileCount {
		return nil, &SecurityError{
			Check:  "file_count",
			Detail: fmt.Sprintf("zip contains %d entries, limit is %d", len(zr.File), limits.MaxFileCount),
			Err:    ErrTooManyFiles,
		}
	}

	seen := make(map[string]bool)
	// First pass: read all entries into memory with security checks.
	entries := make(map[string][]byte, len(zr.File))
	var totalDecompressed int64

	for _, f := range zr.File {
		name := f.Name

		// Validate path.
		if err := validateEntryPath(name); err != nil {
			return nil, err
		}

		// Check duplicates.
		if err := checkDuplicate(name, seen); err != nil {
			return nil, err
		}

		// Check compression ratio.
		if err := checkCompressionRatio(int64(f.CompressedSize64), int64(f.UncompressedSize64), limits); err != nil {
			return nil, err
		}

		// Check per-part size.
		if int64(f.UncompressedSize64) > limits.MaxPartSize {
			return nil, &SecurityError{
				Check:  "part_size",
				Detail: fmt.Sprintf("part %q size %d exceeds limit %d", name, f.UncompressedSize64, limits.MaxPartSize),
				Err:    ErrFileTooLarge,
			}
		}

		// Check total decompressed size.
		totalDecompressed += int64(f.UncompressedSize64)
		if totalDecompressed > limits.MaxDecompressedSize {
			return nil, &SecurityError{
				Check:  "total_size",
				Detail: fmt.Sprintf("total decompressed size %d exceeds limit %d", totalDecompressed, limits.MaxDecompressedSize),
				Err:    ErrZipBomb,
			}
		}

		rc, err := f.Open()
		if err != nil {
			return nil, fmt.Errorf("open %q: %w", name, err)
		}
		data, err := io.ReadAll(rc)
		rc.Close()
		if err != nil {
			return nil, fmt.Errorf("read %q: %w", name, err)
		}

		entries[name] = data
	}

	pkg := &Package{
		Parts: make(map[string]*Part),
	}

	// Parse [Content_Types].xml.
	if ctData, ok := entries["[Content_Types].xml"]; ok {
		ct, err := ParseContentTypes(ctData)
		if err != nil {
			return nil, fmt.Errorf("parse [Content_Types].xml: %w", err)
		}
		pkg.ContentTypes = ct
	} else {
		pkg.ContentTypes = &ContentTypes{
			Defaults:  make(map[string]string),
			Overrides: make(map[string]string),
		}
	}

	// Parse _rels/.rels (package-level relationships).
	if relsData, ok := entries["_rels/.rels"]; ok {
		rels, err := ParseRelationships(relsData)
		if err != nil {
			return nil, fmt.Errorf("parse _rels/.rels: %w", err)
		}
		pkg.Rels = rels
	}

	// Collect part-level rels files for second pass.
	partRels := make(map[string][]Relationship) // source path -> rels

	for name, data := range entries {
		if name == "[Content_Types].xml" || name == "_rels/.rels" {
			continue
		}
		if isRelsPath(name) {
			// Parse part-level rels.
			rels, err := ParseRelationships(data)
			if err != nil {
				return nil, fmt.Errorf("parse %q: %w", name, err)
			}
			sourcePath := relsToSourcePath(name)
			if sourcePath != "" {
				partRels[sourcePath] = rels
			}
			continue
		}

		// Regular part.
		contentType := pkg.ContentTypes.ContentTypeFor(name)
		pkg.Parts[name] = &Part{
			Path:        name,
			ContentType: contentType,
			data:        data,
			loaded:      true,
		}
	}

	// Attach part-level rels to their source parts.
	for sourcePath, rels := range partRels {
		if part, ok := pkg.Parts[sourcePath]; ok {
			part.Rels = rels
		}
	}

	return pkg, nil
}

// OpenFile opens an OPC package from a file path.
func OpenFile(filePath string, opts *OpenOptions) (*Package, error) {
	f, err := os.Open(filePath)
	if err != nil {
		return nil, err
	}
	defer f.Close()

	info, err := f.Stat()
	if err != nil {
		return nil, err
	}

	return Open(f, info.Size(), opts)
}

// countingWriter wraps an io.Writer and counts bytes written.
type countingWriter struct {
	w io.Writer
	n int64
}

func (cw *countingWriter) Write(p []byte) (int, error) {
	n, err := cw.w.Write(p)
	cw.n += int64(n)
	return n, err
}

// WriteTo writes the package as a ZIP archive to w and returns the number
// of bytes written. It implements [io.WriterTo].
func (p *Package) WriteTo(w io.Writer) (int64, error) {
	p.mu.RLock()
	defer p.mu.RUnlock()

	cw := &countingWriter{w: w}
	zw := zip.NewWriter(cw)

	// Write [Content_Types].xml.
	if p.ContentTypes != nil {
		ctData, err := p.ContentTypes.Marshal()
		if err != nil {
			return cw.n, fmt.Errorf("marshal [Content_Types].xml: %w", err)
		}
		fw, err := zw.Create("[Content_Types].xml")
		if err != nil {
			return cw.n, err
		}
		if _, err := fw.Write(ctData); err != nil {
			return cw.n, err
		}
	}

	// Write _rels/.rels.
	if len(p.Rels) > 0 {
		relsData, err := MarshalRelationships(p.Rels)
		if err != nil {
			return cw.n, fmt.Errorf("marshal _rels/.rels: %w", err)
		}
		fw, err := zw.Create("_rels/.rels")
		if err != nil {
			return cw.n, err
		}
		if _, err := fw.Write(relsData); err != nil {
			return cw.n, err
		}
	}

	// Write parts and their rels sidecars.
	for _, part := range p.Parts {
		data, err := part.Data()
		if err != nil {
			return cw.n, fmt.Errorf("read part %q: %w", part.Path, err)
		}
		fw, err := zw.Create(part.Path)
		if err != nil {
			return cw.n, err
		}
		if _, err := fw.Write(data); err != nil {
			return cw.n, err
		}

		// Write part-level rels sidecar if any.
		if len(part.Rels) > 0 {
			relsData, err := MarshalRelationships(part.Rels)
			if err != nil {
				return cw.n, fmt.Errorf("marshal rels for %q: %w", part.Path, err)
			}
			relsPath := sourceToRelsPath(part.Path)
			fw, err := zw.Create(relsPath)
			if err != nil {
				return cw.n, err
			}
			if _, err := fw.Write(relsData); err != nil {
				return cw.n, err
			}
		}
	}

	if err := zw.Close(); err != nil {
		return cw.n, err
	}
	return cw.n, nil
}

// Save writes the package as a ZIP archive to w.
func (p *Package) Save(w io.Writer) error {
	_, err := p.WriteTo(w)
	return err
}

// AddPart adds a new part to the package.
func (p *Package) AddPart(partPath, contentType string, data []byte) *Part {
	p.mu.Lock()
	defer p.mu.Unlock()

	part := &Part{
		Path:        partPath,
		ContentType: contentType,
		data:        data,
		loaded:      true,
	}
	p.Parts[partPath] = part
	return part
}

// AddRelationship adds a package-level relationship.
func (p *Package) AddRelationship(relType, target string) Relationship {
	p.mu.Lock()
	defer p.mu.Unlock()

	rel := Relationship{
		ID:     NextRelID(p.Rels),
		Type:   relType,
		Target: target,
	}
	p.Rels = append(p.Rels, rel)
	return rel
}

// isRelsPath returns true if the path is a relationship file path.
// e.g., "_rels/.rels" or "word/_rels/document.xml.rels"
func isRelsPath(name string) bool {
	dir, file := path.Split(name)
	if !strings.HasSuffix(file, ".rels") {
		return false
	}
	// The directory must end with "_rels/".
	return strings.HasSuffix(dir, "_rels/")
}

// relsToSourcePath converts a rels file path to its source part path.
// e.g., "word/_rels/document.xml.rels" → "word/document.xml"
// Returns "" for package-level rels ("_rels/.rels").
func relsToSourcePath(relsPath string) string {
	dir, file := path.Split(relsPath)
	// Remove trailing "_rels/"
	if !strings.HasSuffix(dir, "_rels/") {
		return ""
	}
	parentDir := strings.TrimSuffix(dir, "_rels/")
	// Remove .rels suffix from filename.
	sourceName := strings.TrimSuffix(file, ".rels")
	if sourceName == "" || sourceName == "." {
		return ""
	}
	return parentDir + sourceName
}

// sourceToRelsPath converts a part path to its rels file path.
// e.g., "word/document.xml" → "word/_rels/document.xml.rels"
func sourceToRelsPath(partPath string) string {
	dir, file := path.Split(partPath)
	return dir + "_rels/" + file + ".rels"
}
