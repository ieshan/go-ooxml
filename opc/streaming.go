package opc

import (
	"archive/zip"
	"fmt"
	"io"
	"path"
	"strings"
)

// binaryExtensions is the set of file extensions treated as binary (lazy-loaded) in streaming mode.
var binaryExtensions = map[string]bool{
	"png":  true,
	"jpg":  true,
	"jpeg": true,
	"gif":  true,
	"bmp":  true,
	"tiff": true,
	"emf":  true,
	"wmf":  true,
}

// isBinaryPart returns true if the part should be lazy-loaded in streaming mode.
// A part is binary if its content type contains "image" or its extension is in binaryExtensions.
func isBinaryPart(name, contentType string) bool {
	if strings.Contains(contentType, "image") {
		return true
	}
	ext := strings.TrimPrefix(path.Ext(name), ".")
	ext = strings.ToLower(ext)
	return binaryExtensions[ext]
}

// OpenStreaming opens an OPC package in streaming mode.
// XML parts are fully loaded; binary parts are lazy-loaded on first access.
// The io.ReaderAt must remain valid until all lazy parts have been read.
func OpenStreaming(r io.ReaderAt, size int64, opts *OpenOptions) (*Package, error) {
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

	// First pass: security checks and eager-load non-binary entries.
	// Binary entries get a lazy reader function.
	type entry struct {
		data   []byte    // non-nil when eagerly loaded
		zipRef *zip.File // non-nil when lazy
		lazy   bool
	}
	entries := make(map[string]*entry, len(zr.File))
	var totalDecompressed int64

	for _, f := range zr.File {
		name := f.Name

		if err := validateEntryPath(name); err != nil {
			return nil, err
		}

		if err := checkDuplicate(name, seen); err != nil {
			return nil, err
		}

		if err := checkCompressionRatio(int64(f.CompressedSize64), int64(f.UncompressedSize64), limits); err != nil {
			return nil, err
		}

		if int64(f.UncompressedSize64) > limits.MaxPartSize {
			return nil, &SecurityError{
				Check:  "part_size",
				Detail: fmt.Sprintf("part %q size %d exceeds limit %d", name, f.UncompressedSize64, limits.MaxPartSize),
				Err:    ErrFileTooLarge,
			}
		}

		totalDecompressed += int64(f.UncompressedSize64)
		if totalDecompressed > limits.MaxDecompressedSize {
			return nil, &SecurityError{
				Check:  "total_size",
				Detail: fmt.Sprintf("total decompressed size %d exceeds limit %d", totalDecompressed, limits.MaxDecompressedSize),
				Err:    ErrZipBomb,
			}
		}

		entries[name] = &entry{zipRef: f}
	}

	pkg := &Package{
		Parts: make(map[string]*Part),
	}

	// Parse [Content_Types].xml eagerly (always XML).
	if e, ok := entries["[Content_Types].xml"]; ok {
		data, err := readZipEntry(e.zipRef)
		if err != nil {
			return nil, fmt.Errorf("read [Content_Types].xml: %w", err)
		}
		ct, err := ParseContentTypes(data)
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

	// Parse _rels/.rels eagerly.
	if e, ok := entries["_rels/.rels"]; ok {
		data, err := readZipEntry(e.zipRef)
		if err != nil {
			return nil, fmt.Errorf("read _rels/.rels: %w", err)
		}
		rels, err := ParseRelationships(data)
		if err != nil {
			return nil, fmt.Errorf("parse _rels/.rels: %w", err)
		}
		pkg.Rels = rels
	}

	// Collect part-level rels (always XML, always eager).
	partRels := make(map[string][]Relationship)

	for name, e := range entries {
		if name == "[Content_Types].xml" || name == "_rels/.rels" {
			continue
		}
		if !isRelsPath(name) {
			continue
		}
		data, err := readZipEntry(e.zipRef)
		if err != nil {
			return nil, fmt.Errorf("read %q: %w", name, err)
		}
		rels, err := ParseRelationships(data)
		if err != nil {
			return nil, fmt.Errorf("parse %q: %w", name, err)
		}
		sourcePath := relsToSourcePath(name)
		if sourcePath != "" {
			partRels[sourcePath] = rels
		}
	}

	// Second pass: build Parts — binary parts are lazy, XML parts are eager.
	for name, e := range entries {
		if name == "[Content_Types].xml" || name == "_rels/.rels" {
			continue
		}
		if isRelsPath(name) {
			continue
		}

		contentType := pkg.ContentTypes.ContentTypeFor(name)

		if isBinaryPart(name, contentType) {
			// Capture loop variable for the closure.
			zipFile := e.zipRef
			pkg.Parts[name] = &Part{
				Path:        name,
				ContentType: contentType,
				loaded:      false,
				reader: func() (io.ReadCloser, error) {
					return zipFile.Open()
				},
			}
		} else {
			// Eager load for XML and other text parts.
			data, err := readZipEntry(e.zipRef)
			if err != nil {
				return nil, fmt.Errorf("read %q: %w", name, err)
			}
			pkg.Parts[name] = &Part{
				Path:        name,
				ContentType: contentType,
				data:        data,
				loaded:      true,
			}
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

// readZipEntry reads all bytes from a zip.File entry.
func readZipEntry(f *zip.File) ([]byte, error) {
	rc, err := f.Open()
	if err != nil {
		return nil, err
	}
	defer rc.Close()
	return io.ReadAll(rc)
}
