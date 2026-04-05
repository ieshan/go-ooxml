package opc

import (
	"fmt"
	"path"
	"strings"
)

// SecurityLimits configures thresholds for ZIP security checks.
type SecurityLimits struct {
	MaxDecompressedSize int64
	MaxPartSize         int64
	MaxFileCount        int
	MaxCompressionRatio int
	MaxNestingDepth     int
}

// DefaultSecurityLimits returns conservative defaults suitable for
// processing untrusted OOXML files.
func DefaultSecurityLimits() *SecurityLimits {
	return &SecurityLimits{
		MaxDecompressedSize: 500 * 1024 * 1024, // 500 MB
		MaxPartSize:         100 * 1024 * 1024, // 100 MB
		MaxFileCount:        10_000,
		MaxCompressionRatio: 100,
		MaxNestingDepth:     0,
	}
}

// validateEntryPath checks a ZIP entry path for path traversal, absolute
// paths, backslashes, and null bytes.
func validateEntryPath(p string) error {
	// Reject null bytes.
	if strings.ContainsRune(p, '\x00') {
		return &SecurityError{
			Check:  "path_traversal",
			Detail: fmt.Sprintf("path contains null byte: %q", p),
			Err:    ErrPathTraversal,
		}
	}

	// Reject backslashes (Windows path separator, not allowed in OPC).
	if strings.ContainsRune(p, '\\') {
		return &SecurityError{
			Check:  "path_traversal",
			Detail: fmt.Sprintf("path contains backslash: %q", p),
			Err:    ErrPathTraversal,
		}
	}

	// Reject absolute paths.
	if strings.HasPrefix(p, "/") {
		return &SecurityError{
			Check:  "path_traversal",
			Detail: fmt.Sprintf("path is absolute: %q", p),
			Err:    ErrPathTraversal,
		}
	}

	// Reject any path that attempts traversal via "..".
	// path.Clean normalises slashes and resolves ".." components.
	cleaned := path.Clean(p)
	if strings.HasPrefix(cleaned, "..") {
		return &SecurityError{
			Check:  "path_traversal",
			Detail: fmt.Sprintf("path traversal detected: %q", p),
			Err:    ErrPathTraversal,
		}
	}

	// Reject if any individual segment is "..".
	for _, seg := range strings.Split(p, "/") {
		if seg == ".." {
			return &SecurityError{
				Check:  "path_traversal",
				Detail: fmt.Sprintf("path traversal detected: %q", p),
				Err:    ErrPathTraversal,
			}
		}
	}

	return nil
}

// checkCompressionRatio returns ErrZipBomb when the ratio of
// uncompressedSize to compressedSize exceeds limits.MaxCompressionRatio.
// A compressedSize of 0 is treated as 1 to avoid division by zero.
func checkCompressionRatio(compressedSize, uncompressedSize int64, limits *SecurityLimits) error {
	if compressedSize <= 0 {
		compressedSize = 1
	}
	ratio := uncompressedSize / compressedSize
	if ratio > int64(limits.MaxCompressionRatio) {
		return &SecurityError{
			Check:  "zip_bomb",
			Detail: fmt.Sprintf("compression ratio %d exceeds limit %d", ratio, limits.MaxCompressionRatio),
			Err:    ErrZipBomb,
		}
	}
	return nil
}

// checkDuplicate checks whether path has already been seen (case-insensitively).
// If it has not been seen, it records the lower-cased path in seen and returns nil.
// If it has been seen, it returns ErrDuplicateEntry.
func checkDuplicate(p string, seen map[string]bool) error {
	key := strings.ToLower(p)
	if seen[key] {
		return &SecurityError{
			Check:  "duplicate_entry",
			Detail: fmt.Sprintf("duplicate zip entry: %q", p),
			Err:    ErrDuplicateEntry,
		}
	}
	seen[key] = true
	return nil
}
