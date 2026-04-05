package opc

import (
	"errors"
	"testing"
)

func TestValidatePath_Clean(t *testing.T) {
	clean := []string{"word/document.xml", "_rels/.rels", "[Content_Types].xml", "word/media/image1.png"}
	for _, p := range clean {
		if err := validateEntryPath(p); err != nil {
			t.Errorf("validateEntryPath(%q) = %v, want nil", p, err)
		}
	}
}

func TestValidatePath_Traversal(t *testing.T) {
	bad := []string{"../etc/passwd", "word/../../../etc/passwd", "word/..\\secret", "/absolute/path", "word/\x00null.xml"}
	for _, p := range bad {
		err := validateEntryPath(p)
		if err == nil {
			t.Errorf("validateEntryPath(%q) = nil, want error", p)
			continue
		}
		if !errors.Is(err, ErrPathTraversal) {
			t.Errorf("validateEntryPath(%q) = %v, want ErrPathTraversal", p, err)
		}
	}
}

func TestValidatePath_Backslash(t *testing.T) {
	err := validateEntryPath("word\\document.xml")
	if !errors.Is(err, ErrPathTraversal) {
		t.Errorf("backslash path error = %v, want ErrPathTraversal", err)
	}
}

func TestDefaultSecurityLimits(t *testing.T) {
	limits := DefaultSecurityLimits()
	if limits.MaxDecompressedSize <= 0 {
		t.Error("MaxDecompressedSize should be positive")
	}
	if limits.MaxFileCount <= 0 {
		t.Error("MaxFileCount should be positive")
	}
	if limits.MaxPartSize <= 0 {
		t.Error("MaxPartSize should be positive")
	}
	if limits.MaxCompressionRatio <= 0 {
		t.Error("MaxCompressionRatio should be positive")
	}
}

func TestCheckCompressionRatio(t *testing.T) {
	limits := DefaultSecurityLimits()
	err := checkCompressionRatio(100, 1000, limits)
	if err != nil {
		t.Errorf("normal ratio: %v", err)
	}
	err = checkCompressionRatio(1, 1000, limits)
	if !errors.Is(err, ErrZipBomb) {
		t.Errorf("extreme ratio: %v, want ErrZipBomb", err)
	}
}

func TestCheckDuplicatePaths(t *testing.T) {
	seen := make(map[string]bool)
	if err := checkDuplicate("word/document.xml", seen); err != nil {
		t.Errorf("first entry: %v", err)
	}
	err := checkDuplicate("word/document.xml", seen)
	if err == nil {
		t.Error("duplicate entry should error")
	}
}

func TestCheckDuplicatePaths_CaseInsensitive(t *testing.T) {
	seen := make(map[string]bool)
	_ = checkDuplicate("Word/Document.xml", seen)
	err := checkDuplicate("word/document.xml", seen)
	if err == nil {
		t.Error("case-insensitive duplicate should error")
	}
}

func TestSecurityError_Unwrap(t *testing.T) {
	se := &SecurityError{Check: "test", Detail: "detail", Err: ErrZipBomb}
	if !errors.Is(se, ErrZipBomb) {
		t.Error("SecurityError should unwrap to ErrZipBomb")
	}
}

func TestPartError_Unwrap(t *testing.T) {
	pe := &PartError{Path: "/word/doc.xml", Op: "unmarshal", Err: ErrMissingPart}
	if !errors.Is(pe, ErrMissingPart) {
		t.Error("PartError should unwrap to ErrMissingPart")
	}
}

func ExampleDefaultSecurityLimits() {
	limits := DefaultSecurityLimits()
	_ = limits.MaxDecompressedSize // 500 * 1024 * 1024
	_ = limits.MaxFileCount        // 10000
}
