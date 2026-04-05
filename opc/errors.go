// Package opc implements the Open Packaging Convention (OPC) used by
// Office Open XML formats. It handles reading and writing ZIP-based
// packages with content types, relationships, and security protections.
package opc

import (
	"errors"
	"fmt"
)

var (
	ErrZipBomb        = errors.New("ooxml: decompressed size exceeds limit")
	ErrPathTraversal  = errors.New("ooxml: zip entry contains path traversal")
	ErrTooManyFiles   = errors.New("ooxml: zip contains too many entries")
	ErrFileTooLarge   = errors.New("ooxml: part exceeds maximum size")
	ErrCorruptZip     = errors.New("ooxml: invalid or corrupt zip archive")
	ErrMissingPart    = errors.New("ooxml: required document part missing")
	ErrDuplicateEntry = errors.New("ooxml: zip contains duplicate entry paths")
)

// SecurityError provides context about a security check failure.
type SecurityError struct {
	Check  string // "zip_bomb", "path_traversal", "xxe", etc.
	Detail string
	Err    error // underlying sentinel error
}

func (e *SecurityError) Error() string {
	return fmt.Sprintf("ooxml: security check %s: %s", e.Check, e.Detail)
}

func (e *SecurityError) Unwrap() error { return e.Err }

// PartError provides context about a failure processing a specific part.
type PartError struct {
	Path string // e.g., "/word/document.xml"
	Op   string // "unmarshal", "marshal", "validate"
	Err  error
}

func (e *PartError) Error() string {
	return fmt.Sprintf("ooxml: %s %s: %s", e.Op, e.Path, e.Err)
}

func (e *PartError) Unwrap() error { return e.Err }
