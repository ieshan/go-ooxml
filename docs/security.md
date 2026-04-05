# Security

Office documents from untrusted sources can be attack vectors. This library applies security checks by default when opening any document.

## Protections

### ZIP Bomb Detection

`.docx` files are ZIP archives. A small ZIP can decompress into gigabytes of data. The library checks:

- **Compression ratio** — rejects entries where `uncompressed / compressed` exceeds the limit (default: 100:1)
- **Per-part size** — rejects individual parts larger than the limit (default: 100 MB)
- **Total decompressed size** — rejects archives whose total decompressed size exceeds the limit (default: 500 MB)
- **File count** — rejects archives with too many entries (default: 10,000)

### XXE Prevention

XML External Entity attacks embed references to local files or network resources in XML. Go's `encoding/xml` parser does not process DTDs or external entities, making it inherently safe. The library additionally rejects any document containing DTD declarations as an extra precaution.

### Path Traversal

Malicious ZIP entries can use paths like `../../etc/cron.d/evil` to write files outside the intended directory. The library validates every ZIP entry path:

- Rejects paths containing `..`
- Rejects absolute paths (starting with `/`)
- Rejects backslashes (`\`)
- Rejects null bytes

### Duplicate ZIP Entries

Some attacks rely on duplicate filenames in a ZIP to confuse parsers. The library detects and rejects duplicate entries (case-insensitive).

## Custom Security Limits

Override defaults when opening a document:

```go
import "github.com/ieshan/go-ooxml/opc"

limits := &opc.SecurityLimits{
	MaxDecompressedSize: 200 * 1024 * 1024,  // 200 MB
	MaxPartSize:         50 * 1024 * 1024,    // 50 MB
	MaxFileCount:        5000,
	MaxCompressionRatio: 50,
}

doc, err := docx.Open("untrusted.docx", &docx.Config{
	Security: limits,
})
```

The default limits are returned by `opc.DefaultSecurityLimits()`:

| Limit | Default |
|-------|---------|
| MaxDecompressedSize | 500 MB |
| MaxPartSize | 100 MB |
| MaxFileCount | 10,000 |
| MaxCompressionRatio | 100:1 |

## Error Handling

Security violations return an `*opc.SecurityError` with a `Check` field identifying which check failed:

```go
doc, err := docx.Open("suspicious.docx", nil)
if err != nil {
	var secErr *opc.SecurityError
	if errors.As(err, &secErr) {
		fmt.Printf("Security check failed: %s — %s\n", secErr.Check, secErr.Detail)
	}
}
```

Check values: `"zip_bomb"`, `"path_traversal"`, `"part_size"`, `"file_count"`, `"total_size"`, `"duplicate_entry"`.
