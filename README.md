# Mbadocx

## Introduction	

**mbadocx** is a Go library for programmatically creating and manipulating Microsoft Word (DOCX) documents.  
It provides a modular, extensible API for generating Word documents with advanced formatting, metadata, and resource management.

## Features

### Document Creation & Structure
- [x] Create DOCX documents from scratch
- [x] Add paragraphs and runs
- [x] Section breaks and page setup
- [ ] Table support (rows, cells, cell merging)
- [ ] Image embedding (PNG, JPEG, etc.)


### Text & Paragraph Formatting
- [x] Apply text formatting (bold, italic, underline, font, color, size)
- [x] Set paragraph alignment (left, right, center, justify)
- [x] Set paragraph spacing (before, after, line spacing)
- [x] Set paragraph indentation (left, right, first line, hanging)
- [x] Add line breaks and page breaks
- [ ] Advanced numbering and lists (bullets, multi-level lists)
- [ ] Extended formatting (borders, shading, tabs, page breaks)

### Styles & Layout
- [x] Define and apply paragraph styles
- [x] Define and apply character styles
- [ ] Table styles
- [ ] Custom style definitions

### Content & Relationships
- [x] Add hyperlinks (internal, external)
- [ ] Add bookmarks
- [ ] Add footnotes and endnotes
- [ ] Custom properties (user-defined metadata)
- [ ] Embedded objects (charts, equations)

### Metadata & Validation
- [x] Manage document metadata (author, title, subject, timestamps)
- [ ] Custom metadata fields
- [ ] Document validation and 
- [ ] Compatibility checks (Word version, schema)

### API & Resource Management
- [x] Thread-safe API (safe for concurrent use)
- [x] Resource management and cleanup (file handles, memory)
- [x] Save to file or write to any `io.Writer`
- [ ] Edit existing DOCX files
- [ ] Import/export from other formats (HTML, Markdown, PDF)
- [ ] Plugin architecture for custom elements

### Testing & Documentation
- [x] Unit tests for core features
- [ ] Integration tests for document output
- [x] API documentation
- [x] Example usage and guides
- [ ] FAQ and troubleshooting

---

## Installation

```sh
go get github.com/didikprabowo/mbadocx
```

---

## Usage Example

```go
package main

import "github.com/didikprabowo/mbadocx"

func main() {
    doc := mbadocx.New()
    para := doc.GetBody().AddParagraph()
    para.AddText("Hello, mbadocx!").SetBold(true)
    doc.Save("output.docx")
}
```

## Contributing

Contributions are welcome!  
Please see [CONTRIBUTING.md](./CONTRIBUTING.md) for guidelines.

---

## License

MIT License. See [LICENSE](./LICENSE) for details.