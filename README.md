<div align="center">
  <img src="./mbadocx.svg" alt="Mbadocx - Go DOCX Library" width="250">
  
  **Go library for creating, reading and manipulating DOCX files**
  
  ![Go Version](https://img.shields.io/badge/Go-1.16+-00ADD8)
 [![Go Reference](https://pkg.go.dev/badge/github.com/didikprabowo/mbadocx.svg)](https://pkg.go.dev/github.com/didikprabowo/mbadocx)
  [![Go Report Card](https://goreportcard.com/badge/github.com/didikprabowo/mbadocx)](https://goreportcard.com/report/github.com/didikprabowo/mbadocx)
  [![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
  
</div>

## Overview

**Mbadocx** is a Go library for programmatically creating and manipulating Microsoft Word (DOCX) documents.  
It provides a modular, extensible API for generating Word documents with advanced formatting, metadata, and resource management.

### Key Features

- ✅ **Document Creation** - Create new DOCX documents from scratch
- ✅ **Text Formatting** - Bold, italic, colors, font sizes
- ✅ **Paragraphs** - Full paragraph management and styling
- ✅ **Tables** - Create and format tables with rows and cells
- ✅ **Hyperlinks** - Add clickable links to documents
- ✅ **Line Breaks** - Control document flow
- ✅ **Fluent API** - Chainable methods for clean code
- ✅ **Pure Go** - No external dependencies required

### Examples
- [Basic Document](./example/basic)
- [List Numbering](./example/list)
- [Text Formating](./example/text-formatting)
- [Tables](./example/table)

### Why mbadocx?

- Simple API - Intuitive and easy to learn
- Idiomatic Go - Follows Go best practices and conventions
- Lightweight - Minimal overhead and dependencies
- Extensible - Modular architecture for easy extension
- MIT Licensed - Use freely in commercial and open-source projects

## Installation

### Requirements
- Go 1.16 or higher
- No external dependencies required

### Install via go get

```bash
go get github.com/didikprabowo/mbadocx
```

### Import in your code

```go
import "github.com/didikprabowo/mbadocx"
```

## Quick Start

Here's a simple example to get you started:

```go
package main

import (
    "log"
    "github.com/didikprabowo/mbadocx"
)

func main() {
    // Create a new document
    doc := mbadocx.New()
    
    // Add a paragraph with formatted text
    para := doc.AddParagraph()
    para.AddText("Hello, World!").SetBold(true).SetFontSize(16)
    
    // Save the document
    if err := doc.Save("hello.docx"); err != nil {
        log.Fatal(err)
    }
}
```

## Contributing

Contributions are welcome!  
Please see [CONTRIBUTING.md](./CONTRIBUTING.md) for guidelines.

---

## License

MIT License. See [LICENSE](./LICENSE) for details.