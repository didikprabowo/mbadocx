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
- ✅ **Image** - Add image and set properties
- ✅ **Hyperlinks** - Add clickable links to documents
- ✅ **Line Breaks** - Control document flow
- ✅ **Fluent API** - Chainable methods for clean code
- ✅ **Pure Go** - No external dependencies required
- ⬜ Headers/Footers (if not implemented)

### Examples
- [Basic Document](./example/basic)
- [List Numbering](./example/list)
- [Text Formating](./example/text-formatting)
- [Tables](./example/table)
- [Image](./example/image)

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

### Creating Tables

```go
package main

import (
    "log"

    "github.com/didikprabowo/mbadocx"
    "github.com/didikprabowo/mbadocx/elements"
)

func main() {
    doc := mbadocx.New()
    // Add title
    doc.AddHeading("Table Examples in mbadocx", 1)
    doc.AddParagraph().AddText("This document demonstrates various table features.")

    // Example 1: Simple table
    doc.AddHeading("Example 1: Basic Table", 2)
    doc.AddParagraph().AddText("A simple 3x3 table:")

    table1 := doc.AddTable(3, 3)
    if table1 != nil {
        // Add headers
        table1.SetCellFormattedText(0, 0, "Name", func(r *elements.Run) {
            r.SetSpacing(2)
            r.SetBold(true)
        })
        table1.SetCellFormattedText(0, 1, "Age", func(r *elements.Run) {
            r.SetBold(true)
        })
        table1.SetCellFormattedText(0, 2, "City", func(r *elements.Run) {
            r.SetBold(true)
        })

        // Add data
        table1.SetCellText(1, 0, "John Doe")
        table1.SetCellText(1, 1, "30")
        table1.SetCellText(1, 2, "New York")

        table1.SetCellText(2, 0, "Jane Smith")
        table1.SetCellText(2, 1, "25")
        table1.SetCellText(2, 2, "London")

        // Set table properties
        table1.SetTableAlignment("center")
    }

    // Save the document
    if err := doc.Save("testdata/table_examples.docx"); err != nil {
        log.Fatalf("Failed to save document: %v", err)
    }
}
```

### Adding Images

```go
package main

import (
    "log"

    "github.com/didikprabowo/mbadocx"
    "github.com/didikprabowo/mbadocx/properties"
)

func main() {
    docx := mbadocx.New()
    docx.AddParagraph().AddText("Image with Borders and Effects")

    img1, err := docx.AddImage("mbadocx_logo.png")
    if err != nil {
        log.Fatal("failed add image")
    }

    img1.
        SetBorder(3, "#FF0000").
        SetAlignment(properties.AlignCenter)

        if err :=  docx.Save("testdata/image.docx"); err != nil {
        log.Fatalf("Failed to save document: %v", err)
    }
}
```


## Contributing

Contributions are welcome!  
Please see [CONTRIBUTING.md](./CONTRIBUTING.md) for guidelines.

---

## License

MIT License. See [LICENSE](./LICENSE) for details.