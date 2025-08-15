<div align="center">
  <img src="./mbadocx.svg" alt="Mbadocx - Go DOCX Library" width="250">
  
  
  **Go library for creating, reading and manipulating DOCX files**
  
  [![Go Reference](https://pkg.go.dev/badge/github.com/yourusername/mbadocx.svg)](https://pkg.go.dev/github.com/yourusername/mbadocx)
  [![Go Report Card](https://goreportcard.com/badge/github.com/yourusername/mbadocx)](https://goreportcard.com/report/github.com/yourusername/mbadocx)
  [![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
  
</div>

# Mbadocx 

**Mbadocx** is a Go library for programmatically creating and manipulating Microsoft Word (DOCX) documents.  
It provides a modular, extensible API for generating Word documents with advanced formatting, metadata, and resource management.

## Features

The feature able to see on [features](./docs/features_documents.md).


## Installation

```sh
go get github.com/didikprabowo/mbadocx
```

## Usage Example

```go
package main

import (
	"github.com/didikprabowo/mbadocx"
)

func main() {
	// Create a new document
	doc := mbadocx.New()

	// Add a paragraph with formatted text
	para := doc.AddParagraph()
	para.AddText("Hello, mbadocx!").SetBold(true).SetColor("#2E86C1").SetFontSize(16)

	// Add a hyperlink
	para.AddHyperlink("Visit GitHub", "https://github.com/didikprabowo/mbadocx")

	// Add a line break and another paragraph
	para.AddLineBreak()
	doc.AddParagraph().AddText("This is a second paragraph.")

	// Save the document
	if err := doc.Save("testdata/basic.docx"); err != nil {
		panic(err)
	}
}
```

## Contributing

Contributions are welcome!  
Please see [CONTRIBUTING.md](./CONTRIBUTING.md) for guidelines.

---

## License

MIT License. See [LICENSE](./LICENSE) for details.