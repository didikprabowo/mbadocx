# Mbadocx

## Introduction

Introduction is library written in GO providing a set of functions that allow you to write to and read from Docx file.

## Features

- [x] Create `.docx` file from scratch
- [x] Add headings 1–6
- [x] Add paragraphs
- [x] Add hyperlinks
- [x] Generate document metadata
- [ ] Insert tables
- [ ] Add images
- [ ] Add headers and footers
- [ ] Custom styles
- [ ] Table of contents (TOC)


## Basic Usage

### Installation

```bash
go get github.com/didikprabowo/mbadocx
```

### Create document

```go
package main

import (
	"os"

	"github.com/didikprabowo/mbadocx"
	"github.com/didikprabowo/mbadocx/types"
)

func main() {
	// Initialize new document
	doc := mbadocx.NewDocument()

	// Add a Heading
	doc.AddHeading("My Document Title", types.Heading1)

	// Add a paragraph with mixed styles
	p := doc.AddParagraph()
	p.AddText("Hello, this is a ")
	p.AddBold("bold")
	p.AddText(" and ")
	p.AddItalic("italic")
	p.AddText(" example.")

	// Add a hyperlink
	p.AddHyperlink("Visit GitHub", "https://github.com/didikprabowo/mbadocx")

	// Add metadata
	doc.SetMetadata(types.Metadata{
		Author:  "Didik Prabowo",
		Company: "Example Inc.",
		Manager: "Tukang ketik",
	})

	// Save to file
	f, err := os.Create("example.docx")
	if err != nil {
		panic(err)
	}
	defer f.Close()

	writer := mbadocx.NewWriter(doc)
	if err := writer.Write(f); err != nil {
		panic(err)
	}
}
```


## Contributing
Contributions, feedback, and suggestions are very welcome! Feel free to open an issue or submit a PR.

## License
MIT License © 2025 Didik Prabowo