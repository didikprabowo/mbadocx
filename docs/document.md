#  Document Plan: Build .docx Generator in Go

## 🎯 Objective

Develop a modular Go library to generate .docx (WordprocessingML) files using standard Go features (encoding/xml, archive/zip, etc.), enabling:

- Paragraphs, runs, and styled text (bold, italic, underline, links)
- Headings and structured styles
- Page settings and margins
- Tables
- Lists (bullets, numbering)
- Media (images, optional phase)
- File relationships (.rels)

## 📁 .docx Structure Overview
```
.docx
├── [Content_Types].xml
├── _rels/
│   └── .rels
├── word/
│   ├── document.xml
│   ├── styles.xml
│   ├── settings.xml
│   ├── fontTable.xml
│   ├── theme/
│   │   └── theme1.xml
│   └── _rels/
│       └── document.xml.rels
└── docProps/
    ├── core.xml
    └── app.xml
```

## Feature Specifications

### Document Elements
**Text Features**
- Plain text
- Formatted text (bold, italic, underline, strikethrough)
- Heading 1-6.
- Font customization (family, size, color)
- Text alignment (left, center, right, justify)
- Line spacing and paragraph spacing
- Bullet points and numbered lists
- Hyperlinks

**Structural Elements**

- Headers and footers
- Page numbers
- Section breaks
- Page breaks
- Table of contents
- Bookmarks and cross-references
## 🔧 Phase-by-Phase Implementation Plan

### ✅ Phase 1: Core Writer & File Structure
 - [ ] Create Writer to wrap zip.Writer
 - [ ] Implement zipWritable interface:

 ```go
 type zipWritable interface {
    io.WriterTo
    Path() string
}
 ```
 - [ ] Write are component
    - [x] `[Content_Types].xml`
    - [x] `_rels/.rels`
    - [x] `word/document.xml`

### ✍️ Phase 2: Text and Paragraph

- [x] Define Paragraph, Run, Text structs
- [ ] Add support for: 
    - Plain text
    - Bold / Italic / Underline
    - Hyperlinks
    - Headings
    - Hyperlinks
    - Headings 1-6