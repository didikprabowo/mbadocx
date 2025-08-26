package writer

// EXAMPLE

// <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
// <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
//             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
//   <w:body>
//     <w:p>
//       <w:r>
//         <w:rPr>
//           <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/>
//           <w:sz w:val="22"/>
//           <w:szCs w:val="22"/>
//         </w:rPr>
//         <w:t>this is example paragraph</w:t>
//       </w:r>
//       <w:hyperlink r:id="rId1" w:history="1">
//         <w:r>
//           <w:rPr>
//             <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/>
//             <w:sz w:val="22"/>
//             <w:szCs w:val="22"/>
//             <w:u w:val="single"/>
//             <w:color w:val="0563C1"/>
//           </w:rPr>
//           <w:t>go to github</w:t>
//         </w:r>
//       </w:hyperlink>
//     </w:p>
//     <w:sectPr>
//       <w:pgSz w:w="11900" w:h="16840"/>
//       <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
//     </w:sectPr>
//   </w:body>
// </w:document>

import (
	"encoding/xml"
	"fmt"
	"io"
	"log"
	"strings"

	"github.com/didikprabowo/mbadocx/types"
)

var _ zipWritable = (*Document)(nil)

type Document struct {
	document types.Document
}

func newDocument(doc types.Document) *Document {
	return &Document{document: doc}
}

func (d *Document) Path() string {
	return "word/document.xml"
}

func (d *Document) Byte() ([]byte, error) {
	var builder strings.Builder
	// Pre-allocate a reasonable capacity to reduce reallocations
	builder.Grow(8192)

	const indent = "  "
	const doubleIndent = "    "

	builder.WriteString(xml.Header)

	// Write <w:document> manually
	builder.WriteString(` <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"`)
	builder.WriteString(` xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"`)
	builder.WriteString(` xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"`)
	builder.WriteString(` xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"`)
	builder.WriteString(` xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"`)
	builder.WriteString(` xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">`)

	// Open body
	// Write <w:body>
	builder.WriteString(indent)
	builder.WriteString("<w:body>\n")

	for _, el := range d.document.Body().GetElements() {
		xmlData, err := el.XML()
		if err != nil {
			return nil, fmt.Errorf("serialize element: %w", err)
		}
		
		// Optimize indentation by processing the XML data more efficiently
		d.indentXMLData(&builder, xmlData, doubleIndent)
	}

	// Close body and document
	builder.WriteString(indent)
	builder.WriteString("</w:body>\n")
	builder.WriteString("</w:document>\n")

	log.Printf("'%s' has been created.\n", d.Path())
	// log.Print(builder.String())

	return []byte(builder.String()), nil
}

// indentXMLData efficiently indents XML data by processing it line by line
func (d *Document) indentXMLData(builder *strings.Builder, xmlData []byte, indent string) {
	// Convert to string once to avoid repeated conversions
	xmlStr := string(xmlData)
	lines := strings.Split(xmlStr, "\n")
	
	for _, line := range lines {
		trimmed := strings.TrimSpace(line)
		if len(trimmed) == 0 {
			continue
		}
		builder.WriteString(indent)
		builder.WriteString(trimmed)
		builder.WriteString("\n")
	}
}

func (d *Document) WriteTo(w io.Writer) (int64, error) {
	data, err := d.Byte()
	if err != nil {
		return 0, err
	}
	n, err := w.Write(data)
	return int64(n), err
}
