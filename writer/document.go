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
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"log"

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
	var buf bytes.Buffer
	indent := "  "
	buf.WriteString(xml.Header)

	// Write <w:document> manually
	buf.WriteString(` <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"`)
	buf.WriteString(` xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"`)
	buf.WriteString(` xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"`)
	buf.WriteString(` xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"`)
	buf.WriteString(` xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"`)
	buf.WriteString(` xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">`)

	// Open body
	// Write <w:body>
	buf.WriteString(indent + "<w:body>\n")

	for _, el := range d.document.Body().GetElements() {
		xmlData, err := el.XML()
		if err != nil {
			return nil, fmt.Errorf("serialize element: %w", err)
		}
		// Indent each line of xmlData
		lines := bytes.Split(xmlData, []byte("\n"))
		for _, line := range lines {
			if len(bytes.TrimSpace(line)) == 0 {
				continue
			}
			buf.WriteString(indent + indent)
			buf.Write(bytes.TrimRight(line, "\r\n"))
			buf.WriteString("\n")
		}
	}

	// Close body and document
	buf.WriteString(indent + "</w:body>\n")
	buf.WriteString("</w:document>\n")

	log.Printf("'%s' has been created.\n", d.Path())
	// log.Print(buf.String())

	return buf.Bytes(), nil
}

func (d *Document) WriteTo(w io.Writer) (int64, error) {
	data, err := d.Byte()
	if err != nil {
		return 0, err
	}
	n, err := w.Write(data)
	return int64(n), err
}
