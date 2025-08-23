package writer

import (
	"bytes"
	"io"
	"log"

	"github.com/didikprabowo/mbadocx/types"
)

type DocumentRelationship struct {
	types.Document
}

var _ zipWritable = (*DocumentRelationship)(nil)

// newDocument wraps types.Document in writer.Document
func newDocumentRelationship(doc types.Document) *DocumentRelationship {
	return &DocumentRelationship{doc}
}

// Path returns the path inside the .docx ZIP file
func (d *DocumentRelationship) Path() string {
	return "word/_rels/document.xml.rels"
}

// Byte generates the XML content for the document
func (d *DocumentRelationship) Byte() ([]byte, error) {
	var buf bytes.Buffer

	rels := d.Relationships()
	docXML, err := rels.DocumentXML()
	if err != nil {
		return nil, err
	}

	// Add XML declaration if missing
	if !bytes.HasPrefix(docXML, []byte("<?xml")) {
		buf.Write([]byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`))
		buf.WriteByte('\n')
	}

	// Append the actual document XML
	buf.Write(docXML)

	log.Printf("'%s' has been created.\n", d.Path())
	// log.Print(buf.String())

	return buf.Bytes(), nil
}

// WriteTo writes the document content to the given writer (implements io.WriterTo)
func (d *DocumentRelationship) WriteTo(w io.Writer) (int64, error) {
	content, err := d.Byte()
	if err != nil {
		return 0, err
	}

	n, err := w.Write(content)
	return int64(n), err
}
