package writer

import (
	"encoding/xml"
	"fmt"
	"io"
	"log"

	"github.com/didikprabowo/mbadocx/types"
)

type DocumentRelationship struct {
	document types.Document
}

func newDocumentRelationship(document types.Document) *DocumentRelationship {
	return &DocumentRelationship{document: document}
}

func (dr *DocumentRelationship) Path() string {
	return "word/_rels/document.xml.rels"
}

// Byte generates the XML content for the document
func (dr *DocumentRelationship) Byte() ([]byte, error) {
	buf := getBuffer()
	defer putBuffer(buf)

	buf.WriteString(xml.Header)

	// Get relationships and generate document relationships XML
	rels := dr.document.Relationships()
	docXML, err := rels.DocumentXML()
	if err != nil {
		return nil, fmt.Errorf("generating document relationships XML: %w", err)
	}

	// Write the raw XML data
	buf.Write(docXML)

	log.Printf("'%s' has been created.\n", dr.Path())
	// log.Print(buf.String())

	// Make a copy of the bytes before returning the buffer to the pool
	result := make([]byte, buf.Len())
	copy(result, buf.Bytes())
	return result, nil
}

var _ zipWritable = (*DocumentRelationship)(nil)

// WriteTo writes the document content to the given writer (implements io.WriterTo)
func (dr *DocumentRelationship) WriteTo(w io.Writer) (int64, error) {
	content, err := dr.Byte()
	if err != nil {
		return 0, err
	}

	n, err := w.Write(content)
	return int64(n), err
}
