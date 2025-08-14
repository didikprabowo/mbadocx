package writer

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"log"

	"github.com/didikprabowo/mbadocx/types"
)

var _ zipWritable = (*ContentTypesWr)(nil)

// ContentTypes represents the [Content_Types].xml part in a DOCX package.
type ContentTypesWr struct {
	document types.Document
}

// newContentType initializes and returns a default ContentTypes definition.
func newContentTypeWr(document types.Document) *ContentTypesWr {
	return &ContentTypesWr{document: document}
}

// Path returns the location of the content types file within the DOCX ZIP.
func (ct *ContentTypesWr) Path() string {
	return "[Content_Types].xml"
}

// Byte serializes the ContentTypes struct to XML with an XML declaration.
func (ct *ContentTypesWr) Byte() ([]byte, error) {
	var buf bytes.Buffer

	// Write XML declaration
	buf.WriteString(xml.Header)

	// Encode the struct
	enc := xml.NewEncoder(&buf)
	enc.Indent("", "  ")

	contentTypes := ct.document.ContentTypes().Get()
	if err := enc.Encode(contentTypes); err != nil {
		return nil, fmt.Errorf("encoding ContentTypes XML: %w", err)
	}

	log.Printf("'%s' has been created.\n", ct.Path())
	// log.Print(buf.String())

	return buf.Bytes(), nil
}

// WriteTo writes the XML to an io.Writer (implements io.WriterTo).
func (ct *ContentTypesWr) WriteTo(w io.Writer) (int64, error) {
	xmlData, err := ct.Byte()
	if err != nil {
		return 0, err
	}

	n, err := w.Write(xmlData)
	return int64(n), err
}
