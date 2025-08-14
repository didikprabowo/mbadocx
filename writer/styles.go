package writer

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"log"

	"github.com/didikprabowo/mbadocx/types"
)

var _ zipWritable = (*StylesWr)(nil)

type StylesWr struct {
	// document
	document types.Document
}

// newStylesWR
func newStylesWr(document types.Document) *StylesWr {
	return &StylesWr{document: document}
}

// Path
func (swr *StylesWr) Path() string {
	return "word/styles.xml"
}

// Byte
func (swr *StylesWr) Byte() ([]byte, error) {
	var buf bytes.Buffer

	// Write XML declaration
	buf.WriteString(xml.Header)

	// Encode the struct
	enc := xml.NewEncoder(&buf)
	enc.Indent("", "  ")

	styles := swr.document.Styles().Get()
	if err := enc.Encode(styles); err != nil {
		return nil, fmt.Errorf("encoding ContentTypes XML: %w", err)
	}

	log.Printf("'%s' has been created.\n", swr.Path())
	// log.Print(buf.String())

	return buf.Bytes(), nil
}

// WriteTo
func (swr *StylesWr) WriteTo(w io.Writer) (int64, error) {
	xmlData, err := swr.Byte()
	if err != nil {
		return 0, err
	}

	n, err := w.Write(xmlData)
	return int64(n), err
}
