package writer

import (
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
func (s *StylesWr) Byte() ([]byte, error) {
	buf := getBuffer()
	defer putBuffer(buf)

	buf.WriteString(xml.Header)

	enc := xml.NewEncoder(buf)
	enc.Indent("", "  ")

	styles := s.document.Styles().Get()
	if err := enc.Encode(styles); err != nil {
		return nil, fmt.Errorf("encoding Styles XML: %w", err)
	}

	log.Printf("'%s' has been created.\n", s.Path())
	// log.Print(buf.String())

	// Make a copy of the bytes before returning the buffer to the pool
	result := make([]byte, buf.Len())
	copy(result, buf.Bytes())
	return result, nil
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
