package writer

import (
	"bytes"
	"fmt"
	"io"
	"log"

	"github.com/didikprabowo/mbadocx/types"
)

var _ zipWritable = (*Relationships)(nil)

// Relationships represents the _rels/.rels file in a DOCX package.
type Relationships struct {
	types.Document
}

// newRelationships wraps types.Document in a writer.Relationships
func newRelationships(doc types.Document) *Relationships {
	return &Relationships{doc}
}

// Path returns the location of the part inside the DOCX ZIP.
func (r *Relationships) Path() string {
	return "_rels/.rels"
}

// Byte returns the full XML content for the _rels/.rels part.
func (r *Relationships) Byte() ([]byte, error) {
	rels := r.GetRelationships()

	relsXML, err := rels.PackageXML()
	if err != nil {
		return nil, fmt.Errorf("generating package relationships XML: %w", err)
	}

	var buf bytes.Buffer

	// Add XML declaration
	buf.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n")

	// Append marshaled <Relationships> content
	buf.Write(relsXML)

	log.Printf("'%s' has been created.\n", r.Path())
	// log.Print(buf.String())

	return buf.Bytes(), nil
}

// WriteTo writes the XML content to the given writer.
func (r *Relationships) WriteTo(w io.Writer) (int64, error) {
	data, err := r.Byte()
	if err != nil {
		return 0, err
	}

	n, err := w.Write(data)
	return int64(n), err
}
