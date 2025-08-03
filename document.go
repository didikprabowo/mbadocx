package mbadocx

import (
	"io"
	"time"

	"github.com/didikprabowo/mbadocx/relationships"
	"github.com/didikprabowo/mbadocx/types"
	"github.com/didikprabowo/mbadocx/writer"
)

type Document struct {
	// Core components
	metadata      *types.Metadata
	body          *Body
	relationships *relationships.Relationships
}

// New creates a new empty document
func New() *Document {
	return &Document{
		body:          &Body{Elements: make([]types.Element, 0)},
		relationships: relationships.NewDefault(),
		metadata:      NewDefaultMetadata(),
	}
}

// NewDefaultMetadata creates default metadata
func NewDefaultMetadata() *types.Metadata {
	return &types.Metadata{
		Creator:  "Go DOCX Library",
		Created:  time.Now(),
		Modified: time.Now(),
		Revision: "1",
		Language: "en-US",
	}
}

// Metadata Methods

// SetMetadata sets complete document metadata
func (d *Document) SetMetadata(metadata *types.Metadata) {
	d.metadata = metadata
}

// GetMetadata returns the document metadata
func (d *Document) GetMetadata() types.Metadata {
	return *d.metadata
}

// Write writes the document to an io.Writer
func (d *Document) Write(w io.Writer) error {
	// Set modified at during write
	d.metadata.Modified = time.Now()

	docWriter := writer.NewWriter(d)
	// Write the document
	return docWriter.Write(w)
}

// Implement DocumentInterface methods

// GetBody
func (d *Document) GetBody() types.Body {
	return d.body
}

// GetRelationships
func (d *Document) GetRelationships() types.Relationships {
	return d.relationships
}
