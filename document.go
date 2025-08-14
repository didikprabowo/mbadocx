package mbadocx

import (
	"io"
	"os"
	"time"

	"github.com/didikprabowo/mbadocx/relationships"
	"github.com/didikprabowo/mbadocx/settings"
	"github.com/didikprabowo/mbadocx/styles"
	"github.com/didikprabowo/mbadocx/types"
	"github.com/didikprabowo/mbadocx/writer"
)

type Document struct {
	// Core components
	body          *Body
	relationships *relationships.Relationships
	settings      *settings.DocumentSettings
	style         *styles.DocumentStyles

	// Metadata
	metadata *types.Metadata
}

// New creates a new empty document
func New() *Document {
	return &Document{
		body:          &Body{Elements: make([]types.Element, 0)},
		relationships: relationships.NewDefault(),
		metadata:      NewDefaultMetadata(),
		settings:      settings.DefaultSettings(),
		style:         styles.DefaultDocumentStyles(),
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
func (d *Document) GetMetadata() *types.Metadata {
	return d.metadata
}

// Save writes the document to a file
func (d *Document) Save(filename string) error {
	file, err := os.Create(filename)
	if err != nil {
		return err
	}
	defer file.Close()

	return d.Write(file)
}

// SaveAs is an alias for Save
func (d *Document) SaveAs(filename string) error {
	return d.Save(filename)
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

// GetSettings
func (d *Document) GetSettings() *settings.DocumentSettings {
	if d.settings == nil {
		d.settings = settings.DefaultSettings()
	}
	return d.settings
}

// SetSettings
func (d *Document) SetSettings(docSettings *settings.DocumentSettings) {
	d.settings = docSettings
}
