package mbadocx

import (
	"fmt"
	"io"
	"os"
	"sync"
	"time"

	contenttypes "github.com/didikprabowo/mbadocx/content_types"
	"github.com/didikprabowo/mbadocx/metadata"
	"github.com/didikprabowo/mbadocx/relationships"
	"github.com/didikprabowo/mbadocx/styles"
	"github.com/didikprabowo/mbadocx/types"
	"github.com/didikprabowo/mbadocx/writer"
)

type Document struct {
	// Core components
	contentTypes  *contenttypes.ContentTypes
	body          *Body
	relationships *relationships.Relationships
	styles        *styles.Styles

	// Metadata
	metadata *metadata.Metadata

	// Internal state
	mu     sync.RWMutex
	closed bool

	// Resources that need cleanup
	openFiles []*os.File
}

// New creates a new empty document
func New() *Document {
	return &Document{
		body:          &Body{Elements: make([]types.Element, 0)},
		relationships: relationships.NewDefault(),
		contentTypes:  contenttypes.NewDefaultContentType(),
		metadata:      metadata.NewDefaultMetadata(),
		styles:        styles.NewDefaultStyles(),
		openFiles:     make([]*os.File, 0),
		closed:        false,
	}
}

// Save writes the document to a file
func (d *Document) Save(filename string) error {
	d.mu.Lock()
	defer d.mu.Unlock()

	if d.closed {
		return fmt.Errorf("document has been closed")
	}

	file, err := os.Create(filename)
	if err != nil {
		return fmt.Errorf("failed to create file: %w", err)
	}

	// Track the file for cleanup
	d.openFiles = append(d.openFiles, file)

	// Ensure file is closed even if Write fails
	defer func() {
		if closeErr := file.Close(); closeErr != nil && err == nil {
			err = closeErr
		}
		// Remove from openFiles after closing
		for i, f := range d.openFiles {
			if f == file {
				d.openFiles = append(d.openFiles[:i], d.openFiles[i+1:]...)
				break
			}
		}
	}()

	if err := d.write(file); err != nil {
		return fmt.Errorf("failed to write document: %w", err)
	}

	return nil
}

// SaveAs is an alias for Save
func (d *Document) SaveAs(filename string) error {
	return d.Save(filename)
}

// Write writes the document to an io.Writer
func (d *Document) Write(w io.Writer) error {
	d.mu.Lock()
	defer d.mu.Unlock()

	if d.closed {
		return fmt.Errorf("document has been closed")
	}

	return d.write(w)
}

// write is the internal write method (must be called with lock held)
func (d *Document) write(w io.Writer) error {
	// Set modified time during write
	d.metadata.Modified = time.Now()

	docWriter := writer.NewWriter(d)

	// Write the document
	if err := docWriter.Write(w); err != nil {
		return err
	}

	return nil
}

// Close releases all resources associated with the document
func (d *Document) Close() error {
	d.mu.Lock()
	defer d.mu.Unlock()

	if d.closed {
		return nil // Already closed, no-op
	}

	var errs []error

	// Close any open files
	for _, file := range d.openFiles {
		if file != nil {
			if err := file.Close(); err != nil {
				errs = append(errs, fmt.Errorf("failed to close file %s: %w", file.Name(), err))
			}
		}
	}
	d.openFiles = nil

	// Clear references to allow garbage collection
	d.body = nil
	d.relationships = nil
	d.contentTypes = nil
	d.metadata = nil
	d.styles = nil

	d.closed = true

	// Combine errors if any occurred
	if len(errs) > 0 {
		return fmt.Errorf("errors during close: %v", errs)
	}

	return nil
}

// IsClosed returns whether the document has been closed
func (d *Document) IsClosed() bool {
	d.mu.RLock()
	defer d.mu.RUnlock()
	return d.closed
}

// GetMetadata returns the document metadata
func (d *Document) Metadata() types.Metadata {
	if d.closed {
		return nil
	}
	return d.metadata
}

// GetBody returns the document body
func (d *Document) GetBody() types.Body {
	if d.closed {
		return nil
	}
	return d.body
}

// GetRelationships returns the document relationships
func (d *Document) GetRelationships() types.Relationships {
	if d.closed {
		return nil
	}
	return d.relationships
}

// GetStyles returns the document styles
func (d *Document) GetStyles() types.Styles {
	if d.closed {
		return nil
	}
	return d.styles
}

// Styles returns the document styles (alias for GetStyles for consistency)
func (d *Document) Styles() types.Styles {
	return d.GetStyles()
}

// ContentTypes returns the document content types
func (d *Document) ContentTypes() types.ContentTypes {
	if d.closed {
		return nil
	}
	return d.contentTypes
}
