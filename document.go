// Package mbadocx provides a pure-Go DOCX generator and writer.
//
// It is designed to create Microsoft Word .docx documents without external
// dependencies, using only the Go standard library and custom struct models.
//
// # Example
//
//	Create a new DOCX file:
//
//		package main
//
//		import (
//			"log"
//
//			"github.com/didikprabowo/mbadocx"
//		)
//
//		func main() {
//			doc := mbadocx.New()
//			doc.AddParagraph("Hello, World!")
//
//			if err := doc.Save("example.docx"); err != nil {
//				log.Fatal(err)
//			}
//		}
//
// This package is modular: subpackages such as `writer`, `elements`, and
// `types` handle XML serialization, building elements, and struct models.
//
// For detailed usage and advanced composition, see the subpackage docs.
//

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

// Document represents a DOCX document and its core components.
type Document struct {
	// Core components
	contentTypes  *contenttypes.ContentTypes   // Content types for the DOCX package
	body          *Body                        // Main document body
	relationships *relationships.Relationships // Relationships (e.g., images, styles)
	styles        *styles.Styles               // Document styles

	// Metadata
	metadata *metadata.Metadata // Document metadata (author, timestamps, etc.)
	media    *Media

	// Internal state
	mu     sync.RWMutex // Mutex for thread safety
	closed bool         // Indicates if the document is closed

	// Resources that need cleanup
	openFiles []*os.File // List of open files for cleanup
}

// New creates a new empty document with default components.
func New() *Document {
	docx := &Document{
		body:          NewBody(),
		relationships: relationships.NewDefault(),
		contentTypes:  contenttypes.NewDefaultContentType(),
		metadata:      metadata.NewDefaultMetadata(),
		styles:        styles.NewDefaultStyles(),
		openFiles:     make([]*os.File, 0),
		media:         &Media{},
		closed:        false,
	}

	return docx
}

// Save writes the document to a file with the given filename.
func (d *Document) Save(filename string) (err error) {
	d.mu.Lock()
	defer d.mu.Unlock()

	if d.closed {
		return fmt.Errorf("document has been closed")
	}

	file, err := os.Create(filename)
	if err != nil {
		return fmt.Errorf("failed to create file: %w", err)
	}

	d.openFiles = append(d.openFiles, file)

	defer func() {
		closeErr := file.Close()
		if closeErr != nil && err == nil {
			err = closeErr
		}
		for i, f := range d.openFiles {
			if f == file {
				d.openFiles = append(d.openFiles[:i], d.openFiles[i+1:]...)
				break
			}
		}
	}()

	if writeErr := d.write(file); writeErr != nil {
		err = fmt.Errorf("failed to write document: %w", writeErr)
	}

	return
}

// SaveAs is an alias for Save, writes the document to a new file.
func (d *Document) SaveAs(filename string) error {
	return d.Save(filename)
}

// Write writes the document to an io.Writer (e.g., file, buffer).
func (d *Document) Write(w io.Writer) error {
	d.mu.Lock()
	defer d.mu.Unlock()

	if d.closed {
		return fmt.Errorf("document has been closed")
	}

	return d.write(w)
}

// write is the internal write method (must be called with lock held).
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

// Close releases all resources associated with the document.
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

// IsClosed returns whether the document has been closed.
func (d *Document) IsClosed() bool {
	d.mu.RLock()
	defer d.mu.RUnlock()
	return d.closed
}

// Metadata returns the document metadata.
func (d *Document) Metadata() types.Metadata {
	if d.closed {
		return nil
	}
	return d.metadata
}

// Body returns the document body.
func (d *Document) Body() types.Body {
	if d.closed {
		return nil
	}
	return d.body
}

// Relationships returns the document relationships.
func (d *Document) Relationships() types.Relationships {
	if d.closed {
		return nil
	}
	return d.relationships
}

// Styles returns the document styles (alias for GetStyles for consistency).
func (d *Document) Styles() types.Styles {
	if d.closed {
		return nil
	}
	return d.styles
}

// ContentTypes returns the document content types.
func (d *Document) ContentTypes() types.ContentTypes {
	if d.closed {
		return nil
	}
	return d.contentTypes
}

// Media
func (d *Document) Media() []types.Media {
	return d.media.Media
}
