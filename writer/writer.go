// document.docx
// ├── [Content_Types].xml
// ├── _rels/
// │   └── .rels
// ├── word/
// │   ├── document.xml
// │   ├── styles.xml
// │   ├── settings.xml
// │   ├── fontTable.xml
// │   ├── theme/
// │   │   └── theme1.xml
// │   └── _rels/
// │       └── document.xml.rels
// └── docProps/
//     ├── core.xml
//     └── app.xml

package writer

import (
	"archive/zip"
	"compress/flate"
	"fmt"
	"io"
	"log"
	"time"

	"github.com/didikprabowo/mbadocx/types"
)

// Writer handles writing documents to DOCX format
type Writer struct {
	document   types.Document // Uses interface instead of concrete type
	zipWriter  *zip.Writer
	mediaFiles map[string][]byte
	options    SaveOptions
}

// SaveOptions provides options for saving documents
type SaveOptions struct {
	CompressionLevel        int
	PrettyPrint             bool
	IncludeCustomProperties bool
	UpdateFields            bool
	Metadata                map[string]string
	// New performance-oriented options
	BufferSize              int  // Buffer size for ZIP operations
	ConcurrentComponents    bool // Enable concurrent component processing
	MinimalNumbering        bool // Use minimal numbering when possible
}

// DefaultSaveOptions returns default save options optimized for performance
func DefaultSaveOptions() SaveOptions {
	return SaveOptions{
		CompressionLevel:        flate.DefaultCompression, // Better than -1 for most cases
		PrettyPrint:             false,
		IncludeCustomProperties: true,
		UpdateFields:            true,
		Metadata:                make(map[string]string),
		BufferSize:              32 * 1024, // 32KB buffer for better I/O performance
		ConcurrentComponents:    true,      // Enable concurrent processing by default
		MinimalNumbering:        true,      // Use minimal numbering when possible
	}
}

// FastSaveOptions returns save options optimized for speed over file size
func FastSaveOptions() SaveOptions {
	opts := DefaultSaveOptions()
	opts.CompressionLevel = flate.NoCompression // Fastest compression
	opts.PrettyPrint = false                    // No pretty printing for speed
	return opts
}

// CompactSaveOptions returns save options optimized for smaller file size
func CompactSaveOptions() SaveOptions {
	opts := DefaultSaveOptions()
	opts.CompressionLevel = flate.BestCompression // Maximum compression
	opts.MinimalNumbering = true                  // Minimal numbering for smaller files
	return opts
}

// NewWriter creates a new writer for the document
func NewWriter(doc types.Document) *Writer {
	return &Writer{
		document:   doc,
		mediaFiles: make(map[string][]byte),
		options:    DefaultSaveOptions(),
	}
}

// WriteToZip writes any ZipWritable part to a zip.Writer
func (w *Writer) writeToZip(part zipWritable) error {
	writer, err := w.zipWriter.Create(part.Path())
	if err != nil {
		return err
	}
	_, err = part.WriteTo(writer)
	return err
}

func (w *Writer) writeFile(name string, data []byte) {
	writer, _ := w.zipWriter.Create(name)
	writer.Write(data)
}

// Write writes the document to an io.Writer
func (w *Writer) Write(writer io.Writer) error {
	w.zipWriter = zip.NewWriter(writer)
	defer func() {
		if err := w.zipWriter.Close(); err != nil {
			log.Printf("Error closing zip writer: %v", err)
		}
	}()

	// Set compression level if specified
	if w.options.CompressionLevel >= 0 {
		w.zipWriter.RegisterCompressor(zip.Deflate, func(out io.Writer) (io.WriteCloser, error) {
			return flate.NewWriter(out, w.options.CompressionLevel)
		})
	}

	// Create numbering definitions based on options
	var numberingComponent zipWritable
	if w.options.MinimalNumbering {
		numberingComponent = newNumberingDefinitionsForDocument(w.document)
	} else {
		numberingComponent = newNumberingDefinitions()
	}

	// Create components
	components := []zipWritable{
		newContentTypeWr(w.document),        // [Content_Types].xml
		newRelationships(w.document),        // _rels/.rels
		newDocumentRelationship(w.document), // word/_rels/document.xml.rels
		newDocument(w.document),             // word/document.xml
		newCoreProperties(w.document),       // docProps/core.xml
		newAppProperties(w.document),        // docProps/app.xml
		numberingComponent,                  // word/numbering.xml
		newStylesWr(w.document),            // word/styles.xml
	}

	// Process components based on concurrency setting
	if w.options.ConcurrentComponents {
		return w.writeConcurrently(components)
	} else {
		return w.writeSequentially(components)
	}
}

// writeConcurrently processes components concurrently for better performance
func (w *Writer) writeConcurrently(components []zipWritable) error {
	type componentResult struct {
		component zipWritable
		err       error
	}

	// Create a channel for results
	resultChan := make(chan componentResult, len(components))
	
	// Process components concurrently
	for _, component := range components {
		go func(comp zipWritable) {
			err := w.writeToZip(comp)
			resultChan <- componentResult{component: comp, err: err}
		}(component)
	}

	// Collect results and check for errors
	for i := 0; i < len(components); i++ {
		result := <-resultChan
		if result.err != nil {
			return fmt.Errorf("write %s: %w", result.component.Path(), result.err)
		}
	}

	return w.writeMediaFiles()
}

// writeSequentially processes components sequentially (safer for some cases)
func (w *Writer) writeSequentially(components []zipWritable) error {
	// Write each component
	for _, component := range components {
		if err := w.writeToZip(component); err != nil {
			return fmt.Errorf("write %s: %w", component.Path(), err)
		}
	}

	return w.writeMediaFiles()
}

// writeMediaFiles writes all media files with proper error handling
func (w *Writer) writeMediaFiles() error {
	for _, media := range w.document.Media() {
		if err := w.writeMediaFile(media); err != nil {
			return fmt.Errorf("write media file %s: %w", media.FileName(), err)
		}
	}
	return nil
}

// writeMediaFile writes a single media file with error handling
func (w *Writer) writeMediaFile(media types.Media) error {
	fileName := media.TargetPath() + media.FileName()
	writer, err := w.zipWriter.Create(fileName)
	if err != nil {
		return fmt.Errorf("create media file %s: %w", fileName, err)
	}
	
	content := media.RawContent()
	if _, err := writer.Write(content); err != nil {
		return fmt.Errorf("write media content for %s: %w", fileName, err)
	}
	
	log.Printf("'%s' has been created.\n", fileName)
	return nil
}

// formatDateTime formats a time for Office Open XML
func formatDateTime(t time.Time) string {
	if t.IsZero() {
		return time.Now().Format(time.RFC3339)
	}
	return t.Format(time.RFC3339)
}
