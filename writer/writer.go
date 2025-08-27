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
}

// DefaultSaveOptions returns default save options
func DefaultSaveOptions() SaveOptions {
	return SaveOptions{
		CompressionLevel:        -1,
		PrettyPrint:             false,
		IncludeCustomProperties: true,
		UpdateFields:            true,
		Metadata:                make(map[string]string),
	}
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

func (w *Writer) writeFile(name string, data []byte) error {
	writer, _ := w.zipWriter.Create(name)
	_, err := writer.Write(data)
	return err
}

// Write writes the document to an io.Writer
func (w *Writer) Write(writer io.Writer) error {
	w.zipWriter = zip.NewWriter(writer)
	defer func() {
		_ = w.zipWriter.Close()
	}()

	// Set compression level if specified
	if w.options.CompressionLevel >= 0 {
		w.zipWriter.RegisterCompressor(zip.Deflate, func(out io.Writer) (io.WriteCloser, error) {
			return flate.NewWriter(out, w.options.CompressionLevel)
		})
	}

	var components []zipWritable

	components = append(components,
		newContentTypeWr(w.document),        // [Content_Types].xml
		newRelationships(w.document),        // _rels/.rels
		newDocumentRelationship(w.document), // word/_rels/document.xml.rels
		newDocument(w.document),             // word/document.xml
		newCoreProperties(w.document),       // docProps/core.xml
		newAppProperties(w.document),        // docProps/app.xml
		newNumberingDefinitions(),           // word/numbering.xml
		newStylesWr(w.document),
		// Add others like styles, header/footer, etc.
	)

	// Write each component
	for _, part := range components {
		if err := w.writeToZip(part); err != nil {
			return fmt.Errorf("write %s: %w", part.Path(), err)
		}
	}

	// Write file
	// word/media/*
	for _, media := range w.document.Media() {
		err := w.writeFile(media.TargetPath()+media.FileName(), media.RawContent())
		if err != nil {
			log.Print(err.Error())
		}
		log.Printf("'%s' has been created.\n", media.TargetPath()+media.FileName())
	}

	return nil
}

// formatDateTime formats a time for Office Open XML
func formatDateTime(t time.Time) string {
	if t.IsZero() {
		return time.Now().Format(time.RFC3339)
	}
	return t.Format(time.RFC3339)
}
