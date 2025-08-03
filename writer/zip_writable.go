package writer

import "io"

// ZipWritable can be written into a ZIP file at a specific path.
type zipWritable interface {
	Byte() ([]byte, error)
	Path() string
	io.WriterTo
}
