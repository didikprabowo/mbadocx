package writer

import (
	"encoding/xml"
	"io"
	"strings"
)

// StreamingXMLWriter provides efficient streaming XML writing capabilities
type StreamingXMLWriter struct {
	writer  io.Writer
	encoder *xml.Encoder
	buffer  *strings.Builder
}

// NewStreamingXMLWriter creates a new streaming XML writer
func NewStreamingXMLWriter(w io.Writer) *StreamingXMLWriter {
	return &StreamingXMLWriter{
		writer:  w,
		encoder: xml.NewEncoder(w),
		buffer:  &strings.Builder{},
	}
}

// WriteHeader writes the XML header
func (sw *StreamingXMLWriter) WriteHeader() error {
	_, err := sw.writer.Write([]byte(xml.Header))
	return err
}

// WriteStartElement writes a start element with attributes
func (sw *StreamingXMLWriter) WriteStartElement(name string, attrs ...xml.Attr) error {
	token := xml.StartElement{
		Name: xml.Name{Local: name},
		Attr: attrs,
	}
	return sw.encoder.EncodeToken(token)
}

// WriteEndElement writes an end element
func (sw *StreamingXMLWriter) WriteEndElement(name string) error {
	token := xml.EndElement{
		Name: xml.Name{Local: name},
	}
	return sw.encoder.EncodeToken(token)
}

// WriteCharData writes character data (text content)
func (sw *StreamingXMLWriter) WriteCharData(data string) error {
	token := xml.CharData(data)
	return sw.encoder.EncodeToken(token)
}

// WriteElement writes a complete element with text content
func (sw *StreamingXMLWriter) WriteElement(name, content string, attrs ...xml.Attr) error {
	if err := sw.WriteStartElement(name, attrs...); err != nil {
		return err
	}
	if content != "" {
		if err := sw.WriteCharData(content); err != nil {
			return err
		}
	}
	return sw.WriteEndElement(name)
}

// Flush flushes any buffered data to the underlying writer
func (sw *StreamingXMLWriter) Flush() error {
	return sw.encoder.Flush()
}

// Close finalizes the XML document and closes the writer
func (sw *StreamingXMLWriter) Close() error {
	return sw.encoder.Flush()
}

// WriteRaw writes raw XML data directly (use with caution)
func (sw *StreamingXMLWriter) WriteRaw(data []byte) error {
	// Flush encoder first to ensure proper ordering
	if err := sw.encoder.Flush(); err != nil {
		return err
	}
	_, err := sw.writer.Write(data)
	return err
}

// SetIndent sets the indentation for pretty printing
func (sw *StreamingXMLWriter) SetIndent(prefix, indent string) {
	sw.encoder.Indent(prefix, indent)
}