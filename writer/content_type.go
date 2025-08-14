package writer

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"log"
)

var _ zipWritable = (*ContentTypes)(nil)

// ContentTypes represents the [Content_Types].xml part in a DOCX package.
type ContentTypes struct {
	XMLName   xml.Name   `xml:"Types"`
	Xmlns     string     `xml:"xmlns,attr"`
	Defaults  []Default  `xml:"Default"`
	Overrides []Override `xml:"Override"`
}

// Default defines a default MIME type for a file extension.
type Default struct {
	Extension   string `xml:"Extension,attr"`
	ContentType string `xml:"ContentType,attr"`
}

// Override defines a specific MIME type for a file path.
type Override struct {
	PartName    string `xml:"PartName,attr"`
	ContentType string `xml:"ContentType,attr"`
}

// newContentType initializes and returns a default ContentTypes definition.
func newContentType() *ContentTypes {
	return &ContentTypes{
		Xmlns: "http://schemas.openxmlformats.org/package/2006/content-types",
		Defaults: []Default{
			{Extension: "rels", ContentType: "application/vnd.openxmlformats-package.relationships+xml"},
			{Extension: "xml", ContentType: "application/xml"},
			{Extension: "png", ContentType: "image/png"},
			{Extension: "jpeg", ContentType: "image/jpeg"},
			{Extension: "jpg", ContentType: "image/jpeg"},
			{Extension: "gif", ContentType: "image/gif"},
			{Extension: "bmp", ContentType: "image/bmp"},
			{Extension: "tiff", ContentType: "image/tiff"},
			{Extension: "tif", ContentType: "image/tiff"},
		},
		Overrides: []Override{
			{PartName: "/word/document.xml", ContentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"},
			{PartName: "/word/numbering.xml", ContentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"},
			{PartName: "/word/styles.xml", ContentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"},
			{PartName: "/word/settings.xml", ContentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"},
			{PartName: "/word/webSettings.xml", ContentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml"},
			{PartName: "/word/fontTable.xml", ContentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"},
			{PartName: "/word/theme/theme1.xml", ContentType: "application/vnd.openxmlformats-officedocument.theme+xml"},
			{PartName: "/docProps/core.xml", ContentType: "application/vnd.openxmlformats-package.core-properties+xml"},
			{PartName: "/docProps/app.xml", ContentType: "application/vnd.openxmlformats-officedocument.extended-properties+xml"},
		},
	}
}

// Path returns the location of the content types file within the DOCX ZIP.
func (ct *ContentTypes) Path() string {
	return "[Content_Types].xml"
}

// Byte serializes the ContentTypes struct to XML with an XML declaration.
func (ct *ContentTypes) Byte() ([]byte, error) {
	var buf bytes.Buffer

	// Write XML declaration
	buf.WriteString(xml.Header)

	// Encode the struct
	enc := xml.NewEncoder(&buf)
	enc.Indent("", "  ")

	if err := enc.Encode(ct); err != nil {
		return nil, fmt.Errorf("encoding ContentTypes XML: %w", err)
	}

	log.Printf("'%s' has been created.\n", ct.Path())
	// log.Print(buf.String())

	return buf.Bytes(), nil
}

// WriteTo writes the XML to an io.Writer (implements io.WriterTo).
func (ct *ContentTypes) WriteTo(w io.Writer) (int64, error) {
	xmlData, err := ct.Byte()
	if err != nil {
		return 0, err
	}

	n, err := w.Write(xmlData)
	return int64(n), err
}
