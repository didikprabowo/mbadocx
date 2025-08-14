// Example metadata XML output:
/*
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
	<dc:title>My Document Title</dc:title>
	<dc:subject>Document Subject</dc:subject>
	<dc:creator>John Doe</dc:creator>
	<cp:keywords>docx, go, example</cp:keywords>
	<dc:description>This is a sample document</dc:description>
	<cp:lastModifiedBy>Jane Smith</cp:lastModifiedBy>
	<cp:revision>2</cp:revision>
	<dcterms:created xsi:type="dcterms:W3CDTF">2024-01-15T10:30:00Z</dcterms:created>
	<dcterms:modified xsi:type="dcterms:W3CDTF">2024-01-15T14:45:00Z</dcterms:modified>
	<cp:category>Reports</cp:category>
	<cp:contentStatus>Final</cp:contentStatus>
	<dc:language>en-US</dc:language>
	<cp:version>1.0</cp:version>
</cp:coreProperties>
*/

package writer

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"log"

	"github.com/didikprabowo/mbadocx/types"
)

var _ zipWritable = (*CoreProperties)(nil)

type CoreProperties struct {
	*CorePropertiesXML
	// Main document
	document types.Document
}

type CorePropertiesXML struct {
	XMLName       xml.Name `xml:"cp:coreProperties"`
	XmlnsCp       string   `xml:"xmlns:cp,attr"`
	XmlnsDc       string   `xml:"xmlns:dc,attr"`
	XmlnsDcterms  string   `xml:"xmlns:dcterms,attr"`
	XmlnsDcmitype string   `xml:"xmlns:dcmitype,attr"`
	XmlnsXsi      string   `xml:"xmlns:xsi,attr"`

	Title          string       `xml:"dc:title,omitempty"`
	Subject        string       `xml:"dc:subject,omitempty"`
	Creator        string       `xml:"dc:creator,omitempty"`
	Keywords       string       `xml:"cp:keywords,omitempty"`
	Description    string       `xml:"dc:description,omitempty"`
	LastModifiedBy string       `xml:"cp:lastModifiedBy,omitempty"`
	Revision       string       `xml:"cp:revision,omitempty"`
	Created        *TimeElement `xml:"dcterms:created"`
	Modified       *TimeElement `xml:"dcterms:modified"`
	Category       string       `xml:"cp:category,omitempty"`
	ContentStatus  string       `xml:"cp:contentStatus,omitempty"`
	Language       string       `xml:"dc:language,omitempty"`
	Version        string       `xml:"cp:version,omitempty"`
}

// TimeElement represents a time element with xsi:type
type TimeElement struct {
	XMLName xml.Name `xml:""`
	Type    string   `xml:"xsi:type,attr"`
	Value   string   `xml:",chardata"`
}

var _ zipWritable = (*DocumentRelationship)(nil)

func newCoreProperties(document types.Document) *CoreProperties {
	return &CoreProperties{
		document: document,
	}
}

func (cp *CoreProperties) Path() string {
	return "docProps/core.xml"
}

func (cr *CoreProperties) Byte() ([]byte, error) {
	metadata := cr.document.Metadata().Get()
	props := &CorePropertiesXML{
		XmlnsCp:       "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
		XmlnsDc:       "http://purl.org/dc/elements/1.1/",
		XmlnsDcterms:  "http://purl.org/dc/terms/",
		XmlnsDcmitype: "http://purl.org/dc/dcmitype/",
		XmlnsXsi:      "http://www.w3.org/2001/XMLSchema-instance",

		Title:          metadata.Title,
		Subject:        metadata.Subject,
		Creator:        metadata.Creator,
		Keywords:       metadata.Keywords,
		Description:    metadata.Description,
		LastModifiedBy: metadata.LastModifiedBy,
		Revision:       metadata.Revision,
		Category:       metadata.Category,
		ContentStatus:  metadata.ContentStatus,
		Language:       metadata.Language,
		Version:        metadata.Version,
	}

	// Format dates
	props.Created = &TimeElement{
		Type:  "dcterms:W3CDTF",
		Value: formatDateTime(metadata.Created),
	}

	props.Modified = &TimeElement{
		Type:  "dcterms:W3CDTF",
		Value: formatDateTime(metadata.Modified),
	}

	cr.CorePropertiesXML = props

	var buf bytes.Buffer

	// Write XML declaration
	buf.WriteString(xml.Header)

	// Encode the struct
	enc := xml.NewEncoder(&buf)
	enc.Indent("", "  ")

	if err := enc.Encode(cr.CorePropertiesXML); err != nil {
		return nil, fmt.Errorf("encoding ContentTypes XML: %w", err)
	}

	log.Printf("'%s' has been created.\n", cr.Path())
	// log.Print(buf.String())

	return buf.Bytes(), nil

}

func (cp *CoreProperties) WriteTo(w io.Writer) (int64, error) {
	content, err := cp.Byte()
	if err != nil {
		return 0, err
	}

	n, err := w.Write(content)
	return int64(n), err
}
