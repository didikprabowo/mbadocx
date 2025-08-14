package contenttypes

import "encoding/xml"

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

func NewDefaultContentType() *ContentTypes {
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

// Get
func (ct *ContentTypes) Get() *ContentTypes {
	return ct
}
