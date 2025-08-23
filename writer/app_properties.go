package writer

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"log"
	"strings"

	"github.com/didikprabowo/mbadocx/types"
)

// AppProperties represents the generator for docProps/app.xml
type AppProperties struct {
	document types.Document
}

// AppPropertiesXML defines the XML structure of docProps/app.xml
type AppPropertiesXML struct {
	XMLName xml.Name `xml:"Properties"`
	Xmlns   string   `xml:"xmlns,attr"`
	XmlnsVt string   `xml:"xmlns:vt,attr"`

	Application          string `xml:"Application,omitempty"`
	AppVersion           string `xml:"AppVersion,omitempty"`
	DocSecurity          int    `xml:"DocSecurity"`
	Lines                int    `xml:"Lines"`
	Paragraphs           int    `xml:"Paragraphs"`
	Words                int    `xml:"Words"`
	Characters           int    `xml:"Characters"`
	CharactersWithSpaces int    `xml:"CharactersWithSpaces"`
	Pages                int    `xml:"Pages"`
	Company              string `xml:"Company,omitempty"`
	Manager              string `xml:"Manager,omitempty"`
	LinksUpToDate        bool   `xml:"LinksUpToDate"`
	ScaleCrop            bool   `xml:"ScaleCrop"`
	SharedDoc            bool   `xml:"SharedDoc"`
	HyperlinksChanged    bool   `xml:"HyperlinksChanged"`
	Template             string `xml:"Template,omitempty"`
	TotalTime            int    `xml:"TotalTime,omitempty"`
	HiddenSlides         int    `xml:"HiddenSlides,omitempty"`
	MMClips              int    `xml:"MMClips,omitempty"`
	Notes                int    `xml:"Notes,omitempty"`
	Slides               int    `xml:"Slides,omitempty"`
}

// Ensure AppProperties implements zipWritable interface
var _ zipWritable = (*AppProperties)(nil)

// newAppProperties creates a new AppProperties writer
func newAppProperties(document types.Document) *AppProperties {
	return &AppProperties{document: document}
}

// Path returns the ZIP path for app properties
func (ap *AppProperties) Path() string {
	return "docProps/app.xml"
}

// Byte generates the XML content for docProps/app.xml
func (ap *AppProperties) Byte() ([]byte, error) {
	metadata := ap.document.Metadata().Get()

	props := &AppPropertiesXML{
		Xmlns:   "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
		XmlnsVt: "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",

		Application:          "Go DOCX Library",
		AppVersion:           "1.0",
		DocSecurity:          0,
		Lines:                ap.countLines(),
		Paragraphs:           ap.countParagraphs(),
		Words:                ap.countWords(),
		Characters:           ap.countCharacters(),
		CharactersWithSpaces: ap.countCharactersWithSpaces(),
		Pages:                1, // Approximation; Word will recalculate
		Company:              metadata.Company,
		Manager:              metadata.Manager,
		LinksUpToDate:        false,
		ScaleCrop:            false,
		SharedDoc:            false,
		HyperlinksChanged:    false,
	}

	var buf bytes.Buffer
	buf.WriteString(xml.Header)

	enc := xml.NewEncoder(&buf)
	enc.Indent("", "  ")

	if err := enc.Encode(props); err != nil {
		return nil, fmt.Errorf("encoding AppProperties XML: %w", err)
	}

	log.Printf("'%s' has been created.\n", ap.Path())
	// log.Print(buf.String())

	return buf.Bytes(), nil
}

// WriteTo writes the XML to an io.Writer (implements io.WriterTo).
func (ap *AppProperties) WriteTo(w io.Writer) (int64, error) {
	xmlData, err := ap.Byte()
	if err != nil {
		return 0, err
	}

	n, err := w.Write(xmlData)
	return int64(n), err
}

// countCharactersWithSpaces returns total character count including spaces
func (ap *AppProperties) countCharactersWithSpaces() int {
	total := 0
	for _, elem := range ap.document.Body().GetElements() {
		text := ap.getElementText(elem)
		for _, r := range text {
			if r != '\n' && r != '\r' {
				total++
			}
		}
	}
	return total
}

// countParagraphs returns the number of paragraph elements
func (ap *AppProperties) countParagraphs() int {
	count := 0
	for _, elem := range ap.document.Body().GetElements() {
		if elem.Type() == "paragraph" {
			count++
		}
	}
	return count
}

// countWords estimates the number of words in the document
func (ap *AppProperties) countWords() int {
	total := 0
	for _, elem := range ap.document.Body().GetElements() {
		text := ap.getElementText(elem)
		if text != "" {
			total += len(splitWords(text))
		}
	}
	return total
}

// countLines estimates number of lines based on character length per paragraph
func (ap *AppProperties) countLines() int {
	lines := 0
	for _, elem := range ap.document.Body().GetElements() {
		if elem.Type() == "paragraph" {
			lines++
			text := ap.getElementText(elem)
			if len(text) > 80 {
				lines += len(text) / 80
			}
		}
	}
	if lines == 0 {
		lines = 1
	}
	return lines
}

// countCharacters returns total non-whitespace character count
func (ap *AppProperties) countCharacters() int {
	total := 0
	for _, elem := range ap.document.Body().GetElements() {
		text := ap.getElementText(elem)
		for _, r := range text {
			if r != ' ' && r != '\t' && r != '\n' && r != '\r' {
				total++
			}
		}
	}
	return total
}

// getElementText extracts visible text from a raw XML element
func (ap *AppProperties) getElementText(elem types.Element) string {
	xmlData, err := elem.XML()
	if err != nil {
		return ""
	}
	text := string(xmlData)
	// Naively strip tags (for accurate results, implement a real parser)
	for {
		start := strings.Index(text, "<")
		if start == -1 {
			break
		}
		end := strings.Index(text[start:], ">")
		if end == -1 {
			break
		}
		text = text[:start] + text[start+end+1:]
	}
	return strings.TrimSpace(text)
}

// splitWords separates text into word-like tokens
func splitWords(text string) []string {
	var words []string
	var word strings.Builder

	for _, r := range text {
		if isWordChar(r) {
			word.WriteRune(r)
		} else if word.Len() > 0 {
			words = append(words, word.String())
			word.Reset()
		}
	}
	if word.Len() > 0 {
		words = append(words, word.String())
	}
	return words
}

// isWordChar checks if a rune is part of a word
func isWordChar(r rune) bool {
	return (r >= 'a' && r <= 'z') ||
		(r >= 'A' && r <= 'Z') ||
		(r >= '0' && r <= '9') ||
		r == '\'' || r == '-' || r == '_'
}
