package elements

import (
	"bytes"
	"encoding/xml"
	"strings"
)

// Text represents text content
type Text struct {
	Value         string
	PreserveSpace bool
}

// NewText creates a new text element
func NewText(value string) *Text {
	return &Text{
		Value:         value,
		PreserveSpace: strings.Contains(value, "  ") || strings.HasPrefix(value, " ") || strings.HasSuffix(value, " "),
	}
}

// Type returns the element type
func (t *Text) Type() string {
	return "text"
}

// XML generates the XML representation of the text
func (t *Text) XML() ([]byte, error) {
	var buf bytes.Buffer

	// Start text tag
	if t.PreserveSpace {
		buf.WriteString(`<w:t xml:space="preserve">`)
	} else {
		buf.WriteString(`<w:t>`)
	}

	// Escape XML special characters
	err := xml.EscapeText(&buf, []byte(t.Value))
	if err != nil {
		return nil, err
	}

	// Close text tag
	buf.WriteString(`</w:t>`)

	return buf.Bytes(), nil
}
