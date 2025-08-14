package elements

// PageBreak represents a page break element
type PageBreak struct{}

// NewPageBreak creates a new page break
func NewPageBreak() *PageBreak {
	return &PageBreak{}
}

// Type returns the element type
func (pb *PageBreak) Type() string {
	return "pageBreak"
}

// XML generates the XML representation
func (pb *PageBreak) XML() ([]byte, error) {
	// Page break is a paragraph with a page break run
	return []byte(`<w:p><w:r><w:br w:type="page"/></w:r></w:p>`), nil
}
