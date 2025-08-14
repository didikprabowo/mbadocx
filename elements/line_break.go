package elements

// LineBreak represents a line break within a run
type LineBreak struct {
	Typ   string // Type of break: "textWrapping" (default), "page", "column", "textWrapping"
	Clear string // Clear type for text wrapping: "none", "left", "right", "all"
}

// NewLineBreak creates a new line break
func NewLineBreak() *LineBreak {
	return &LineBreak{
		Typ: "textWrapping",
	}
}

// NewColumnBreak creates a new column break
func NewColumnBreak() *LineBreak {
	return &LineBreak{
		Typ: "column",
	}
}

// NewTextWrappingBreak creates a line break that clears floating objects
func NewTextWrappingBreak(clear string) *LineBreak {
	return &LineBreak{
		Typ:   "textWrapping",
		Clear: clear,
	}
}

// Type returns the element type
func (lb *LineBreak) Type() string {
	return "lineBreak"
}

// XML generates the XML representation
func (lb *LineBreak) XML() ([]byte, error) {
	if lb.Typ == "" || lb.Typ == "textWrapping" {
		if lb.Clear != "" && lb.Clear != "none" {
			return []byte(`<w:br w:type="textWrapping" w:clear="` + lb.Clear + `"/>`), nil
		}
		return []byte(`<w:br/>`), nil
	}

	return []byte(`<w:br w:type="` + lb.Typ + `"/>`), nil
}
