package mbadocx

import (
	"github.com/didikprabowo/mbadocx/types"
)

// Body contains the main content of the document
type Body struct {
	Elements []types.Element
}

func NewBody() *Body {
	return &Body{Elements: make([]types.Element, 0)}
}

// AddElement adds any element to the body
func (b *Body) AddElement(element types.Element) {
	b.Elements = append(b.Elements, element)
}

// GetElements
func (b *Body) GetElements() []types.Element {
	return b.Elements
}

// Clear removes all elements from the body
func (b *Body) Clear() {
	b.Elements = b.Elements[:0]
}
