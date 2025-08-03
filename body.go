package mbadocx

import "github.com/didikprabowo/mbadocx/types"

// Body contains the main content of the document
type Body struct {
	Elements []types.Element
}

func (b *Body) GetElements() []types.Element {
	return b.Elements
}
