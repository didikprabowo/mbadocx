package mbadocx

import "github.com/didikprabowo/mbadocx/elements"

// Content Methods

// AddParagraph
func (d *Document) AddParagraph() *elements.Paragraph {
	paragraphElem := elements.NewParagraph(d)
	d.body.AddElement(paragraphElem)
	return paragraphElem
}
