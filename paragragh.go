package mbadocx

import "github.com/didikprabowo/mbadocx/elements"

// Content Methods

// AddParagraph
func (d *Document) AddParagraph() *elements.Paragraph {
	p := elements.NewParagraph(d)
	d.body.Elements = append(d.body.Elements, p)
	return p
}
