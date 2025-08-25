package mbadocx

import (
	"github.com/didikprabowo/mbadocx/elements"
)

// AddImage adds an image to the document
func (d *Document) AddImage(imagePath string) (*elements.Image, error) {
	img, err := elements.NewImage(d, imagePath)
	if err != nil {
		return nil, err
	}

	p := elements.NewParagraph(d)
	// Add to children
	p.AddChildren(img)
	// Add to element
	d.body.AddElement(p)
	// Add to media
	d.media.AddMedia(img)

	return img, nil
}
