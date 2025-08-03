package mbadocx

import (
	"fmt"
	"log"

	"github.com/didikprabowo/mbadocx/elements"
)

// TextStyle represents a paragraph or run style
type TextStyle string

const (
	StyleNormal   TextStyle = "Normal"
	StyleHeading  TextStyle = "Heading"
	StyleHeading1 TextStyle = "Heading1"
	StyleHeading2 TextStyle = "Heading2"
	StyleHeading3 TextStyle = "Heading3"
	StyleHeading4 TextStyle = "Heading4"
	StyleHeading5 TextStyle = "Heading5"
	StyleHeading6 TextStyle = "Heading6"
)

var (
	ListStyleHeading = []TextStyle{
		StyleNormal,
		StyleHeading,
		StyleHeading1,
		StyleHeading2,
		StyleHeading3,
		StyleHeading4,
		StyleHeading5,
		StyleHeading6,
	}
)

// Validate
func (style TextStyle) Validate(textStyle TextStyle) error {
	for _, style := range ListStyleHeading {
		if textStyle == style {
			return nil
		}
	}

	return fmt.Errorf("the heading style is invalid")
}

// AddHeading
func (d *Document) AddHeading(style TextStyle) *elements.Paragraph {
	err := style.Validate(style)
	if err != nil {
		log.Printf("ERROR: %v", err)
		style = StyleNormal
	}

	p := elements.NewParagraph(d)
	p.SetStyle(string(style))

	d.body.Elements = append(d.body.Elements, p)
	return p
}
