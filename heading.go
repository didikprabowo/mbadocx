package mbadocx

import (
	"fmt"

	"github.com/didikprabowo/mbadocx/elements"
)

// AddHeading adds a heading paragraph
func (d *Document) AddHeading(text string, level int) *elements.Paragraph {
	if level < 1 || level > 9 {
		level = 1
	}
	styleID := fmt.Sprintf("Heading%d", level)

	p := d.AddParagraph()
	p.SetStyle(styleID)
	if text != "" {
		p.AddText(text)
	}

	return p
}
