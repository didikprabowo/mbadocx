package mbadocx

import "github.com/didikprabowo/mbadocx/elements"

// addList is a private helper method that handles all list types
func (d *Document) addList(items []string, listType elements.ListType, lvl int) *elements.Paragraph {
	p := elements.NewParagraph(d)

	for _, item := range items {
		p.SetNumbering(listType, lvl).AddText(item)
	}

	d.body.AddElement(p)
	return p
}

// AddBulletList
func (d *Document) AddBulletList(items []string, lvl int) *elements.Paragraph {
	return d.addList(items, elements.ListTypeBullet, lvl)
}

// AddNumberedList
func (d *Document) AddNumberedList(items []string, lvl int) *elements.Paragraph {
	return d.addList(items, elements.ListTypeDecimal, lvl)
}

// AddLegalList
func (d *Document) AddLegalList(items []string, lvl int) *elements.Paragraph {
	return d.addList(items, elements.ListTypeLegal, lvl)
}

// AddRomanList
func (d *Document) AddRomanList(items []string, lvl int) *elements.Paragraph {
	return d.addList(items, elements.ListTypeRoman, lvl)
}
