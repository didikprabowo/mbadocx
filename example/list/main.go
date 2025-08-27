package main

import (
	"log"

	"github.com/didikprabowo/mbadocx"
	"github.com/didikprabowo/mbadocx/elements"
)

func main() {
	doc := mbadocx.New()

	// Section: Bulleted List
	doc.AddHeading("Bulleted List Example", 2)

	// Top-level bullet items
	bulletItems := []string{
		"First bullet item",
		"Second bullet item",
	}
	for _, text := range bulletItems {
		p := doc.AddParagraph()
		p.SetNumbering(elements.ListTypeBullet, 0) // Level 0 (top-level)
		p.AddText(text)
	}

	// Sub-list items (nested bullets)
	subBulletItems := []string{
		"Sub bullet item 1",
		"Sub bullet item 2",
		"Sub bullet item 3",
		"Sub bullet item 4",
	}
	for i, text := range subBulletItems {
		level := 1
		if i == 3 {
			level = 2 // Example: deeper nesting for last sub-item
		}
		p := doc.AddParagraph()
		p.SetNumbering(elements.ListTypeBullet, level)
		p.AddText(text)
	}

	// Section: Numbered List
	doc.AddHeading("Numbered List Example", 2)

	numberedItems := []string{
		"First numbered item",
		"Second numbered item",
	}
	for _, text := range numberedItems {
		p := doc.AddParagraph()
		p.SetNumbering(elements.ListTypeDecimal, 0) // Level 0 (top-level)
		p.AddText(text)
	}

	if err := doc.Save("testdata/list_example.docx"); err != nil {
		log.Fatal(err)
	}
}
