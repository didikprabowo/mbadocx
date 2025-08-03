package main

import (
	"log"
	"os"

	"github.com/didikprabowo/mbadocx"
)

func main() {
	docx := mbadocx.New()

	// Heading 1
	h1 := docx.AddHeading(mbadocx.StyleHeading1)
	h1.AddText("Go DOCX writer")

	// Paragraph 1
	p1 := docx.AddParagraph()

	p1.AddText("Hello from Go DOCX writer!")
	p1.AddText("This is bold.").SetBold(true)
	p1.AddText("Italic.").SetItalic(true)
	p1.AddText("And a ")
	p1.AddHyperlink("link.", "http://github.com")

	// Create document
	file, err := os.Create("example.docx")
	if err != nil {
		log.Fatal(err)
	}

	defer file.Close()

	err = docx.Write(file)
	if err != nil {
		log.Fatal(err)
	}
}
