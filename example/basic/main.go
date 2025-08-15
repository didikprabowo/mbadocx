package main

import (
	"github.com/didikprabowo/mbadocx"
)

func main() {
	// Create a new document
	doc := mbadocx.New()

	// Add a paragraph with formatted text
	para := doc.AddParagraph()
	para.AddText("Hello, mbadocx!").SetBold(true).SetColor("#2E86C1").SetFontSize(16)

	// Add a hyperlink
	para.AddHyperlink("Visit GitHub", "https://github.com/didikprabowo/mbadocx")

	// Add a line break and another paragraph
	para.AddLineBreak()
	doc.AddParagraph().AddText("This is a second paragraph.")

	// Save the document
	if err := doc.Save("testdata/basic.docx"); err != nil {
		panic(err)
	}
}
