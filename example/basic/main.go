package main

import (
	"log"

	"github.com/didikprabowo/mbadocx"
)

func main() {
	docx := mbadocx.New()
	defer docx.Close()

	// Add heading 2
	docx.AddHeading("Lorem Ipsum", 2)

	// Add 1st paragraph
	p1 := docx.AddParagraph()
	p1.AddText("Lorem Ipsum").AddSpace(1).SetBold(true)
	p1.AddText("is simply dummy text of the printing and typesetting industry.").AddSpace(1)
	p1.AddHyperlink("More detail able to clik", "http://tes.com")

	// Add text on the next page
	docx.AddParagraph().AddPageBreak().AddText("Example")

	err := docx.Save("testdata/basic.docx")
	if err != nil {
		log.Fatal(err)
	}
}
