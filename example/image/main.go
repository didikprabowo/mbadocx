package main

import (
	"log"

	"github.com/didikprabowo/mbadocx"
	"github.com/didikprabowo/mbadocx/properties"
)

func main() {
	docx := mbadocx.New()

	// imageWithBorderaAndEffects
	imageWithBordersAndEffects(docx)

	// floatingImageExample
	floatingImageExample(docx)

	if err := docx.Save("testdata/image.docx"); err != nil {
		log.Fatal(err)
	}
}

func imageWithBordersAndEffects(docx *mbadocx.Document) {
	docx.AddParagraph().AddText("Image with Borders and Effects")

	img1, err := docx.AddImage("mbadocx_logo.png")
	if err != nil {
		log.Fatal("failed add image")
	}

	img1.
		SetBorder(3, "#FF0000").
		SetAlignment(properties.AlignCenter)

	img2, err := docx.AddImage("mbadocx_logo.png")
	if err != nil {
		log.Fatal("failed add image")
	}

	img2.
		SetShadow(true).
		SetAlignment(properties.AlignCenter).
		SetRotation(15)
}

func floatingImageExample(doc *mbadocx.Document) {
	doc.AddParagraph().AddText("Example 4: Floating Images with Text Wrapping")

	// Add some text
	para := doc.AddParagraph().AddText("This is a paragraph with a floating image. ")
	para.AddText("The image will be positioned to the right of this text with square text wrapping. ")
	para.AddText("Lorem ipsum dolor sit amet, consectetur adipiscing elit. ")
	para.AddText("Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. ")
	para.AddText("Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris.")

	// Floating image with square wrapping
	img1, err := doc.AddImage("mbadocx_logo.png")
	if err != nil {
		log.Printf("Error loading image: %v", err)
		return
	}

	img1.
		SetSize(2, 2).
		SetWrapStyle(properties.WrapSquare).
		SetFloating(properties.HorizontalAnchorInsideMargin, properties.VerticalAnchorParagraph).
		SetOffsetInches(4, 0) // Position 4 inches from left margin

	// Add more text
	para2 := doc.AddParagraph().AddText("This is another paragraph demonstrating text wrapping around images. ")
	para2.AddText("The image can be positioned in various ways relative to the text. ")
	para2.AddText("You can control the wrapping style, position, and offset of the image.")
}
