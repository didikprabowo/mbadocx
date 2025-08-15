package main

import (
	"github.com/didikprabowo/mbadocx"
)

func main() {
	doc := mbadocx.New()

	// custom font size heading 1
	heading := doc.AddHeading("", 1)
	heading.AddText("This is heading 1").SetFontSize(24)

	para1 := doc.AddParagraph()
	// Add bold text
	para1.AddText("Bold Text").SetBold(true)
	// Add italic text
	para1.AddText(" Italic Text").SetItalic(true)
	// Add underlined and colored text
	para1.AddText(" Underlined & Colored").SetUnderline("single").SetColor("#E74C3C")
	// Add large font text
	para1.AddText(" Large Font").SetFontSize(20)

	para2 := doc.AddParagraph()
	// Set paragraph alignment to center
	para2.SetAlignment("center")
	// Set spacing before and after the paragraph
	para2.SetSpacing(12, 12) // 12pt before, 12pt after
	// Set indentation: 20pt left, 0pt right, 30pt first line
	para2.SetIndentation(20, 0, 30)
	para2.AddText("This paragraph is centered, spaced, and indented.")

	para3 := doc.AddParagraph()
	para3.SetAlignment("right")
	para3.AddText("This paragraph on right")

	para4 := doc.AddParagraph()
	para4.SetAlignment("both")
	para4.AddText("This is a paragraph demonstrating justified alignment. When text is long enough to wrap across multiple lines.")

	doc.Save("testdata/formatted_text.docx")
}
