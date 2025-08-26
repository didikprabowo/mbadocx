package mbadocx

// AddPageBreak inserts a page break in the document.
// This forces content following the break to begin on a new page.
// The page break is added as a new paragraph containing only the break element.
//
// Example:
//
//	doc := mbadocx.New()
//	doc.AddParagraph().AddText("Chapter 1 content here")
//	doc.AddPageBreak() // Start Chapter 2 on a new page
//	doc.AddParagraph().AddText("Chapter 2 content here")
//
// Note: The page break creates a new empty paragraph internally,
// so you don't need to add a paragraph before or after calling this method.
func (d *Document) AddPageBreak() {
	d.AddParagraph().AddPageBreak()
}

// AddLineBreak inserts a line break (soft return) in the document.
// Unlike AddPageBreak, this only moves to the next line within the current page,
// similar to pressing Shift+Enter in Microsoft Word.
// The line break is added as a new paragraph containing only the break element.
//
// Example:
//
//	doc := mbadocx.New()
//	doc.AddParagraph().AddText("First line")
//	doc.AddLineBreak() // Move to next line without starting a new paragraph
//	doc.AddParagraph().AddText("Second line")
//
// Note: This creates a visual line break while maintaining paragraph formatting.
// For multiple line breaks, call this method multiple times or use AddParagraph()
// for proper paragraph spacing.
func (d *Document) AddLineBreak() {
	d.AddParagraph().AddLineBreak()
}
