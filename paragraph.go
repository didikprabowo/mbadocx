package mbadocx

import (
	"github.com/didikprabowo/mbadocx/elements"
)

// AddParagraph creates and adds a new paragraph element to the document body.
//
// A paragraph is the fundamental text container in Word documents. It can contain
// formatted text, hyperlinks, images, and other inline elements. Each paragraph
// represents a distinct block of content that ends with a paragraph mark (¶).
//
// The returned Paragraph pointer allows for method chaining to add content
// and apply formatting in a fluent interface style.
//
// Returns:
//   - *elements.Paragraph: A pointer to the newly created paragraph that can be
//     used to add text, apply formatting, set alignment, etc.
//
// Example:
//
//	doc := mbadocx.New()
//
//	// Create a simple paragraph with text
//	doc.AddParagraph().AddText("This is a simple paragraph.")
//
//	// Create a paragraph with multiple formatting
//	para := doc.AddParagraph()
//	para.AddText("This text is ").
//	     AddText("bold").SetBold(true).
//	     AddText(" and this is ").
//	     AddText("italic").SetItalic(true).
//	     AddText(".")
//
//	// Create a paragraph with alignment and spacing
//	doc.AddParagraph().
//	    AddText("Centered text").
//	    SetAlignment(elements.AlignCenter).
//	    SetSpacingAfter(240) // 240 twips = 12pt
//
//	// Add a paragraph with a hyperlink
//	doc.AddParagraph().
//	    AddText("Visit our ").
//	    AddHyperlink("website", "https://example.com").
//	    AddText(" for more information.")
//
// Common paragraph operations after creation:
//   - AddText(string): Add plain or formatted text
//   - AddHyperlink(text, url): Add a clickable link
//   - AddImage(path): Add an inline image
//   - AddLineBreak(): Add a line break within the paragraph
//   - AddPageBreak(): Insert a page break after this paragraph
//   - SetAlignment(align): Set text alignment (left, center, right, justify)
//   - SetIndentation(left, right, firstLine): Set paragraph indentation
//   - SetSpacing(before, after, line): Set paragraph spacing
//   - SetStyle(styleName): Apply a predefined paragraph style
//   - SetNumbering(type, level): Convert to a numbered or bulleted list item
//
// Note: Empty paragraphs are valid and often used for spacing. In Word,
// pressing Enter creates a new paragraph, even if it contains no text.
//
// Performance: O(1) - Direct insertion at the end of the document body.
// Memory: Minimal allocation for the paragraph structure.
//
// Thread Safety: This method is not thread-safe. If multiple goroutines need
// to add paragraphs concurrently, synchronization should be handled externally.
func (d *Document) AddParagraph() *elements.Paragraph {
	// Create a new paragraph element associated with this document
	// The paragraph maintains a reference to its parent document for
	// context-aware operations like style inheritance and numbering
	paragraphElem := elements.NewParagraph(d)

	// Add the paragraph to the document body's element collection
	// This ensures the paragraph appears in the document's XML structure
	// and will be rendered when the document is saved
	d.body.AddElement(paragraphElem)

	// Return the paragraph reference for method chaining
	// This allows users to immediately add content and formatting
	return paragraphElem
}
