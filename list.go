package mbadocx

import "github.com/didikprabowo/mbadocx/elements"

// addList is a private helper method that handles creation of all list types.
// It provides a single implementation for adding lists to avoid code duplication
// across the various public list methods.
//
// This method creates a single paragraph containing all list items with the
// specified numbering style and indentation level.
//
// Parameters:
//   - items: Array of strings representing each list item's text content
//   - listType: The numbering/bullet style to apply (bullet, decimal, roman, etc.)
//   - lvl: Indentation level (0-8), where 0 is the root level and higher numbers
//     create nested sub-lists
//
// Returns:
//   - *elements.Paragraph: The paragraph containing the formatted list items
//
// Implementation note: All items are added to a single paragraph with numbering
// properties. For more complex list structures with multiple paragraphs,
// consider using separate paragraph elements for each item.
func (d *Document) addList(items []string, listType elements.ListType, lvl int) *elements.Paragraph {
	// Create a new paragraph that will contain all list items
	p := elements.NewParagraph(d)

	// Apply numbering style and add text for each item
	for _, item := range items {
		p.SetNumbering(listType, lvl).AddText(item)
	}

	// Add the completed paragraph to the document body
	d.body.AddElement(p)
	return p
}

// AddBulletList creates a bulleted list in the document.
// Each item appears with a bullet point marker (•, ○, ■, etc. depending on level).
//
// Parameters:
//   - items: String array where each element becomes a list item
//   - lvl: Indentation level (0-8). Level 0 is the main list, higher levels create sub-lists
//
// Returns:
//   - *elements.Paragraph: The paragraph containing the bulleted list for further customization
//
// Example:
//
//	doc := mbadocx.New()
//
//	// Create a simple bulleted list
//	doc.AddBulletList([]string{
//	    "First item",
//	    "Second item",
//	    "Third item",
//	}, 0)
//
//	// Create a nested sub-list
//	doc.AddBulletList([]string{
//	    "Sub-item A",
//	    "Sub-item B",
//	}, 1)
//
// Note: Bullet styles typically vary by level (e.g., • for level 0, ○ for level 1, ■ for level 2)
func (d *Document) AddBulletList(items []string, lvl int) *elements.Paragraph {
	return d.addList(items, elements.ListTypeBullet, lvl)
}

// AddNumberedList creates a numbered list with decimal numbering (1, 2, 3, ...).
// This is the most common numbered list format used in documents.
//
// Parameters:
//   - items: String array where each element becomes a numbered item
//   - lvl: Indentation level (0-8). Nested levels may use different numbering schemes
//
// Returns:
//   - *elements.Paragraph: The paragraph containing the numbered list
//
// Example:
//
//	doc := mbadocx.New()
//
//	// Create a numbered list of steps
//	doc.AddNumberedList([]string{
//	    "Download the software",
//	    "Run the installer",
//	    "Follow the setup wizard",
//	    "Restart your computer",
//	}, 0)
//
// Note: Numbering automatically increments. For nested lists, numbering may change
// to letters (a, b, c) or other schemes depending on the document template.
func (d *Document) AddNumberedList(items []string, lvl int) *elements.Paragraph {
	return d.addList(items, elements.ListTypeDecimal, lvl)
}

// AddLegalList creates a legal-style numbered list (1., 1.1., 1.1.1., ...).
// This format is commonly used in legal documents, contracts, and formal specifications
// where hierarchical numbering with full context is required.
//
// Parameters:
//   - items: String array where each element becomes a legal-numbered item
//   - lvl: Indentation level (0-8). Each level adds another decimal segment
//
// Returns:
//   - *elements.Paragraph: The paragraph containing the legal-style list
//
// Example:
//
//	doc := mbadocx.New()
//
//	// Create a contract-style numbered list
//	doc.AddLegalList([]string{
//	    "Terms and Conditions",
//	    "Payment Terms",
//	    "Delivery Schedule",
//	}, 0)
//
//	// Add sub-clauses
//	doc.AddLegalList([]string{
//	    "Payment due within 30 days",
//	    "Late fees apply after grace period",
//	}, 1)
//
// Output format: 1. Terms and Conditions
//
//	1.1. Payment due within 30 days
//	1.2. Late fees apply after grace period
func (d *Document) AddLegalList(items []string, lvl int) *elements.Paragraph {
	return d.addList(items, elements.ListTypeLegal, lvl)
}

// AddRomanList creates a list with Roman numeral numbering (I, II, III, IV, ...).
// This format is often used for major sections, chapters, or formal outlines
// in academic and classical documents.
//
// Parameters:
//   - items: String array where each element becomes a Roman-numbered item
//   - lvl: Indentation level (0-8). Nested levels may use lowercase romans (i, ii, iii)
//
// Returns:
//   - *elements.Paragraph: The paragraph containing the Roman numeral list
//
// Example:
//
//	doc := mbadocx.New()
//
//	// Create chapter headings with Roman numerals
//	doc.AddRomanList([]string{
//	    "Introduction",
//	    "Literature Review",
//	    "Methodology",
//	    "Results",
//	    "Conclusion",
//	}, 0)
//
// Note: Level 0 typically uses uppercase (I, II, III), while level 1 often
// uses lowercase (i, ii, iii). Very large numbers may not display correctly
// in some Word versions (e.g., numbers above 3999).
func (d *Document) AddRomanList(items []string, lvl int) *elements.Paragraph {
	return d.addList(items, elements.ListTypeRoman, lvl)
}
