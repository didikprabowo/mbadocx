// Word UI Option	XML Needed
// Align Top Justified	<w:vAlign w:val="top"/> + <w:jc w:val="both"/>
// Align Top Center	<w:vAlign w:val="top"/> + <w:jc w:val="center"/>
// Align Top Right	<w:vAlign w:val="top"/> + <w:jc w:val="right"/>
// Align Center Justified	<w:vAlign w:val="center"/> + <w:jc w:val="both"/>
// Align Center	<w:vAlign w:val="center"/> + <w:jc w:val="center"/>
// Align Center Right	<w:vAlign w:val="center"/> + <w:jc w:val="right"/>
// Align Bottom Justified	<w:vAlign w:val="bottom"/> + <w:jc w:val="both"/>
// Align Bottom Center (U)	<w:vAlign w:val="bottom"/> + <w:jc w:val="center"/>
// Align Bottom Right	<w:vAlign w:val="bottom"/> + <w:jc w:val="right"/>

package main

import (
	"log"

	"github.com/didikprabowo/mbadocx"
	"github.com/didikprabowo/mbadocx/elements"
)

func main() {
	doc := mbadocx.New()

	// Add title
	doc.AddHeading("Table Examples in mbadocx", 1)
	doc.AddParagraph().AddText("This document demonstrates various table features.")

	// Example 1: Simple table
	doc.AddHeading("Example 1: Basic Table", 2)
	doc.AddParagraph().AddText("A simple 3x3 table:")

	table1 := doc.AddTable(3, 3)

	// Add headers
	_ = table1.SetCellFormattedText(0, 0, "Name", func(r *elements.Run) {
		r.SetSpacing(2)
		r.SetBold(true)
	})
	_ = table1.SetCellFormattedText(0, 1, "Age", func(r *elements.Run) {
		r.SetBold(true)
	})
	_ = table1.SetCellFormattedText(0, 2, "City", func(r *elements.Run) {
		r.SetBold(true)
	})

	// Add data
	_ = table1.SetCellText(1, 0, "John Doe")
	_ = table1.SetCellText(1, 1, "30")
	_ = table1.SetCellText(1, 2, "New York")

	_ = table1.SetCellText(2, 0, "Jane Smith")
	_ = table1.SetCellText(2, 1, "25")
	_ = table1.SetCellText(2, 2, "London")

	// Set table properties
	table1.SetTableAlignment("center")

	doc.AddParagraph().AddText("") // Add spacing

	// Example 2: Table from data array
	doc.AddHeading("Example 2: Table from Data", 2)
	doc.AddParagraph().AddText("Creating table from 2D array:")

	data := [][]string{
		{"Product", "Price", "Stock"},
		{"Laptop", "$999", "15"},
		{"Mouse", "$25", "50"},
		{"Keyboard", "$75", "30"},
	}

	table2 := doc.AddTableWithData(data)
	if table2 != nil {
		// Make first row bold (headers)
		for i := 0; i < 3; i++ {
			cell := table2.Rows[0].Cells[i]
			if len(cell.Paragraphs) > 0 {
				// Clear and re-add with formatting
				text := getTextFromCell(cell)
				cell.Paragraphs[0].Clear()
				cell.Paragraphs[0].AddFormattedText(text, func(r *elements.Run) {
					r.SetBold(true)
				})
			}
		}

		// Add shading to header row
		for i := 0; i < 3; i++ {
			_ = table2.SetCellShading(0, i, "D9D9D9")
		}
	}

	// Example 3: Table with headers helper
	doc.AddHeading("Example 3: Table with Headers", 2)
	doc.AddParagraph().AddText("") // Add spacing
	doc.AddParagraph().AddText("Using the AddTableWithHeaders method:")

	headers := []string{"Employee", "Department", "Salary", "Status"}
	employees := [][]string{
		{"Alice Johnson", "Engineering", "$120,000", "Full-time"},
		{"Bob Williams", "Marketing", "$95,000", "Full-time"},
		{"Charlie Brown", "Sales", "$85,000", "Contract"},
		{"Diana Prince", "HR", "$90,000", "Full-time"},
	}

	table3 := doc.AddTableWithHeaders(headers, employees)
	if table3 != nil {
		// Set table to full width
		table3.SetTableWidth("pct", "5000") // 100% width

		// Add alternating row colors
		for i := 1; i <= len(employees); i++ {
			if i%2 == 0 {
				for j := 0; j < len(headers); j++ {
					_ = table3.SetCellShading(i, j, "F0F0F0")
				}
			}
		}

		// Highlight contract employees
		for i, emp := range employees {
			if emp[3] == "Contract" {
				_ = table3.SetCellShading(i+1, 3, "FFEB9C") // Yellow highlight
			}
		}
	}

	doc.AddParagraph().AddText("")

	// Example 4: Complex table with merged cells
	doc.AddHeading("Example 4: Merged Cells", 2)
	doc.AddParagraph().AddText("Table with merged cells:")

	table4 := doc.AddTable(5, 4)

	// Title row - merge all columns
	_ = table4.SetCellFormattedText(0, 0, "Quarterly Sales Report 2024", func(r *elements.Run) {
		r.SetBold(true)
		r.SetFontSize(14)
	})
	_ = table4.MergeCells(0, 0, 3)
	_ = table4.SetCellShading(0, 0, "4472C4")

	// Headers
	_ = table4.SetCellText(1, 0, "Region")
	_ = table4.SetCellText(1, 1, "Q1")
	_ = table4.SetCellText(1, 2, "Q2")
	_ = table4.SetCellText(1, 3, "Total")

	// Make headers bold
	for i := 0; i < 4; i++ {
		cell := table4.Rows[1].Cells[i]
		if len(cell.Paragraphs) > 0 {
			text := getTextFromCell(cell)
			cell.Paragraphs[0].Clear()
			cell.Paragraphs[0].AddFormattedText(text, func(r *elements.Run) {
				r.SetBold(true)
			})
		}
	}

	// Data
	regions := [][]string{
		{"North", "$50,000", "$55,000", "$105,000"},
		{"South", "$45,000", "$48,000", "$93,000"},
		{"Total", "$95,000", "$103,000", "$198,000"},
	}

	for i, region := range regions {
		for j, value := range region {
			_ = table4.SetCellText(i+2, j, value)
		}
	}

	// Style the total row
	for j := 0; j < 4; j++ {
		_ = table4.SetCellShading(4, j, "E0E0E0")
		cell := table4.Rows[4].Cells[j]
		if len(cell.Paragraphs) > 0 {
			text := getTextFromCell(cell)
			cell.Paragraphs[0].Clear()
			cell.Paragraphs[0].AddFormattedText(text, func(r *elements.Run) {
				r.SetBold(true)
				r.SetVerticalAlign("baseline")
			})
		}
	}

	doc.AddParagraph().AddText("")

	// Example 5: Custom column widths
	doc.AddHeading("Example 5: Custom Column Widths", 2)
	doc.AddParagraph().AddText("Table with custom column widths:")

	table5 := doc.AddTable(3, 3)

	// Set different widths for columns
	_ = table5.SetColumnWidth(0, "1500") // Narrow
	_ = table5.SetColumnWidth(1, "3000") // Medium
	_ = table5.SetColumnWidth(2, "4500") // Wide

	// Add content
	_ = table5.SetCellText(0, 0, "Narrow")
	_ = table5.SetCellText(0, 1, "Medium Width")
	_ = table5.SetCellText(0, 2, "Wide Column")

	_ = table5.SetCellText(1, 0, "A")
	_ = table5.SetCellText(1, 1, "This is medium")
	_ = table5.SetCellText(1, 2, "This column has more space")

	_ = table5.SetCellText(2, 0, "B")
	_ = table5.SetCellText(2, 1, "Another row")
	_ = table5.SetCellText(2, 2, "With different content lengths")

	// Add borders highlight
	table5.Properties.Borders = &elements.TableBorders{
		Top:     &elements.BorderStyle{Value: "double", Size: "6", Color: "4472C4"},
		Bottom:  &elements.BorderStyle{Value: "double", Size: "6", Color: "4472C4"},
		Left:    &elements.BorderStyle{Value: "single", Size: "4", Color: "4472C4"},
		Right:   &elements.BorderStyle{Value: "single", Size: "4", Color: "4472C4"},
		InsideH: &elements.BorderStyle{Value: "single", Size: "2", Color: "808080"},
		InsideV: &elements.BorderStyle{Value: "single", Size: "2", Color: "808080"},
	}

	// Add some final text
	doc.AddParagraph().AddText("")
	doc.AddHeading("Conclusion", 2)
	doc.AddParagraph().AddText("These examples demonstrate the table functionality in mbadocx library.")
	doc.AddParagraph().AddText("Tables can be customized with:")

	// Add a bullet list
	lists := []string{
		"Custom column widths",
		"Cell merging",
		"Background colors (shading)",
		"Text formatting (bold, colors, sizes)",
		"Borders and alignment",
		"Header rows",
	}

	for _, text := range lists {
		p := doc.AddParagraph()
		p.SetNumbering(elements.ListTypeBullet, 0) // Level 0 (top-level)
		p.AddText(text)
	}

	// Save the document
	if err := doc.Save("testdata/table_examples.docx"); err != nil {
		log.Fatalf("Failed to save document: %v", err)
	}
}

// Helper function to get text from a cell
func getTextFromCell(cell *elements.TableCell) string {
	if len(cell.Paragraphs) > 0 && len(cell.Paragraphs[0].Children) > 0 {
		if run, ok := cell.Paragraphs[0].Children[0].(*elements.Run); ok {
			if text, ok := run.Children[0].(*elements.Text); ok {
				return text.Value
			}
		}
	}
	return ""
}
