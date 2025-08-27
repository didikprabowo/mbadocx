package mbadocx

import "github.com/didikprabowo/mbadocx/elements"

// AddTable creates and adds a new table with the specified dimensions to the document.
//
// This method creates an empty table structure that can be populated with data
// after creation. The table is initialized with the specified number of rows
// and columns, with all cells empty.
//
// Parameters:
//   - rows: Number of rows in the table (must be > 0)
//   - cols: Number of columns in the table (must be > 0)
//
// Returns:
//   - *elements.Table: A pointer to the newly created table for further customization
//     such as adding content, styling, borders, etc.
//
// Example:
//
//	doc := mbadocx.New()
//
//	// Create a 3x4 table (3 rows, 4 columns)
//	table := doc.AddTable(3, 4)
//
//	// Populate the table cells
//	table.SetCellText(0, 0, "Name")
//	table.SetCellText(0, 1, "Age")
//	table.SetCellText(0, 2, "Department")
//	table.SetCellText(0, 3, "Salary")
//
//	// Add data rows
//	table.SetCellText(1, 0, "John Doe")
//	table.SetCellText(1, 1, "30")
//	table.SetCellText(1, 2, "Engineering")
//	table.SetCellText(1, 3, "$75,000")
//
//	// Apply table styling
//	table.SetStyle("TableGrid")
//	table.SetWidth(100, elements.WidthPercent)
//
// Note: Empty tables are valid in Word documents. Cell indices are zero-based,
// so a 3x4 table has rows 0-2 and columns 0-3.
func (d *Document) AddTable(rows, cols int) *elements.Table {
	// Create a new table element with specified dimensions
	// The table maintains a reference to the document for style inheritance
	tableElem := elements.NewTable(d, rows, cols)

	// Add the table to the document body
	// This ensures the table appears in the document flow
	d.body.AddElement(tableElem)

	// Return the table for further configuration
	return tableElem
}

// AddTableWithData creates and populates a table from a 2D string array.
//
// This convenience method automatically determines table dimensions from the
// provided data and handles jagged arrays (rows with different column counts)
// by using the maximum column count found across all rows.
//
// Thread Safety: This method is thread-safe through mutex locking to prevent
// concurrent modifications to the document structure.
//
// Parameters:
//   - data: 2D array where each inner array represents a table row.
//     Can be jagged (rows with different lengths).
//
// Returns:
//   - *elements.Table: The created and populated table, or nil if:
//   - The document is closed
//   - The data array is empty
//   - All rows are empty (no columns)
//
// Example:
//
//	doc := mbadocx.New()
//
//	// Create a table with product data
//	data := [][]string{
//	    {"Product", "Price", "Stock"},
//	    {"Laptop", "$999", "15"},
//	    {"Mouse", "$25", "50"},
//	    {"Keyboard", "$75", "30"},
//	}
//
//	table := doc.AddTableWithData(data)
//	if table != nil {
//	    table.SetStyle("LightGrid")
//	}
//
//	// Handle jagged data (different column counts)
//	jaggedData := [][]string{
//	    {"Name", "Value"},
//	    {"Temperature", "25°C", "Normal"},  // Extra column
//	    {"Pressure", "1013 mbar"},          // Standard columns
//	}
//	doc.AddTableWithData(jaggedData) // Creates 3-column table
//
// Note: Empty cells are created for rows with fewer columns than the maximum.
// For example, if one row has 5 columns and others have 3, all rows will have
// 5 columns with empty cells where data is missing.
func (d *Document) AddTableWithData(data [][]string) *elements.Table {
	// Acquire lock for thread safety during document modification
	d.mu.Lock()
	defer d.mu.Unlock()

	// Validate preconditions: document must be open and data non-empty
	if d.closed || len(data) == 0 {
		return nil
	}

	// Calculate maximum column count across all rows
	// This handles jagged arrays gracefully
	cols := 0
	for _, row := range data {
		if len(row) > cols {
			cols = len(row)
		}
	}

	// If no columns found (all rows empty), return nil
	if cols == 0 {
		return nil
	}

	// Create table with calculated dimensions
	rows := len(data)
	table := elements.NewTable(d, rows, cols)

	// Populate table cells with provided data
	// Handles rows with fewer columns than maximum
	for i, row := range data {
		for j, cell := range row {
			if j < cols { // Safety check (should always be true)
				_ = table.SetCellText(i, j, cell)
			}
		}
	}

	// Add completed table to document body
	d.body.AddElement(table)
	return table
}

// AddTableWithHeaders creates a table with a formatted header row and data rows.
//
// This method creates a professional-looking table with:
//   - A header row with bold formatting
//   - The header row marked for proper Word table handling
//   - Automatic column count based on headers
//   - Data rows populated from the provided 2D array
//
// Parameters:
//   - headers: Array of column header texts. Determines the number of columns.
//   - data: 2D array of data rows (excluding headers). Each inner array is a row.
//
// Returns:
//   - *elements.Table: The created table with headers and data, or nil if headers is empty
//
// Example:
//
//	doc := mbadocx.New()
//
//	// Create an employee table with headers
//	headers := []string{"ID", "Name", "Position", "Department"}
//	data := [][]string{
//	    {"001", "Alice Johnson", "Manager", "Sales"},
//	    {"002", "Bob Smith", "Developer", "IT"},
//	    {"003", "Carol White", "Analyst", "Finance"},
//	}
//
//	table := doc.AddTableWithHeaders(headers, data)
//
//	// The table now has 4 rows total (1 header + 3 data)
//	// with the header row in bold and marked as a header
//
//	// Apply additional styling
//	table.SetStyle("TableGrid")
//	table.SetAutoFit(true)
//
//	// Handle case where data has fewer columns than headers
//	sparseData := [][]string{
//	    {"001", "John"},  // Missing Position and Department
//	    {"002", "Jane", "Developer"},  // Missing Department
//	}
//	doc.AddTableWithHeaders(headers, sparseData) // Empty cells filled automatically
//
// Note:
//   - The header row is automatically formatted in bold for visual distinction
//   - SetHeaderRow(0) ensures the header repeats on new pages for long tables
//   - If data rows have fewer columns than headers, empty cells are created
//   - If data rows have more columns than headers, extra columns are ignored
func (d *Document) AddTableWithHeaders(headers []string, data [][]string) *elements.Table {
	// Validate that headers exist (required for column structure)
	if len(headers) == 0 {
		return nil
	}

	// Calculate dimensions: +1 row for headers
	rows := len(data) + 1
	cols := len(headers)

	// Create table with calculated dimensions
	table := elements.NewTable(d, rows, cols)

	// Populate header row with bold formatting
	// Using a callback function to apply formatting to each header cell
	for i, header := range headers {
		_ = table.SetCellFormattedText(0, i, header, func(r *elements.Run) {
			r.SetBold(true) // Apply bold formatting to header text
		})
	}

	// Mark first row as header for proper Word table behavior
	// This ensures headers repeat on new pages and are recognized by screen readers
	_ = table.SetHeaderRow(0)

	// Populate data rows starting from row index 1 (after headers)
	for i, row := range data {
		for j, cellData := range row {
			// Only set cells within column bounds (ignore extra columns)
			if j < cols {
				_ = table.SetCellText(i+1, j, cellData) // i+1 because row 0 is headers
			}
		}
		// Note: If row has fewer columns than headers, cells remain empty
	}

	// Add completed table to document body
	d.body.AddElement(table)
	return table
}
