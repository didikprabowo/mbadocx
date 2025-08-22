package mbadocx

import "github.com/didikprabowo/mbadocx/elements"

// AddTable adds a table to the document
func (d *Document) AddTable(rows, cols int) *elements.Table {
	tableElem := elements.NewTable(d, rows, cols)
	d.body.AddElement(tableElem)
	return tableElem
}

// AddTableWithData adds a table with initial data
func (d *Document) AddTableWithData(data [][]string) *elements.Table {
	d.mu.Lock()
	defer d.mu.Unlock()

	if d.closed || len(data) == 0 {
		return nil
	}

	// Find max columns
	cols := 0
	for _, row := range data {
		if len(row) > cols {
			cols = len(row)
		}
	}

	if cols == 0 {
		return nil
	}

	rows := len(data)
	table := elements.NewTable(d, rows, cols)

	// Fill table with data
	for i, row := range data {
		for j, cell := range row {
			if j < cols {
				table.SetCellText(i, j, cell)
			}
		}
	}

	d.body.AddElement(table)
	return table
}

// AddTableWithHeaders creates a table with header row
func (d *Document) AddTableWithHeaders(headers []string, data [][]string) *elements.Table {
	if len(headers) == 0 {
		return nil
	}

	rows := len(data) + 1
	cols := len(headers)
	table := elements.NewTable(d, rows, cols)

	// Set headers with bold formatting
	for i, header := range headers {
		table.SetCellFormattedText(0, i, header, func(r *elements.Run) {
			r.SetBold(true)
		})
	}

	// Mark first row as header
	table.SetHeaderRow(0)

	// Add data rows
	for i, row := range data {
		for j, cellData := range row {
			if j < cols {
				table.SetCellText(i+1, j, cellData)
			}
		}
	}

	d.body.AddElement(table)
	return table
}
