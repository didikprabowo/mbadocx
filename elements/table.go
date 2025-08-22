// elements/table.go
package elements

import (
	"bytes"
	"fmt"

	"github.com/didikprabowo/mbadocx/types"
)

// Vertical alignment options
type VerticalAlign string

const (
	AlignTop    VerticalAlign = "top"
	AlignCenter VerticalAlign = "center"
	AlignBottom VerticalAlign = "bottom"
)

// Horizontal alignment options
type TableAlign string

const (
	AlignLeft    TableAlign = "left" // not shown in Word UI, but valid
	AlignCenterH TableAlign = "center"
	AlignRight   TableAlign = "right"
	AlignJustify TableAlign = "both"
)

// Table represents a table element in a Word document
type Table struct {
	document   types.Document
	Properties *TableProperties
	Grid       *TableGrid
	Rows       []*TableRow
}

// TableProperties represents table properties
type TableProperties struct {
	Width       *TableWidth
	Alignment   *TableAlignment
	Borders     *TableBorders
	CellMargin  *TableCellMargin
	Style       *TableStyle
	Look        *TableLook
	CellSpacing *TableCellSpacing
	Indent      *TableIndent
}

// TableWidth represents table width
type TableWidth struct {
	Type  string // auto, dxa, pct
	Value string
}

// TableAlignment represents table alignment
type TableAlignment struct {
	Value TableAlign // left, center, right
}

// TableStyle represents table style reference
type TableStyle struct {
	Value string
}

// TableLook represents table look properties
type TableLook struct {
	FirstRow    string
	LastRow     string
	FirstColumn string
	LastColumn  string
	NoHBand     string
	NoVBand     string
}

// TableGrid represents table grid
type TableGrid struct {
	Columns []*TableGridCol
}

// TableGridCol represents a table grid column
type TableGridCol struct {
	Width string
}

// TableRow represents a table row
type TableRow struct {
	Properties *TableRowProperties
	Cells      []*TableCell
}

// TableRowProperties represents table row properties
type TableRowProperties struct {
	Height      *TableRowHeight
	CantSplit   bool
	TableHeader bool
}

// TableRowHeight represents row height
type TableRowHeight struct {
	Value string
	Rule  string // auto, exact, atLeast
}

// TableCell represents a table cell
type TableCell struct {
	Properties *TableCellProperties
	Paragraphs []*Paragraph
}

// TableCellProperties represents table cell properties
type TableCellProperties struct {
	Width         *TableCellWidth
	GridSpan      int
	VerticalMerge *VerticalMerge
	Borders       *TableCellBorders
	Shading       *TableCellShading
	Margins       *TableCellMargins
	VerticalAlign VerticalAlign // top, center, bottom
}

// TableCellWidth represents cell width
type TableCellWidth struct {
	Type  string // auto, dxa, pct
	Value string
}

// VerticalMerge represents vertical cell merge
type VerticalMerge struct {
	Value string // restart, continue
}

// TableBorders represents table borders
type TableBorders struct {
	Top     *BorderStyle
	Left    *BorderStyle
	Bottom  *BorderStyle
	Right   *BorderStyle
	InsideH *BorderStyle
	InsideV *BorderStyle
}

// TableCellBorders represents cell borders
type TableCellBorders struct {
	Top    *BorderStyle
	Left   *BorderStyle
	Bottom *BorderStyle
	Right  *BorderStyle
}

// BorderStyle represents a border style
type BorderStyle struct {
	Value string // single, double, dotted, dashed, thick, none
	Size  string
	Space string
	Color string
}

// TableCellMargin represents table cell margins
type TableCellMargin struct {
	Top    *MarginValue
	Left   *MarginValue
	Bottom *MarginValue
	Right  *MarginValue
}

// TableCellMargins represents individual cell margins
type TableCellMargins struct {
	Top    *MarginValue
	Left   *MarginValue
	Bottom *MarginValue
	Right  *MarginValue
}

// MarginValue represents a margin value
type MarginValue struct {
	Width string
	Type  string // dxa, auto
}

// TableCellSpacing represents cell spacing
type TableCellSpacing struct {
	Width string
	Type  string
}

// TableIndent represents table indentation
type TableIndent struct {
	Width string
	Type  string
}

// TableCellShading represents cell shading
type TableCellShading struct {
	Value string // clear, solid
	Color string
	Fill  string
}

// NewTable creates a new table with specified rows and columns
func NewTable(document types.Document, rows, cols int) *Table {
	table := &Table{
		document: document,
		Properties: &TableProperties{
			Alignment: &TableAlignment{
				Value: "both",
			},
			Indent: &TableIndent{
				Width: "0",
				Type:  "dxa",
			},
			Width: &TableWidth{
				Type:  "auto",
				Value: "0",
			},
			Borders: DefaultTableBorders(),
			Look: &TableLook{
				FirstRow:    "1",
				LastRow:     "0",
				FirstColumn: "1",
				LastColumn:  "0",
				NoHBand:     "0",
				NoVBand:     "1",
			},
			CellMargin: &TableCellMargin{ // default Word padding
				Top:    &MarginValue{Width: "0", Type: "dxa"},
				Bottom: &MarginValue{Width: "0", Type: "dxa"},
				Left:   &MarginValue{Width: "80", Type: "dxa"},
				Right:  &MarginValue{Width: "80", Type: "dxa"},
			},
		},
		Grid: &TableGrid{
			Columns: make([]*TableGridCol, cols),
		},
		Rows: make([]*TableRow, rows),
	}

	// Initialize grid columns
	for i := 0; i < cols; i++ {
		table.Grid.Columns[i] = &TableGridCol{
			Width: "2880", // Default width (2 inches in twips)
		}
	}

	// Initialize rows and cells
	for i := 0; i < rows; i++ {
		row := &TableRow{
			Cells: make([]*TableCell, cols),
			Properties: &TableRowProperties{
				Height: &TableRowHeight{
					Value: "auto", // 300 twips = 15pt
					Rule:  "atLeast",
				},
			},
		}
		for j := 0; j < cols; j++ {
			row.Cells[j] = &TableCell{
				Properties: &TableCellProperties{
					VerticalAlign: "center",
					Width: &TableCellWidth{
						Type:  "dxa",
						Value: "2880",
					},
				},
				Paragraphs: []*Paragraph{
					NewParagraph(document), // Add empty paragraph to each cell
				},
			}
		}
		table.Rows[i] = row
	}

	return table
}

// Type returns the element type
func (t *Table) Type() string {
	return "table"
}

// DefaultTableBorders returns default table borders
func DefaultTableBorders() *TableBorders {
	return &TableBorders{
		Top:     &BorderStyle{Value: "single", Size: "4", Space: "0", Color: "auto"},
		Left:    &BorderStyle{Value: "single", Size: "4", Space: "0", Color: "auto"},
		Bottom:  &BorderStyle{Value: "single", Size: "4", Space: "0", Color: "auto"},
		Right:   &BorderStyle{Value: "single", Size: "4", Space: "0", Color: "auto"},
		InsideH: &BorderStyle{Value: "single", Size: "4", Space: "0", Color: "auto"},
		InsideV: &BorderStyle{Value: "single", Size: "4", Space: "0", Color: "auto"},
	}
}

// SetCellText sets text in a specific cell
func (t *Table) SetCellText(row, col int, text string) error {
	if row >= len(t.Rows) || col >= len(t.Rows[row].Cells) {
		return fmt.Errorf("cell position out of bounds")
	}

	cell := t.Rows[row].Cells[col]
	if len(cell.Paragraphs) == 0 {
		cell.Paragraphs = []*Paragraph{NewParagraph(t.document)}
	}

	// Clear existing content and add new text
	cell.Paragraphs[0].Clear()
	// Clear existing content and add formatted text
	cell.Paragraphs[0].Clear()
	cell.Paragraphs[0].Properties.SpacingBefore = 0
	cell.Paragraphs[0].Properties.SpacingAfter = 0
	cell.Paragraphs[0].Properties.LineSpacing = 1.15 // Default Word line spacing
	cell.Paragraphs[0].Properties.LineSpacingRule = "auto"

	cell.Paragraphs[0].AddText(text)

	return nil
}

// SetCellFormattedText sets formatted text in a specific cell
func (t *Table) SetCellFormattedText(row, col int, text string, format func(*Run)) error {
	if row >= len(t.Rows) || col >= len(t.Rows[row].Cells) {
		return fmt.Errorf("cell position out of bounds")
	}

	cell := t.Rows[row].Cells[col]
	if len(cell.Paragraphs) == 0 {
		cell.Paragraphs = []*Paragraph{NewTableCellParagraph(t.document)}
	}

	// Clear existing content and add formatted text
	cell.Paragraphs[0].Clear()
	cell.Paragraphs[0].Properties.SpacingBefore = 0
	cell.Paragraphs[0].Properties.SpacingAfter = 0
	cell.Paragraphs[0].Properties.LineSpacing = 1.15 // Default Word line spacing
	cell.Paragraphs[0].Properties.LineSpacingRule = "auto"

	cell.Paragraphs[0].AddFormattedText(text, format)

	return nil
}

// AddRow adds a new row to the table
func (t *Table) AddRow() *TableRow {
	cols := len(t.Grid.Columns)
	row := &TableRow{
		Cells: make([]*TableCell, cols),
	}

	for i := 0; i < cols; i++ {
		row.Cells[i] = &TableCell{
			Properties: &TableCellProperties{
				Width: &TableCellWidth{
					Type:  "dxa",
					Value: t.Grid.Columns[i].Width,
				},
			},
			Paragraphs: []*Paragraph{
				NewTableCellParagraph(t.document),
			},
		}
	}

	t.Rows = append(t.Rows, row)
	return row
}

// SetColumnWidth sets the width of a specific column
func (t *Table) SetColumnWidth(col int, width string) error {
	if col >= len(t.Grid.Columns) {
		return fmt.Errorf("column index out of bounds")
	}

	t.Grid.Columns[col].Width = width

	// Update cell widths in all rows
	for _, row := range t.Rows {
		if col < len(row.Cells) {
			if row.Cells[col].Properties == nil {
				row.Cells[col].Properties = &TableCellProperties{}
			}
			if row.Cells[col].Properties.Width == nil {
				row.Cells[col].Properties.Width = &TableCellWidth{Type: "dxa"}
			}
			row.Cells[col].Properties.Width.Value = width
		}
	}

	return nil
}

// SetTableWidth sets the overall table width
func (t *Table) SetTableWidth(widthType, value string) {
	if t.Properties == nil {
		t.Properties = &TableProperties{}
	}
	if t.Properties.Width == nil {
		t.Properties.Width = &TableWidth{}
	}
	t.Properties.Width.Type = widthType
	t.Properties.Width.Value = value
}

// SetTableAlignment sets table alignment (left, center, right)
func (t *Table) SetTableAlignment(alignment TableAlign) {
	if t.Properties == nil {
		t.Properties = &TableProperties{}
	}
	t.Properties.Alignment = &TableAlignment{
		Value: alignment,
	}
}

// MergeCells merges cells horizontally
func (t *Table) MergeCells(row, startCol, endCol int) error {
	if row >= len(t.Rows) || startCol >= len(t.Rows[row].Cells) || endCol >= len(t.Rows[row].Cells) {
		return fmt.Errorf("merge position out of bounds")
	}

	span := endCol - startCol + 1
	if t.Rows[row].Cells[startCol].Properties == nil {
		t.Rows[row].Cells[startCol].Properties = &TableCellProperties{}
	}
	t.Rows[row].Cells[startCol].Properties.GridSpan = span

	// Remove the merged cells
	t.Rows[row].Cells = append(t.Rows[row].Cells[:startCol+1], t.Rows[row].Cells[endCol+1:]...)

	return nil
}

// SetCellShading sets background color for a cell
func (t *Table) SetCellShading(row, col int, color string) error {
	if row >= len(t.Rows) || col >= len(t.Rows[row].Cells) {
		return fmt.Errorf("cell position out of bounds")
	}

	cell := t.Rows[row].Cells[col]
	if cell.Properties == nil {
		cell.Properties = &TableCellProperties{}
	}

	cell.Properties.Shading = &TableCellShading{
		Value: "clear",
		Color: "auto",
		Fill:  color,
	}

	return nil
}

// SetCellVerticalAlignment sets vertical alignment for a cell
func (t *Table) SetCellVerticalAlignment(row, col int, alignment VerticalAlign) error {
	if row >= len(t.Rows) || col >= len(t.Rows[row].Cells) {
		return fmt.Errorf("cell position out of bounds")
	}

	cell := t.Rows[row].Cells[col]
	if cell.Properties == nil {
		cell.Properties = &TableCellProperties{}
	}

	cell.Properties.VerticalAlign = alignment

	return nil
}

// SetRowHeight sets the height of a specific row
func (t *Table) SetRowHeight(row int, height string, rule string) error {
	if row >= len(t.Rows) {
		return fmt.Errorf("row index out of bounds")
	}

	if t.Rows[row].Properties == nil {
		t.Rows[row].Properties = &TableRowProperties{}
	}

	t.Rows[row].Properties.Height = &TableRowHeight{
		Value: height,
		Rule:  rule,
	}

	return nil
}

// SetHeaderRow marks a row as a header row
func (t *Table) SetHeaderRow(row int) error {
	if row >= len(t.Rows) {
		return fmt.Errorf("row index out of bounds")
	}

	if t.Rows[row].Properties == nil {
		t.Rows[row].Properties = &TableRowProperties{}
	}

	t.Rows[row].Properties.TableHeader = true

	return nil
}

// XML generates the XML representation of the table
func (t *Table) XML() ([]byte, error) {
	var buf bytes.Buffer

	buf.WriteString(`<w:tbl>`)

	// Table properties
	if t.Properties != nil {
		propXML, err := t.generatePropertiesXML()
		if err != nil {
			return nil, fmt.Errorf("generating table properties: %w", err)
		}
		buf.Write(propXML)
	}

	// Table grid
	if t.Grid != nil {
		gridXML, err := t.generateGridXML()
		if err != nil {
			return nil, fmt.Errorf("generating table grid: %w", err)
		}
		buf.Write(gridXML)
	}

	// Table rows
	for _, row := range t.Rows {
		rowXML, err := t.generateRowXML(row)
		if err != nil {
			return nil, fmt.Errorf("generating table row: %w", err)
		}
		buf.Write(rowXML)
	}

	buf.WriteString(`</w:tbl>`)

	return buf.Bytes(), nil
}

// generatePropertiesXML generates the table properties XML
func (t *Table) generatePropertiesXML() ([]byte, error) {
	var buf bytes.Buffer
	buf.WriteString(`<w:tblPr>`)

	if t.Properties.Indent != nil {
		buf.WriteString(fmt.Sprintf(`<w:tblInd w:w="%s" w:type="%s"/>`, t.Properties.Indent.Width, t.Properties.Indent.Type))
	}

	// Table style
	if t.Properties.Style != nil {
		buf.WriteString(fmt.Sprintf(`<w:tblStyle w:val="%s"/>`, t.Properties.Style.Value))
	}

	// Table width
	if t.Properties.Width != nil {
		buf.WriteString(fmt.Sprintf(`<w:tblW w:type="%s" w:w="%s"/>`,
			t.Properties.Width.Type, t.Properties.Width.Value))
	}

	// Table alignment
	if t.Properties.Alignment != nil {
		buf.WriteString(fmt.Sprintf(`<w:jc w:val="%s"/>`, t.Properties.Alignment.Value))
	}

	// Table borders
	if t.Properties.Borders != nil {
		bordersXML, err := t.generateBordersXML(t.Properties.Borders)
		if err != nil {
			return nil, err
		}
		buf.Write(bordersXML)
	}

	// CellMargin
	if t.Properties.CellMargin != nil {
		buf.WriteString("<w:tblCellMar>")
		if t.Properties.CellMargin.Top != nil {
			buf.WriteString(fmt.Sprintf(`<w:top w:w="%s" w:type="%s"/>`,
				t.Properties.CellMargin.Top.Width, t.Properties.CellMargin.Top.Type))
		}
		if t.Properties.CellMargin.Left != nil {
			buf.WriteString(fmt.Sprintf(`<w:left w:w="%s" w:type="%s"/>`,
				t.Properties.CellMargin.Left.Width, t.Properties.CellMargin.Left.Type))
		}
		if t.Properties.CellMargin.Bottom != nil {
			buf.WriteString(fmt.Sprintf(`<w:bottom w:w="%s" w:type="%s"/>`,
				t.Properties.CellMargin.Bottom.Width, t.Properties.CellMargin.Bottom.Type))
		}
		if t.Properties.CellMargin.Right != nil {
			buf.WriteString(fmt.Sprintf(`<w:right w:w="%s" w:type="%s"/>`,
				t.Properties.CellMargin.Right.Width, t.Properties.CellMargin.Right.Type))
		}
		buf.WriteString("</w:tblCellMar>")
	}

	// Table look
	if t.Properties.Look != nil {
		buf.WriteString(fmt.Sprintf(`<w:tblLook w:firstRow="%s" w:lastRow="%s" w:firstColumn="%s" w:lastColumn="%s" w:noHBand="%s" w:noVBand="%s"/>`,
			t.Properties.Look.FirstRow, t.Properties.Look.LastRow,
			t.Properties.Look.FirstColumn, t.Properties.Look.LastColumn,
			t.Properties.Look.NoHBand, t.Properties.Look.NoVBand))
	}

	buf.WriteString(`</w:tblPr>`)
	return buf.Bytes(), nil
}

// generateGridXML generates the table grid XML
func (t *Table) generateGridXML() ([]byte, error) {
	var buf bytes.Buffer
	buf.WriteString(`<w:tblGrid>`)

	for _, col := range t.Grid.Columns {
		buf.WriteString(fmt.Sprintf(`<w:gridCol w:w="%s"/>`, col.Width))
	}

	buf.WriteString(`</w:tblGrid>`)
	return buf.Bytes(), nil
}

// generateRowXML generates a table row XML
func (t *Table) generateRowXML(row *TableRow) ([]byte, error) {
	var buf bytes.Buffer
	buf.WriteString(`<w:tr>`)

	// Row properties
	if row.Properties != nil {
		propXML, err := t.generateRowPropertiesXML(row.Properties)
		if err != nil {
			return nil, err
		}
		buf.Write(propXML)
	}

	// Row cells
	for _, cell := range row.Cells {
		cellXML, err := t.generateCellXML(cell)
		if err != nil {
			return nil, err
		}
		buf.Write(cellXML)
	}

	buf.WriteString(`</w:tr>`)
	return buf.Bytes(), nil
}

// generateRowPropertiesXML generates row properties XML
func (t *Table) generateRowPropertiesXML(props *TableRowProperties) ([]byte, error) {
	var buf bytes.Buffer
	buf.WriteString(`<w:trPr>`)

	if props.Height != nil {
		buf.WriteString(fmt.Sprintf(`<w:trHeight w:val="%s"`, props.Height.Value))
		if props.Height.Rule == "exact" || props.Height.Rule == "atLeast" {
			buf.WriteString(fmt.Sprintf(` w:hRule="%s"`, props.Height.Rule))
		}
		buf.WriteString(`/>`)
	}

	if props.CantSplit {
		buf.WriteString(`<w:cantSplit/>`)
	}

	if props.TableHeader {
		buf.WriteString(`<w:tblHeader/>`)
	}

	buf.WriteString(`</w:trPr>`)
	return buf.Bytes(), nil
}

// generateCellXML generates a table cell XML
func (t *Table) generateCellXML(cell *TableCell) ([]byte, error) {
	var buf bytes.Buffer
	buf.WriteString(`<w:tc>`)

	// Cell properties
	if cell.Properties != nil {
		propXML, err := t.generateCellPropertiesXML(cell.Properties)
		if err != nil {
			return nil, err
		}
		buf.Write(propXML)
	}

	// Cell paragraphs
	for _, para := range cell.Paragraphs {
		paraXML, err := para.XML()
		if err != nil {
			return nil, err
		}
		buf.Write(paraXML)
	}

	buf.WriteString(`</w:tc>`)
	return buf.Bytes(), nil
}

// generateCellPropertiesXML generates cell properties XML
func (t *Table) generateCellPropertiesXML(props *TableCellProperties) ([]byte, error) {
	var buf bytes.Buffer
	buf.WriteString(`<w:tcPr>`)

	// Cell width
	if props.Width != nil {
		buf.WriteString(fmt.Sprintf(`<w:tcW w:type="%s" w:w="%s"/>`,
			props.Width.Type, props.Width.Value))
	}

	// Grid span
	if props.GridSpan > 1 {
		buf.WriteString(fmt.Sprintf(`<w:gridSpan w:val="%d"/>`, props.GridSpan))
	}

	// Vertical merge
	if props.VerticalMerge != nil {
		buf.WriteString(`<w:vMerge`)
		if props.VerticalMerge.Value != "" {
			buf.WriteString(fmt.Sprintf(` w:val="%s"`, props.VerticalMerge.Value))
		}
		buf.WriteString(`/>`)
	}

	// Vertical alignment
	// Vertical alignment - FIX: use "center" or "top", not "left"
	if props.VerticalAlign != "" {
		// Convert "left" to "top" for proper alignment
		valign := props.VerticalAlign
		if valign == "left" {
			valign = "top"
		}
		buf.WriteString(fmt.Sprintf(`<w:vAlign w:val="%s"/>`, valign))
	}

	// Cell shading
	if props.Shading != nil {
		buf.WriteString(`<w:shd`)
		if props.Shading.Value != "" {
			buf.WriteString(fmt.Sprintf(` w:val="%s"`, props.Shading.Value))
		}
		if props.Shading.Color != "" {
			buf.WriteString(fmt.Sprintf(` w:color="%s"`, props.Shading.Color))
		}
		if props.Shading.Fill != "" {
			buf.WriteString(fmt.Sprintf(` w:fill="%s"`, props.Shading.Fill))
		}
		buf.WriteString(`/>`)
	}

	// Cell margins (padding)
	if props.Margins != nil {
		buf.WriteString(`<w:tcMar>`)
		if props.Margins.Top != nil {
			buf.WriteString(fmt.Sprintf(`<w:top w:w="%s" w:type="%s"/>`,
				props.Margins.Top.Width, props.Margins.Top.Type))
		}
		if props.Margins.Left != nil {
			buf.WriteString(fmt.Sprintf(`<w:left w:w="%s" w:type="%s"/>`,
				props.Margins.Left.Width, props.Margins.Left.Type))
		}
		if props.Margins.Bottom != nil {
			buf.WriteString(fmt.Sprintf(`<w:bottom w:w="%s" w:type="%s"/>`,
				props.Margins.Bottom.Width, props.Margins.Bottom.Type))
		}
		if props.Margins.Right != nil {
			buf.WriteString(fmt.Sprintf(`<w:right w:w="%s" w:type="%s"/>`,
				props.Margins.Right.Width, props.Margins.Right.Type))
		}
		buf.WriteString(`</w:tcMar>`)
	}

	buf.WriteString(`</w:tcPr>`)
	return buf.Bytes(), nil
}

// generateBordersXML generates borders XML
func (t *Table) generateBordersXML(borders *TableBorders) ([]byte, error) {
	var buf bytes.Buffer
	buf.WriteString(`<w:tblBorders>`)

	if borders.Top != nil {
		buf.Write(t.generateBorderXML("top", borders.Top))
	}
	if borders.Left != nil {
		buf.Write(t.generateBorderXML("left", borders.Left))
	}
	if borders.Bottom != nil {
		buf.Write(t.generateBorderXML("bottom", borders.Bottom))
	}
	if borders.Right != nil {
		buf.Write(t.generateBorderXML("right", borders.Right))
	}
	if borders.InsideH != nil {
		buf.Write(t.generateBorderXML("insideH", borders.InsideH))
	}
	if borders.InsideV != nil {
		buf.Write(t.generateBorderXML("insideV", borders.InsideV))
	}

	buf.WriteString(`</w:tblBorders>`)
	return buf.Bytes(), nil
}

// generateBorderXML generates a single border XML
func (t *Table) generateBorderXML(position string, border *BorderStyle) []byte {
	var buf bytes.Buffer
	buf.WriteString(fmt.Sprintf(`<w:%s w:val="%s"`, position, border.Value))

	if border.Size != "" {
		buf.WriteString(fmt.Sprintf(` w:sz="%s"`, border.Size))
	}
	if border.Space != "" {
		buf.WriteString(fmt.Sprintf(` w:space="%s"`, border.Space))
	}
	if border.Color != "" {
		buf.WriteString(fmt.Sprintf(` w:color="%s"`, border.Color))
	}

	buf.WriteString(`/>`)
	return buf.Bytes()
}

func NewTableCellParagraph(document types.Document) *Paragraph {
	p := NewParagraph(document)
	// Set explicit zero spacing for table cells
	p.Properties.SpacingBefore = 0
	p.Properties.SpacingAfter = 0
	p.Properties.LineSpacing = 1.0 // Default Word line spacing
	p.Properties.LineSpacingRule = "auto"
	return p
}
