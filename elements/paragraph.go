// File: elements/paragraph.go
package elements

import (
	"bytes"
	"fmt"

	"github.com/didikprabowo/mbadocx/properties"
	"github.com/didikprabowo/mbadocx/types"
)

type ListType string

const (
	ListTypeBullet  ListType = "Bullet"
	ListTypeDecimal ListType = "Decimal numbering"
	ListTypeLegal   ListType = "Legal numbering"
	ListTypeRoman   ListType = "Roman numerals"
	ListTypeCustom  ListType = "Custom symbols"
)

// Paragraph represents a paragraph element
type Paragraph struct {
	document   types.Document
	Properties *properties.ParagraphProperties
	Children   []ParagraphChild
}

// ParagraphChild interface for elements that can be children of a paragraph
type ParagraphChild interface {
	Type() string
	XML() ([]byte, error)
}

// NewParagraph creates a new paragraph
func NewParagraph(document types.Document) *Paragraph {
	return &Paragraph{
		document:   document,
		Properties: properties.NewParagraphProperties(),
		Children:   make([]ParagraphChild, 0),
	}
}

// Type returns the element type
func (p *Paragraph) Type() string {
	return "paragraph"
}

// AddRun adds a new run to the paragraph
func (p *Paragraph) AddRun() *Run {
	r := NewRun()
	p.Children = append(p.Children, r)
	return r
}

// AddText is a convenience method to add a text run
func (p *Paragraph) AddText(text string) *Run {
	r := p.AddRun()
	r.AddText(text)
	return r
}

// AddHyperlink
func (pb *Paragraph) AddHyperlink(text, url string) *Paragraph {
	h := NewHyperlink(text, url)

	if h.Typ == HyperlinkTypeExternal {
		rel := pb.document.GetRelationships().GetOrCreateHyperlink(url)
		h.ID = rel.ID
	}

	pb.Children = append(pb.Children, h)
	return pb
}

// AddFormattedText adds text with specific formatting
func (p *Paragraph) AddFormattedText(text string, format func(*Run)) *Run {
	r := p.AddRun()
	r.AddText(text)
	if format != nil {
		format(r)
	}
	return r
}

// AddLineBreak adds a paragraph
func (p *Paragraph) AddLineBreak() *Paragraph {
	run := p.AddRun()
	run.AddBreak()
	return p
}

// AddPageBreak adds a page break to the paragraph
func (p *Paragraph) AddPageBreak() *Paragraph {
	run := p.AddRun()
	run.AddPageBreak()
	return p
}

// SetAlignment sets the paragraph alignment
func (p *Paragraph) SetAlignment(alignment string) *Paragraph {
	// Map "justify" to "both" for DOCX compatibility
	if alignment == "justify" {
		alignment = "both"
	}
	p.Properties.Alignment = alignment
	return p
}

// SetStyle sets the paragraph style
func (p *Paragraph) SetStyle(styleID string) *Paragraph {
	p.Properties.StyleID = styleID
	return p
}

// SetSpacing sets spacing before and after the paragraph
func (p *Paragraph) SetSpacing(before, after float64) *Paragraph {
	p.Properties.SpacingBefore = before
	p.Properties.SpacingAfter = after
	return p
}

// SetLineSpacing sets the line spacing
func (p *Paragraph) SetLineSpacing(spacing float64, rule string) *Paragraph {
	p.Properties.LineSpacing = spacing
	p.Properties.LineSpacingRule = rule
	return p
}

// SetIndentation sets paragraph indentation
func (p *Paragraph) SetIndentation(left, right, firstLine float64) *Paragraph {
	p.Properties.IndentLeft = left
	p.Properties.IndentRight = right
	p.Properties.IndentFirstLine = firstLine
	return p
}

// SetHangingIndent sets a hanging indent
func (p *Paragraph) SetHangingIndent(hanging float64) *Paragraph {
	p.Properties.IndentFirstLine = -hanging
	return p
}

// SetKeepNext sets whether to keep this paragraph with the next one
func (p *Paragraph) SetKeepNext(keep bool) *Paragraph {
	p.Properties.KeepNext = keep
	return p
}

// SetKeepLines sets whether to keep all lines together
func (p *Paragraph) SetKeepLines(keep bool) *Paragraph {
	p.Properties.KeepLines = keep
	return p
}

// SetPageBreakBefore sets whether to insert a page break before
func (p *Paragraph) SetPageBreakBefore(pageBreak bool) *Paragraph {
	p.Properties.PageBreakBefore = pageBreak
	return p
}

// SetWidowControl sets widow/orphan control
func (p *Paragraph) SetWidowControl(control bool) *Paragraph {
	p.Properties.WidowControl = control
	return p
}

// SetNumbering sets numbering properties
//
//	ID: 1 -> Bullet list
//	ID: 2 -> Decimal numbering
//	ID: 3 -> Legal numbering
//	ID: 4 -> Roman numerals
//	ID: 5 -> Custom symbols
func (p *Paragraph) SetNumbering(listType ListType, level int) *Paragraph {
	var numID string
	switch listType {
	case ListTypeBullet:
		numID = "1"
	case ListTypeDecimal:
		numID = "2"
	case ListTypeLegal:
		numID = "3"
	case ListTypeRoman:
		numID = "4"
	case ListTypeCustom:
		numID = "5"
	default:
		numID = "1"
	}
	p.Properties.NumberingID = numID
	p.Properties.NumberingLevel = level
	return p
}

// SetOutlineLevel sets the outline level for TOC
func (p *Paragraph) SetOutlineLevel(level int) *Paragraph {
	p.Properties.OutlineLevel = level
	return p
}

// SetBorders sets paragraph borders
func (p *Paragraph) SetBorders(borders *properties.ParagraphBorders) *Paragraph {
	p.Properties.Borders = borders
	return p
}

// SetShading sets paragraph shading
func (p *Paragraph) SetShading(shading *properties.ParagraphShading) *Paragraph {
	p.Properties.Shading = shading
	return p
}

// SetTabs sets custom tab stops
func (p *Paragraph) SetTabs(tabs []properties.TabStop) *Paragraph {
	p.Properties.Tabs = tabs
	return p
}

// Clone creates a deep copy of the paragraph
func (p *Paragraph) Clone() *Paragraph {
	newPara := &Paragraph{
		Properties: p.Properties.Clone(),
		Children:   make([]ParagraphChild, 0, len(p.Children)),
	}

	// Clone children
	for _, child := range p.Children {
		switch c := child.(type) {
		case *Run:
			newPara.Children = append(newPara.Children, c.Clone())
		case *Hyperlink:
			newPara.Children = append(newPara.Children, c.Clone())
			// Add other child types as needed
		}
	}

	return newPara
}

// Clear removes all content from the paragraph
func (p *Paragraph) Clear() {
	p.Children = p.Children[:0]
}

// Validate checks if the paragraph is valid
func (p *Paragraph) Validate() error {
	if p.Properties != nil {
		if err := p.Properties.Validate(); err != nil {
			return fmt.Errorf("invalid paragraph properties: %w", err)
		}
	}

	for i, child := range p.Children {
		// Validate each child if it has a Validate method
		if validator, ok := child.(interface{ Validate() error }); ok {
			if err := validator.Validate(); err != nil {
				return fmt.Errorf("invalid child at index %d: %w", i, err)
			}
		}
	}

	return nil
}

// XML generates the XML representation of the paragraph
func (p *Paragraph) XML() ([]byte, error) {
	var buf bytes.Buffer

	// Start with XML declaration
	buf.WriteString(`<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"`)

	// Add namespace for hyperlinks if needed
	hasHyperlinks := false
	for _, child := range p.Children {
		if _, ok := child.(*Hyperlink); ok {
			hasHyperlinks = true
			break
		}
	}

	if hasHyperlinks {
		buf.WriteString(` xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"`)
	}

	buf.WriteString(`>`)

	// Add properties if they exist
	if p.Properties != nil && !p.Properties.IsEmpty() {
		propXML, err := p.generatePropertiesXML()
		if err != nil {
			return nil, fmt.Errorf("generating properties XML: %w", err)
		}
		buf.Write(propXML)
	}

	// Add children (runs, hyperlinks, etc.)
	for _, child := range p.Children {
		childXML, err := child.XML()
		if err != nil {
			return nil, fmt.Errorf("generating child XML: %w", err)
		}
		buf.Write(childXML)
	}

	// Close paragraph tag
	buf.WriteString(`</w:p>`)

	return buf.Bytes(), nil
}

// generatePropertiesXML generates the properties XML
func (p *Paragraph) generatePropertiesXML() ([]byte, error) {
	pp := p.Properties
	if pp == nil {
		return nil, nil
	}

	var buf bytes.Buffer
	buf.WriteString(`<w:pPr>`)

	// Style
	if pp.StyleID != "" {
		buf.WriteString(fmt.Sprintf(`<w:pStyle w:val="%s"/>`, pp.StyleID))
	}

	// Keep properties
	if pp.KeepNext {
		buf.WriteString(`<w:keepNext/>`)
	}
	if pp.KeepLines {
		buf.WriteString(`<w:keepLines/>`)
	}
	if pp.PageBreakBefore {
		buf.WriteString(`<w:pageBreakBefore/>`)
	}

	// Widow control
	if !pp.WidowControl {
		buf.WriteString(`<w:widowControl w:val="false"/>`)
	}

	// Numbering
	if pp.NumberingID != "" {
		buf.WriteString(`<w:numPr>`)
		buf.WriteString(fmt.Sprintf(`<w:ilvl w:val="%d"/>`, pp.NumberingLevel))
		buf.WriteString(fmt.Sprintf(`<w:numId w:val="%s"/>`, pp.NumberingID))
		buf.WriteString(`</w:numPr>`)
	}

	// Borders
	if pp.Borders != nil {
		bordersXML, err := pp.Borders.XML()
		if err != nil {
			return nil, err
		}
		buf.Write(bordersXML)
	}

	// Shading
	if pp.Shading != nil {
		shadingXML, err := pp.Shading.XML()
		if err != nil {
			return nil, err
		}
		buf.Write(shadingXML)
	}

	// Tabs
	if len(pp.Tabs) > 0 {
		buf.WriteString(`<w:tabs>`)
		for _, tab := range pp.Tabs {
			buf.WriteString(fmt.Sprintf(`<w:tab w:val="%s" w:pos="%d"`, tab.Alignment, tab.Position))
			if tab.Leader != "" {
				buf.WriteString(fmt.Sprintf(` w:leader="%s"`, tab.Leader))
			}
			buf.WriteString(`/>`)
		}
		buf.WriteString(`</w:tabs>`)
	}

	// Suppress auto hyphens
	if pp.SuppressAutoHyphens {
		buf.WriteString(`<w:suppressAutoHyphens/>`)
	}

	// Alignment
	if pp.Alignment != "" && pp.Alignment != "left" {
		buf.WriteString(fmt.Sprintf(`<w:jc w:val="%s"/>`, pp.Alignment))
	}

	// Outline level
	if pp.OutlineLevel > 0 {
		buf.WriteString(fmt.Sprintf(`<w:outlineLvl w:val="%d"/>`, pp.OutlineLevel-1)) // 0-based in XML
	}

	// Indentation
	if pp.IndentLeft != 0 || pp.IndentRight != 0 || pp.IndentFirstLine != 0 {
		buf.WriteString(`<w:ind`)

		if pp.IndentLeft != 0 {
			buf.WriteString(fmt.Sprintf(` w:left="%d"`, int(pp.IndentLeft*20))) // Convert to twips
		}

		if pp.IndentRight != 0 {
			buf.WriteString(fmt.Sprintf(` w:right="%d"`, int(pp.IndentRight*20)))
		}

		if pp.IndentFirstLine > 0 {
			buf.WriteString(fmt.Sprintf(` w:firstLine="%d"`, int(pp.IndentFirstLine*20)))
		} else if pp.IndentFirstLine < 0 {
			buf.WriteString(fmt.Sprintf(` w:hanging="%d"`, int(-pp.IndentFirstLine*20)))
		}

		buf.WriteString(`/>`)
	}

	// Spacing
	if pp.SpacingBefore != 0 || pp.SpacingAfter != 0 || pp.LineSpacing != 0 {
		buf.WriteString(`<w:spacing`)

		if pp.SpacingBefore > 0 {
			buf.WriteString(fmt.Sprintf(` w:before="%d"`, int(pp.SpacingBefore*20)))
		}

		if pp.SpacingAfter > 0 {
			buf.WriteString(fmt.Sprintf(` w:after="%d"`, int(pp.SpacingAfter*20)))
		}

		if pp.LineSpacing > 0 {
			switch pp.LineSpacingRule {
			case "exact":
				buf.WriteString(fmt.Sprintf(` w:line="%d" w:lineRule="exact"`, int(pp.LineSpacing*20)))
			case "atLeast":
				buf.WriteString(fmt.Sprintf(` w:line="%d" w:lineRule="atLeast"`, int(pp.LineSpacing*20)))
			default: // auto
				// For auto, 240 = single space, 360 = 1.5, 480 = double
				buf.WriteString(fmt.Sprintf(` w:line="%d" w:lineRule="auto"`, int(pp.LineSpacing*240)))
			}
		}

		buf.WriteString(`/>`)
	}

	buf.WriteString(`</w:pPr>`)

	return buf.Bytes(), nil
}
