// File: elements/break.go
package elements

import "strings"

// LineBreak represents a line break within a run
type LineBreak struct {
	Typ   string // Type of break: "textWrapping" (default), "page", "column", "textWrapping"
	Clear string // Clear type for text wrapping: "none", "left", "right", "all"
}

// NewLineBreak creates a new line break
func NewLineBreak() *LineBreak {
	return &LineBreak{
		Typ: "textWrapping",
	}
}

// NewPageBreak creates a new page break
func NewPageBreak() *PageBreak {
	return &PageBreak{}
}

// NewColumnBreak creates a new column break
func NewColumnBreak() *LineBreak {
	return &LineBreak{
		Typ: "column",
	}
}

// NewTextWrappingBreak creates a line break that clears floating objects
func NewTextWrappingBreak(clear string) *LineBreak {
	return &LineBreak{
		Typ:   "textWrapping",
		Clear: clear,
	}
}

// Type returns the element type
func (lb *LineBreak) Type() string {
	return "lineBreak"
}

// SetType sets the break type
func (lb *LineBreak) SetType(breakType string) *LineBreak {
	lb.Typ = breakType
	return lb
}

// SetClear sets the clear type for text wrapping breaks
func (lb *LineBreak) SetClear(clear string) *LineBreak {
	lb.Clear = clear
	return lb
}

// XML generates the XML representation
func (lb *LineBreak) XML() ([]byte, error) {
	if lb.Typ == "" || lb.Typ == "textWrapping" {
		if lb.Clear != "" && lb.Clear != "none" {
			return []byte(`<w:br w:type="textWrapping" w:clear="` + lb.Clear + `"/>`), nil
		}
		return []byte(`<w:br/>`), nil
	}

	return []byte(`<w:br w:type="` + lb.Typ + `"/>`), nil
}

// PageBreak represents a page break element
type PageBreak struct{}

// Type returns the element type
func (pb *PageBreak) Type() string {
	return "pageBreak"
}

// XML generates the XML representation
func (pb *PageBreak) XML() ([]byte, error) {
	// Page break is a paragraph with a page break run
	return []byte(`<w:p><w:r><w:br w:type="page"/></w:r></w:p>`), nil
}

// SectionBreak represents a section break
type SectionBreak struct {
	Typ string // Type: "nextPage", "continuous", "evenPage", "oddPage", "nextColumn"
}

// NewSectionBreak creates a new section break
func NewSectionBreak(breakType string) *SectionBreak {
	return &SectionBreak{
		Typ: breakType,
	}
}

// Type returns the element type
func (sb *SectionBreak) Type() string {
	return "sectionBreak"
}

// XML generates the XML representation
func (sb *SectionBreak) XML() ([]byte, error) {
	// Section breaks are implemented through paragraph properties
	return []byte(`<w:p><w:pPr><w:sectPr><w:type w:val="` + sb.Typ + `"/></w:sectPr></w:pPr></w:p>`), nil
}

// Tab represents a tab character
type Tab struct {
	Position  int    // Position in twips (optional, for positional tabs)
	Alignment string // Alignment: "left", "center", "right", "decimal", "bar"
	Leader    string // Leader character: "dot", "hyphen", "underscore", "heavy", "middleDot"
}

// NewTab creates a new tab
func NewTab() *Tab {
	return &Tab{}
}

// NewPositionalTab creates a tab at a specific position
func NewPositionalTab(position int, alignment string) *Tab {
	return &Tab{
		Position:  position,
		Alignment: alignment,
	}
}

// Type returns the element type
func (t *Tab) Type() string {
	return "tab"
}

// SetPosition sets the tab position in twips
func (t *Tab) SetPosition(position int) *Tab {
	t.Position = position
	return t
}

// SetAlignment sets the tab alignment
func (t *Tab) SetAlignment(alignment string) *Tab {
	t.Alignment = alignment
	return t
}

// SetLeader sets the tab leader character
func (t *Tab) SetLeader(leader string) *Tab {
	t.Leader = leader
	return t
}

// XML generates the XML representation
func (t *Tab) XML() ([]byte, error) {
	// Simple tab (most common case)
	if t.Position == 0 && t.Alignment == "" && t.Leader == "" {
		return []byte(`<w:tab/>`), nil
	}

	// Positional tab (requires run properties)
	// This would typically be handled at the paragraph level
	// For now, return a simple tab
	return []byte(`<w:tab/>`), nil
}

// CarriageReturn represents a carriage return (CR) character
type CarriageReturn struct{}

// NewCarriageReturn creates a new carriage return
func NewCarriageReturn() *CarriageReturn {
	return &CarriageReturn{}
}

// Type returns the element type
func (cr *CarriageReturn) Type() string {
	return "carriageReturn"
}

// XML generates the XML representation
func (cr *CarriageReturn) XML() ([]byte, error) {
	return []byte(`<w:cr/>`), nil
}

// NoBreakHyphen represents a non-breaking hyphen
type NoBreakHyphen struct{}

// NewNoBreakHyphen creates a new non-breaking hyphen
func NewNoBreakHyphen() *NoBreakHyphen {
	return &NoBreakHyphen{}
}

// Type returns the element type
func (nbh *NoBreakHyphen) Type() string {
	return "noBreakHyphen"
}

// XML generates the XML representation
func (nbh *NoBreakHyphen) XML() ([]byte, error) {
	return []byte(`<w:noBreakHyphen/>`), nil
}

// SoftHyphen represents a soft hyphen (optional hyphen)
type SoftHyphen struct{}

// NewSoftHyphen creates a new soft hyphen
func NewSoftHyphen() *SoftHyphen {
	return &SoftHyphen{}
}

// Type returns the element type
func (sh *SoftHyphen) Type() string {
	return "softHyphen"
}

// XML generates the XML representation
func (sh *SoftHyphen) XML() ([]byte, error) {
	return []byte(`<w:softHyphen/>`), nil
}

// Symbol represents a symbol character
type Symbol struct {
	Font string // Font name
	Char string // Character code (hex)
}

// NewSymbol creates a new symbol
func NewSymbol(font, char string) *Symbol {
	return &Symbol{
		Font: font,
		Char: char,
	}
}

// Type returns the element type
func (s *Symbol) Type() string {
	return "symbol"
}

// XML generates the XML representation
func (s *Symbol) XML() ([]byte, error) {
	return []byte(`<w:sym w:font="` + s.Font + `" w:char="` + s.Char + `"/>`), nil
}

// LastRenderedPageBreak represents the position of a page break in the last rendering
type LastRenderedPageBreak struct{}

// NewLastRenderedPageBreak creates a new last rendered page break marker
func NewLastRenderedPageBreak() *LastRenderedPageBreak {
	return &LastRenderedPageBreak{}
}

// Type returns the element type
func (lrpb *LastRenderedPageBreak) Type() string {
	return "lastRenderedPageBreak"
}

// XML generates the XML representation
func (lrpb *LastRenderedPageBreak) XML() ([]byte, error) {
	return []byte(`<w:lastRenderedPageBreak/>`), nil
}

// Helper functions for common use cases

// AddLineBreakToRun adds a line break to a run
func AddLineBreakToRun(run *Run) {
	run.Children = append(run.Children, NewLineBreak())
}

// AddPageBreakToParagraph adds a page break run to a paragraph
func AddPageBreakToParagraph(p *Paragraph) {
	run := NewRun()
	run.Children = append(run.Children, &LineBreak{Typ: "page"})
	p.Children = append(p.Children, run)
}

// AddTabToRun adds a tab to a run
func AddTabToRun(run *Run) {
	run.Children = append(run.Children, NewTab())
}

// AddMultipleTabsToRun adds multiple tabs to a run
func AddMultipleTabsToRun(run *Run, count int) {
	for i := 0; i < count; i++ {
		run.Children = append(run.Children, NewTab())
	}
}

// AddCarriageReturnToRun adds a carriage return to a run
func AddCarriageReturnToRun(run *Run) {
	run.Children = append(run.Children, NewCarriageReturn())
}

// AddSymbolToRun adds a symbol to a run
func AddSymbolToRun(run *Run, font, char string) {
	run.Children = append(run.Children, NewSymbol(font, char))
}

// Common symbols as constants
const (
	// Wingdings symbols
	SymbolCheckmark  = "F0FE" // ✓
	SymbolCross      = "F0FB" // ✗
	SymbolStar       = "F0AB" // ★
	SymbolArrowRight = "F0E0" // →
	SymbolArrowLeft  = "F0DF" // ←
	SymbolArrowUp    = "F0DD" // ↑
	SymbolArrowDown  = "F0DE" // ↓

	// Common Unicode symbols (using Symbol font)
	SymbolCopyright  = "00A9" // ©
	SymbolRegistered = "00AE" // ®
	SymbolTrademark  = "2122" // ™
	SymbolDegree     = "00B0" // °
	SymbolPlusMinus  = "00B1" // ±
	SymbolEuro       = "20AC" // €
	SymbolPound      = "00A3" // £
	SymbolYen        = "00A5" // ¥
)

// AddCommonSymbol adds a common symbol to a run
func AddCommonSymbol(run *Run, symbolCode string) {
	font := "Symbol"
	if symbolCode == SymbolCheckmark || symbolCode == SymbolCross ||
		symbolCode == SymbolStar || strings.HasPrefix(symbolCode, "F0") {
		font = "Wingdings"
	}
	AddSymbolToRun(run, font, symbolCode)
}
