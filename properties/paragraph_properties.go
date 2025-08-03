// File: properties/paragraph_properties.go
package properties

import (
	"fmt"
)

// ParagraphProperties defines paragraph formatting
type ParagraphProperties struct {
	// Alignment
	Alignment     string // left, center, right, justify, distribute, start, end
	TextAlignment string // auto, baseline, bottom, center, top

	// Indentation (in points)
	IndentLeft      float64 // Left indentation
	IndentRight     float64 // Right indentation
	IndentFirstLine float64 // First line indent (negative for hanging)
	IndentHanging   float64 // Hanging indent (alternative to negative FirstLine)

	// Spacing (in points)
	SpacingBefore     float64 // Space before paragraph
	SpacingAfter      float64 // Space after paragraph
	SpacingBeforeAuto bool    // Auto space before
	SpacingAfterAuto  bool    // Auto space after

	// Line spacing
	LineSpacing     float64 // Line spacing value
	LineSpacingRule string  // auto, exact, atLeast

	// Keep properties
	KeepNext        bool // Keep with next paragraph
	KeepLines       bool // Keep lines together
	PageBreakBefore bool // Page break before paragraph
	WidowControl    bool // Widow/orphan control

	// Paragraph style
	StyleID string // Reference to paragraph style

	// Outline and numbering
	OutlineLevel   int    // Outline level (0 = body text, 1-9 = heading levels)
	NumberingID    string // Numbering definition ID
	NumberingLevel int    // Numbering level (0-8)

	// Borders
	Borders *ParagraphBorders

	// Shading
	Shading *ParagraphShading

	// Tabs
	Tabs []TabStop

	// Text direction
	BiDi          bool   // Right-to-left paragraph
	TextDirection string // lr, tb, tbV, rl, rlV

	// Additional properties
	SuppressAutoHyphens bool   // Suppress automatic hyphenation
	SuppressLineNumbers bool   // Suppress line numbers
	SuppressOverlap     bool   // Suppress overlapping text
	ContextualSpacing   bool   // Use contextual spacing
	MirrorIndents       bool   // Mirror indents (for facing pages)
	AdjustRightInd      bool   // Automatically adjust right indent
	SnapToGrid          bool   // Snap to document grid
	DivID               string // Associated HTML div ID

	// Frame properties
	Frame *ParagraphFrame

	// Section properties (for last paragraph in section)
	SectionProperties *SectionProperties
}

// ParagraphBorders defines paragraph borders
type ParagraphBorders struct {
	Top     *Border
	Bottom  *Border
	Left    *Border
	Right   *Border
	Between *Border // Border between paragraphs
	Bar     *Border // Vertical bar
}

// Border defines a single border
type Border struct {
	Type   string // single, double, triple, thick, dotted, dashed, etc.
	Width  int    // Width in eighths of a point
	Space  int    // Space in points
	Color  string // RGB hex color
	Shadow bool   // Shadow effect
	Frame  bool   // Frame effect
}

// ParagraphShading defines paragraph background
type ParagraphShading struct {
	Fill         string // Background color (RGB hex)
	Color        string // Foreground pattern color
	Pattern      string // Pattern type: clear, solid, horzStripe, etc.
	PatternColor string // Pattern color
}

// TabStop defines a tab stop
type TabStop struct {
	Position  int    // Position in twips
	Alignment string // left, center, right, decimal, bar
	Leader    string // none, dot, hyphen, underscore, heavy, middleDot
}

// ParagraphFrame defines text frame properties
type ParagraphFrame struct {
	Width            int    // Frame width in twips
	Height           int    // Frame height in twips
	HorizontalRule   string // off, exact, atLeast, auto
	HorizontalAnchor string // margin, page, text
	VerticalAnchor   string // margin, page, text
	HorizontalSpace  int    // Horizontal padding
	VerticalSpace    int    // Vertical padding
	Wrap             string // around, tight, through, none
	VAnchorLock      bool   // Lock vertical anchor
	HAnchorLock      bool   // Lock horizontal anchor
	XAlign           string // left, center, right, inside, outside
	YAlign           string // inline, top, center, bottom, inside, outside
	X                int    // Absolute X position
	Y                int    // Absolute Y position
}

// SectionProperties defines section formatting
type SectionProperties struct {
	Type           string // continuous, nextPage, nextColumn, evenPage, oddPage
	PageSize       *PageSize
	PageMargins    *PageMargins
	Columns        *Columns
	PageNumbering  *PageNumbering
	HeaderDistance int
	FooterDistance int
	Gutter         int
	LineNumbers    *LineNumbers
	DocGrid        *DocumentGrid
	FormProtection bool
	VerticalAlign  string // top, center, bottom, justify
	BiDi           bool   // Right-to-left section
}

// PageSize defines page dimensions
type PageSize struct {
	Width       int
	Height      int
	Orientation string // portrait, landscape
	Code        int    // Paper size code
}

// PageMargins defines page margins
type PageMargins struct {
	Top    int
	Right  int
	Bottom int
	Left   int
	Header int
	Footer int
	Gutter int
}

// Columns defines column layout
type Columns struct {
	Count      int
	Space      int
	Separator  bool
	EqualWidth bool
	Columns    []Column
}

// Column defines a single column
type Column struct {
	Width int
	Space int
}

// PageNumbering defines page numbering
type PageNumbering struct {
	Start        int
	Format       string // decimal, upperRoman, lowerRoman, upperLetter, lowerLetter
	ChapterSep   string // hyphen, period, colon, emDash, enDash
	ChapterStyle string
}

// LineNumbers defines line numbering
type LineNumbers struct {
	CountBy  int
	Start    int
	Restart  string // continuous, newPage, newSection
	Distance int
}

// DocumentGrid defines document grid
type DocumentGrid struct {
	Type      string // default, lines, linesAndChars, snapToChars
	LinePitch int
	CharSpace int
}

// NewParagraphProperties creates new paragraph properties with defaults
func NewParagraphProperties() *ParagraphProperties {
	return &ParagraphProperties{
		Alignment:       "left",
		LineSpacing:     1.15, // Default 1.15 line spacing
		LineSpacingRule: "auto",
		WidowControl:    true, // Default widow control on
		SnapToGrid:      true, // Default snap to grid
		SpacingAfter:    8,    // Default 8pt after
	}
}

// Clone creates a deep copy of ParagraphProperties
func (pp *ParagraphProperties) Clone() *ParagraphProperties {
	if pp == nil {
		return nil
	}

	clone := &ParagraphProperties{
		Alignment:           pp.Alignment,
		TextAlignment:       pp.TextAlignment,
		IndentLeft:          pp.IndentLeft,
		IndentRight:         pp.IndentRight,
		IndentFirstLine:     pp.IndentFirstLine,
		IndentHanging:       pp.IndentHanging,
		SpacingBefore:       pp.SpacingBefore,
		SpacingAfter:        pp.SpacingAfter,
		SpacingBeforeAuto:   pp.SpacingBeforeAuto,
		SpacingAfterAuto:    pp.SpacingAfterAuto,
		LineSpacing:         pp.LineSpacing,
		LineSpacingRule:     pp.LineSpacingRule,
		KeepNext:            pp.KeepNext,
		KeepLines:           pp.KeepLines,
		PageBreakBefore:     pp.PageBreakBefore,
		WidowControl:        pp.WidowControl,
		StyleID:             pp.StyleID,
		OutlineLevel:        pp.OutlineLevel,
		NumberingID:         pp.NumberingID,
		NumberingLevel:      pp.NumberingLevel,
		BiDi:                pp.BiDi,
		TextDirection:       pp.TextDirection,
		SuppressAutoHyphens: pp.SuppressAutoHyphens,
		SuppressLineNumbers: pp.SuppressLineNumbers,
		SuppressOverlap:     pp.SuppressOverlap,
		ContextualSpacing:   pp.ContextualSpacing,
		MirrorIndents:       pp.MirrorIndents,
		AdjustRightInd:      pp.AdjustRightInd,
		SnapToGrid:          pp.SnapToGrid,
		DivID:               pp.DivID,
	}

	// Clone complex properties
	if pp.Borders != nil {
		clone.Borders = pp.Borders.Clone()
	}

	if pp.Shading != nil {
		clone.Shading = &ParagraphShading{
			Fill:         pp.Shading.Fill,
			Color:        pp.Shading.Color,
			Pattern:      pp.Shading.Pattern,
			PatternColor: pp.Shading.PatternColor,
		}
	}

	if len(pp.Tabs) > 0 {
		clone.Tabs = make([]TabStop, len(pp.Tabs))
		copy(clone.Tabs, pp.Tabs)
	}

	if pp.Frame != nil {
		clone.Frame = pp.Frame.Clone()
	}

	if pp.SectionProperties != nil {
		clone.SectionProperties = pp.SectionProperties.Clone()
	}

	return clone
}

// Merge merges another ParagraphProperties into this one
func (pp *ParagraphProperties) Merge(other *ParagraphProperties) {
	if other == nil {
		return
	}

	// Merge simple properties
	if other.Alignment != "" {
		pp.Alignment = other.Alignment
	}
	if other.TextAlignment != "" {
		pp.TextAlignment = other.TextAlignment
	}
	if other.StyleID != "" {
		pp.StyleID = other.StyleID
	}

	// Merge numeric properties (only if non-zero)
	if other.IndentLeft != 0 {
		pp.IndentLeft = other.IndentLeft
	}
	if other.IndentRight != 0 {
		pp.IndentRight = other.IndentRight
	}
	if other.IndentFirstLine != 0 {
		pp.IndentFirstLine = other.IndentFirstLine
	}
	if other.SpacingBefore != 0 {
		pp.SpacingBefore = other.SpacingBefore
	}
	if other.SpacingAfter != 0 {
		pp.SpacingAfter = other.SpacingAfter
	}
	if other.LineSpacing != 0 {
		pp.LineSpacing = other.LineSpacing
	}

	// Merge boolean properties (always take from other)
	pp.KeepNext = other.KeepNext
	pp.KeepLines = other.KeepLines
	pp.PageBreakBefore = other.PageBreakBefore
	pp.WidowControl = other.WidowControl
	pp.BiDi = other.BiDi

	// Merge complex properties
	if other.Borders != nil {
		pp.Borders = other.Borders.Clone()
	}
	if other.Shading != nil {
		pp.Shading = other.Shading
	}
	if len(other.Tabs) > 0 {
		pp.Tabs = make([]TabStop, len(other.Tabs))
		copy(pp.Tabs, other.Tabs)
	}
}

// Reset clears all formatting
func (pp *ParagraphProperties) Reset() {
	*pp = *NewParagraphProperties()
}

// IsEmpty returns true if no formatting is applied
func (pp *ParagraphProperties) IsEmpty() bool {
	if pp == nil {
		return true
	}

	def := NewParagraphProperties()

	return pp.Alignment == def.Alignment &&
		pp.TextAlignment == "" &&
		pp.IndentLeft == 0 &&
		pp.IndentRight == 0 &&
		pp.IndentFirstLine == 0 &&
		pp.SpacingBefore == 0 &&
		pp.SpacingAfter == def.SpacingAfter &&
		pp.LineSpacing == def.LineSpacing &&
		pp.LineSpacingRule == def.LineSpacingRule &&
		!pp.KeepNext &&
		!pp.KeepLines &&
		!pp.PageBreakBefore &&
		pp.WidowControl == def.WidowControl &&
		pp.StyleID == "" &&
		pp.OutlineLevel == 0 &&
		pp.NumberingID == "" &&
		pp.Borders == nil &&
		pp.Shading == nil &&
		len(pp.Tabs) == 0
}

// Validate validates the paragraph properties
func (pp *ParagraphProperties) Validate() error {
	// Validate alignment
	validAlignments := map[string]bool{
		"left": true, "center": true, "right": true, "justify": true,
		"distribute": true, "start": true, "end": true, "": true,
	}
	if !validAlignments[pp.Alignment] {
		return fmt.Errorf("invalid alignment: %s", pp.Alignment)
	}

	// Validate text alignment
	validTextAlignments := map[string]bool{
		"auto": true, "baseline": true, "bottom": true, "center": true,
		"top": true, "": true,
	}
	if !validTextAlignments[pp.TextAlignment] {
		return fmt.Errorf("invalid text alignment: %s", pp.TextAlignment)
	}

	// Validate line spacing rule
	validLineSpacingRules := map[string]bool{
		"auto": true, "exact": true, "atLeast": true, "": true,
	}
	if !validLineSpacingRules[pp.LineSpacingRule] {
		return fmt.Errorf("invalid line spacing rule: %s", pp.LineSpacingRule)
	}

	// Validate outline level
	if pp.OutlineLevel < 0 || pp.OutlineLevel > 9 {
		return fmt.Errorf("outline level must be between 0 and 9: %d", pp.OutlineLevel)
	}

	// Validate numbering level
	if pp.NumberingLevel < 0 || pp.NumberingLevel > 8 {
		return fmt.Errorf("numbering level must be between 0 and 8: %d", pp.NumberingLevel)
	}

	// Validate borders
	if pp.Borders != nil {
		if err := pp.Borders.Validate(); err != nil {
			return fmt.Errorf("invalid borders: %w", err)
		}
	}

	// Validate tabs
	for i, tab := range pp.Tabs {
		if err := tab.Validate(); err != nil {
			return fmt.Errorf("invalid tab stop %d: %w", i, err)
		}
	}

	return nil
}

// Clone creates a copy of ParagraphBorders
func (pb *ParagraphBorders) Clone() *ParagraphBorders {
	if pb == nil {
		return nil
	}

	clone := &ParagraphBorders{}

	if pb.Top != nil {
		clone.Top = pb.Top.Clone()
	}
	if pb.Bottom != nil {
		clone.Bottom = pb.Bottom.Clone()
	}
	if pb.Left != nil {
		clone.Left = pb.Left.Clone()
	}
	if pb.Right != nil {
		clone.Right = pb.Right.Clone()
	}
	if pb.Between != nil {
		clone.Between = pb.Between.Clone()
	}
	if pb.Bar != nil {
		clone.Bar = pb.Bar.Clone()
	}

	return clone
}

// Validate validates paragraph borders
func (pb *ParagraphBorders) Validate() error {
	borders := []*Border{pb.Top, pb.Bottom, pb.Left, pb.Right, pb.Between, pb.Bar}

	for _, border := range borders {
		if border != nil {
			if err := border.Validate(); err != nil {
				return err
			}
		}
	}

	return nil
}

// XML generates XML for paragraph borders
func (pb *ParagraphBorders) XML() ([]byte, error) {
	// Implementation would generate proper OOXML
	return nil, nil
}

// Clone creates a copy of Border
func (b *Border) Clone() *Border {
	if b == nil {
		return nil
	}

	return &Border{
		Type:   b.Type,
		Width:  b.Width,
		Space:  b.Space,
		Color:  b.Color,
		Shadow: b.Shadow,
		Frame:  b.Frame,
	}
}

// Validate validates a border
func (b *Border) Validate() error {
	validTypes := map[string]bool{
		"single": true, "double": true, "triple": true, "thick": true,
		"dotted": true, "dashed": true, "dotDash": true, "dotDotDash": true,
		"wave": true, "doubleWave": true, "dashSmallGap": true,
		"dashDotStroked": true, "threeDEmboss": true, "threeDEngrave": true,
		"outset": true, "inset": true, "": true,
	}

	if !validTypes[b.Type] {
		return fmt.Errorf("invalid border type: %s", b.Type)
	}

	if b.Width < 0 {
		return fmt.Errorf("border width cannot be negative: %d", b.Width)
	}

	if b.Space < 0 {
		return fmt.Errorf("border space cannot be negative: %d", b.Space)
	}

	return nil
}

// XML generates XML for paragraph shading
func (ps *ParagraphShading) XML() ([]byte, error) {
	// Implementation would generate proper OOXML
	return nil, nil
}

// Validate validates a tab stop
func (ts TabStop) Validate() error {
	validAlignments := map[string]bool{
		"left": true, "center": true, "right": true, "decimal": true,
		"bar": true, "clear": true, "": true,
	}

	if !validAlignments[ts.Alignment] {
		return fmt.Errorf("invalid tab alignment: %s", ts.Alignment)
	}

	validLeaders := map[string]bool{
		"none": true, "dot": true, "hyphen": true, "underscore": true,
		"heavy": true, "middleDot": true, "": true,
	}

	if !validLeaders[ts.Leader] {
		return fmt.Errorf("invalid tab leader: %s", ts.Leader)
	}

	if ts.Position < 0 {
		return fmt.Errorf("tab position cannot be negative: %d", ts.Position)
	}

	return nil
}

// Clone creates a copy of ParagraphFrame
func (pf *ParagraphFrame) Clone() *ParagraphFrame {
	if pf == nil {
		return nil
	}

	return &ParagraphFrame{
		Width:            pf.Width,
		Height:           pf.Height,
		HorizontalRule:   pf.HorizontalRule,
		HorizontalAnchor: pf.HorizontalAnchor,
		VerticalAnchor:   pf.VerticalAnchor,
		HorizontalSpace:  pf.HorizontalSpace,
		VerticalSpace:    pf.VerticalSpace,
		Wrap:             pf.Wrap,
		VAnchorLock:      pf.VAnchorLock,
		HAnchorLock:      pf.HAnchorLock,
		XAlign:           pf.XAlign,
		YAlign:           pf.YAlign,
		X:                pf.X,
		Y:                pf.Y,
	}
}

// Clone creates a copy of SectionProperties
func (sp *SectionProperties) Clone() *SectionProperties {
	if sp == nil {
		return nil
	}

	clone := &SectionProperties{
		Type:           sp.Type,
		HeaderDistance: sp.HeaderDistance,
		FooterDistance: sp.FooterDistance,
		Gutter:         sp.Gutter,
		FormProtection: sp.FormProtection,
		VerticalAlign:  sp.VerticalAlign,
		BiDi:           sp.BiDi,
	}

	if sp.PageSize != nil {
		clone.PageSize = &PageSize{
			Width:       sp.PageSize.Width,
			Height:      sp.PageSize.Height,
			Orientation: sp.PageSize.Orientation,
			Code:        sp.PageSize.Code,
		}
	}

	if sp.PageMargins != nil {
		clone.PageMargins = &PageMargins{
			Top:    sp.PageMargins.Top,
			Right:  sp.PageMargins.Right,
			Bottom: sp.PageMargins.Bottom,
			Left:   sp.PageMargins.Left,
			Header: sp.PageMargins.Header,
			Footer: sp.PageMargins.Footer,
			Gutter: sp.PageMargins.Gutter,
		}
	}

	return clone
}

// SetLineSpacingSingle sets single line spacing
func (pp *ParagraphProperties) SetLineSpacingSingle() *ParagraphProperties {
	pp.LineSpacing = 1.0
	pp.LineSpacingRule = "auto"
	return pp
}

// SetLineSpacingOneAndHalf sets 1.5 line spacing
func (pp *ParagraphProperties) SetLineSpacingOneAndHalf() *ParagraphProperties {
	pp.LineSpacing = 1.5
	pp.LineSpacingRule = "auto"
	return pp
}

// SetLineSpacingDouble sets double line spacing
func (pp *ParagraphProperties) SetLineSpacingDouble() *ParagraphProperties {
	pp.LineSpacing = 2.0
	pp.LineSpacingRule = "auto"
	return pp
}

// SetLineSpacingExact sets exact line spacing in points
func (pp *ParagraphProperties) SetLineSpacingExact(points float64) *ParagraphProperties {
	pp.LineSpacing = points
	pp.LineSpacingRule = "exact"
	return pp
}

// SetLineSpacingAtLeast sets minimum line spacing in points
func (pp *ParagraphProperties) SetLineSpacingAtLeast(points float64) *ParagraphProperties {
	pp.LineSpacing = points
	pp.LineSpacingRule = "atLeast"
	return pp
}

// Common border types as constants
const (
	BorderTypeSingle     = "single"
	BorderTypeDouble     = "double"
	BorderTypeTriple     = "triple"
	BorderTypeThick      = "thick"
	BorderTypeDotted     = "dotted"
	BorderTypeDashed     = "dashed"
	BorderTypeDotDash    = "dotDash"
	BorderTypeDotDotDash = "dotDotDash"
	BorderTypeWave       = "wave"
	BorderTypeDoubleWave = "doubleWave"
)

// Common shading patterns
const (
	ShadingPatternClear      = "clear"
	ShadingPatternSolid      = "solid"
	ShadingPatternHorzStripe = "horzStripe"
	ShadingPatternVertStripe = "vertStripe"
	ShadingPatternDiagStripe = "diagStripe"
	ShadingPatternPct10      = "pct10"
	ShadingPatternPct20      = "pct20"
	ShadingPatternPct25      = "pct25"
	ShadingPatternPct50      = "pct50"
)
