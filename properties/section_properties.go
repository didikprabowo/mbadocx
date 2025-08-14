package properties

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
