// styles/styles.go
package styles

import "github.com/didikprabowo/mbadocx/properties"

// DocumentStyles represents document style configuration
type DocumentStyles struct {
	// Default styles
	Normal   *ParagraphStyle
	Heading1 *ParagraphStyle
	Heading2 *ParagraphStyle
	Heading3 *ParagraphStyle
	Heading4 *ParagraphStyle
	Heading5 *ParagraphStyle
	Heading6 *ParagraphStyle

	// Special styles
	Title    *ParagraphStyle
	Subtitle *ParagraphStyle
	Quote    *ParagraphStyle
	Caption  *ParagraphStyle

	// List styles
	ListParagraph *ParagraphStyle

	// Character styles
	DefaultCharacter *CharacterStyle
	Emphasis         *CharacterStyle
	Strong           *CharacterStyle
	Hyperlink        *CharacterStyle

	// Table styles
	TableNormal *TableStyle

	// Custom styles
	Custom map[string]*ParagraphStyle
}

// ParagraphStyle represents a paragraph style
type ParagraphStyle struct {
	StyleId      string
	Name         string
	Type         string // "paragraph", "character", "table"
	BasedOn      string
	NextStyle    string
	IsDefault    bool
	IsQFormat    bool
	OutlineLevel int

	// Paragraph properties
	Paragraph *properties.ParagraphProperties

	// Run properties
	Run *properties.RunProperties
}

// CharacterStyle represents a character style
type CharacterStyle struct {
	StyleId   string
	Name      string
	BasedOn   string
	IsDefault bool

	// Run properties
	Run *properties.RunProperties
}

// TableStyle represents a table style
type TableStyle struct {
	StyleId   string
	Name      string
	BasedOn   string
	IsDefault bool

	// Table properties would go here
}

// ParagraphProperties represents paragraph-level formatting
type ParagraphProperties struct {
	// Spacing
	SpacingBefore int    // Twips
	SpacingAfter  int    // Twips
	LineSpacing   int    // Twips
	LineRule      string // "auto", "exact", "atLeast"

	// Indentation
	LeftIndent  int // Twips
	RightIndent int // Twips
	FirstIndent int // Twips (can be negative for hanging)

	// Alignment
	Alignment string // "left", "center", "right", "justify"

	// Behavior
	KeepNext     bool // Keep with next paragraph
	KeepLines    bool // Keep lines together
	PageBreak    bool // Page break before
	WidowControl bool // Widow/orphan control

	// Numbering
	NumberingId  int // Numbering definition ID
	NumberingLvl int // Numbering level

	// Borders and shading
	Borders *BorderProperties
	Shading *ShadingProperties

	// Tabs
	Tabs []TabStop
}

// BorderProperties represents border formatting
type BorderProperties struct {
	Top    *Border
	Bottom *Border
	Left   *Border
	Right  *Border
}

// Border represents a single border
type Border struct {
	Style string // "single", "double", "thick", etc.
	Width int    // Eighth-points
	Color string // Hex color
	Space int    // Points
}

// ShadingProperties represents shading/background
type ShadingProperties struct {
	Pattern string // "solid", "pct10", etc.
	Fill    string // Background color (hex)
	Color   string // Foreground color (hex)
}

// TabStop represents a tab stop
type TabStop struct {
	Position int    // Twips from left margin
	Align    string // "left", "center", "right", "decimal"
	Leader   string // "none", "dot", "hyphen", "underscore"
}

var isTrue = func(b bool) *bool { return &b }(true)

// DefaultDocumentStyles returns default document styles
func DefaultDocumentStyles() *DocumentStyles {
	return &DocumentStyles{
		Normal: &ParagraphStyle{
			StyleId:   "Normal",
			Name:      "Normal",
			Type:      "paragraph",
			IsDefault: true,
			IsQFormat: true,
			Paragraph: &properties.ParagraphProperties{
				SpacingAfter:    200, // 10pt
				LineSpacing:     276, // 1.15 line spacing
				LineSpacingRule: "auto",
				Alignment:       "left",
			},
			Run: &properties.RunProperties{
				FontFamily: "Calibri",
				FontSize:   22, // 11pt
				Language:   "en-US",
			},
		},

		Heading1: &ParagraphStyle{
			StyleId:      "Heading1",
			Name:         "heading 1",
			Type:         "paragraph",
			BasedOn:      "Normal",
			NextStyle:    "Normal",
			IsQFormat:    true,
			OutlineLevel: 0,
			Paragraph: &properties.ParagraphProperties{
				SpacingBefore: 240, // 12pt
				SpacingAfter:  0,
				KeepNext:      true,
				KeepLines:     true,
			},
			Run: &properties.RunProperties{
				FontFamily: "Calibri Light",
				FontSize:   32,       // 16pt
				Color:      "2F5496", // Blue
			},
		},

		Heading2: &ParagraphStyle{
			StyleId:      "Heading2",
			Name:         "heading 2",
			Type:         "paragraph",
			BasedOn:      "Normal",
			NextStyle:    "Normal",
			IsQFormat:    true,
			OutlineLevel: 1,
			Paragraph: &properties.ParagraphProperties{
				SpacingBefore: 200, // 10pt
				SpacingAfter:  0,
				KeepNext:      true,
				KeepLines:     true,
			},
			Run: &properties.RunProperties{
				FontFamily: "Calibri Light",
				FontSize:   26,       // 13pt
				Color:      "2F5496", // Blue
			},
		},

		Heading3: &ParagraphStyle{
			StyleId:      "Heading3",
			Name:         "heading 3",
			Type:         "paragraph",
			BasedOn:      "Normal",
			NextStyle:    "Normal",
			IsQFormat:    true,
			OutlineLevel: 2,
			Paragraph: &properties.ParagraphProperties{
				SpacingBefore: 200, // 10pt
				SpacingAfter:  0,
				KeepNext:      true,
				KeepLines:     true,
			},
			Run: &properties.RunProperties{
				FontFamily: "Calibri Light",
				FontSize:   24,       // 12pt
				Color:      "1F3763", // Dark blue
			},
		},
		Heading4: &ParagraphStyle{
			StyleId:      "Heading4",
			Name:         "heading 4",
			Type:         "paragraph",
			BasedOn:      "Normal",
			NextStyle:    "Normal",
			IsQFormat:    true,
			OutlineLevel: 3,
			Paragraph: &properties.ParagraphProperties{
				SpacingBefore: 200,
				SpacingAfter:  0,
				KeepNext:      true,
				KeepLines:     true,
			},
			Run: &properties.RunProperties{
				FontFamily: "Calibri Light",
				FontSize:   22, // 11pt
				Color:      "2F5496",
				Italic:     isTrue,
			},
		},

		Heading5: &ParagraphStyle{
			StyleId:      "Heading5",
			Name:         "heading 5",
			Type:         "paragraph",
			BasedOn:      "Normal",
			NextStyle:    "Normal",
			IsQFormat:    true,
			OutlineLevel: 4,
			Paragraph: &properties.ParagraphProperties{
				SpacingBefore: 200,
				SpacingAfter:  0,
				KeepNext:      true,
				KeepLines:     true,
			},
			Run: &properties.RunProperties{
				FontFamily: "Calibri Light",
				FontSize:   22, // 11pt
				Color:      "2F5496",
			},
		},

		Heading6: &ParagraphStyle{
			StyleId:      "Heading6",
			Name:         "heading 6",
			Type:         "paragraph",
			BasedOn:      "Normal",
			NextStyle:    "Normal",
			IsQFormat:    true,
			OutlineLevel: 5,
			Paragraph: &properties.ParagraphProperties{
				SpacingBefore: 200,
				SpacingAfter:  0,
				KeepNext:      true,
				KeepLines:     true,
			},
			Run: &properties.RunProperties{
				FontFamily: "Calibri Light",
				FontSize:   20, // 10pt
				Color:      "1F3763",
			},
		},

		Title: &ParagraphStyle{
			StyleId:   "Title",
			Name:      "Title",
			Type:      "paragraph",
			BasedOn:   "Normal",
			NextStyle: "Normal",
			IsQFormat: true,
			Paragraph: &properties.ParagraphProperties{
				SpacingBefore: 0,
				SpacingAfter:  0,
				Alignment:     "center",
			},
			Run: &properties.RunProperties{
				FontFamily: "Calibri Light",
				FontSize:   56, // 28pt
				Color:      "2F5496",
			},
		},

		Subtitle: &ParagraphStyle{
			StyleId:   "Subtitle",
			Name:      "Subtitle",
			Type:      "paragraph",
			BasedOn:   "Normal",
			NextStyle: "Normal",
			IsQFormat: true,
			Paragraph: &properties.ParagraphProperties{
				SpacingBefore: 0,
				SpacingAfter:  200,
				Alignment:     "center",
			},
			Run: &properties.RunProperties{
				FontFamily: "Calibri",
				FontSize:   30,       // 15pt
				Color:      "595959", // Gray
				Italic:     isTrue,
			},
		},

		Quote: &ParagraphStyle{
			StyleId:   "Quote",
			Name:      "Quote",
			Type:      "paragraph",
			BasedOn:   "Normal",
			NextStyle: "Normal",
			IsQFormat: true,
			Paragraph: &properties.ParagraphProperties{
				SpacingBefore: 200,
				SpacingAfter:  200,
				IndentLeft:    720, // 0.5 inch
				IndentRight:   720, // 0.5 inch
				Alignment:     "center",
			},
			Run: &properties.RunProperties{
				FontFamily: "Calibri",
				FontSize:   22,
				Color:      "404040",
				Italic:     isTrue,
			},
		},

		Caption: &ParagraphStyle{
			StyleId:   "Caption",
			Name:      "Caption",
			Type:      "paragraph",
			BasedOn:   "Normal",
			NextStyle: "Normal",
			IsQFormat: true,
			Paragraph: &properties.ParagraphProperties{
				SpacingBefore: 120,
				SpacingAfter:  120,
			},
			Run: &properties.RunProperties{
				FontFamily: "Calibri",
				FontSize:   18, // 9pt
				Color:      "404040",
			},
		},

		ListParagraph: &ParagraphStyle{
			StyleId:   "ListParagraph",
			Name:      "List Paragraph",
			Type:      "paragraph",
			BasedOn:   "Normal",
			NextStyle: "Normal",
			IsQFormat: true,
			Paragraph: &properties.ParagraphProperties{
				IndentLeft: 720, // 0.5 inch
			},
			Run: &properties.RunProperties{
				FontFamily: "Calibri",
				FontSize:   22,
			},
		},

		DefaultCharacter: &CharacterStyle{
			StyleId:   "DefaultParagraphFont",
			Name:      "Default Paragraph Font",
			IsDefault: true,
			Run: &properties.RunProperties{
				FontFamily: "Calibri",
				FontSize:   22,
			},
		},

		Emphasis: &CharacterStyle{
			StyleId: "Emphasis",
			Name:    "Emphasis",
			Run: &properties.RunProperties{
				Italic: isTrue,
			},
		},

		Strong: &CharacterStyle{
			StyleId: "Strong",
			Name:    "Strong",
			Run: &properties.RunProperties{
				Bold: isTrue,
			},
		},

		Hyperlink: &CharacterStyle{
			StyleId: "Hyperlink",
			Name:    "Hyperlink",
			Run: &properties.RunProperties{
				Color:     "0563C1", // Blue
				Underline: "single",
			},
		},

		TableNormal: &TableStyle{
			StyleId:   "TableNormal",
			Name:      "Normal Table",
			IsDefault: true,
		},

		Custom: make(map[string]*ParagraphStyle),
	}
}

// BusinessDocumentStyles returns styles suitable for business documents
func BusinessDocumentStyles() *DocumentStyles {
	styles := DefaultDocumentStyles()

	// Override with more conservative business styles
	styles.Normal.Run.FontFamily = "Times New Roman"
	styles.Normal.Run.FontSize = 24 // 12pt

	styles.Heading1.Run.FontFamily = "Times New Roman"
	styles.Heading1.Run.FontSize = 32    // 16pt
	styles.Heading1.Run.Color = "000000" // Black
	styles.Heading1.Run.Bold = isTrue

	styles.Heading2.Run.FontFamily = "Times New Roman"
	styles.Heading2.Run.FontSize = 28 // 14pt
	styles.Heading2.Run.Color = "000000"
	styles.Heading2.Run.Bold = isTrue

	styles.Heading3.Run.FontFamily = "Times New Roman"
	styles.Heading3.Run.FontSize = 26 // 13pt
	styles.Heading3.Run.Color = "000000"
	styles.Heading3.Run.Bold = isTrue

	return styles
}

// ModernDocumentStyles returns modern, clean styles
func ModernDocumentStyles() *DocumentStyles {
	styles := DefaultDocumentStyles()

	// Use modern fonts
	styles.Normal.Run.FontFamily = "Segoe UI"
	styles.Heading1.Run.FontFamily = "Segoe UI"
	styles.Heading2.Run.FontFamily = "Segoe UI"
	styles.Heading3.Run.FontFamily = "Segoe UI"

	// Increase spacing for modern look
	styles.Normal.Paragraph.SpacingAfter = 240    // 12pt
	styles.Heading1.Paragraph.SpacingBefore = 480 // 24pt
	styles.Heading2.Paragraph.SpacingBefore = 360 // 18pt
	styles.Heading3.Paragraph.SpacingBefore = 240 // 12pt

	return styles
}

// AcademicDocumentStyles returns styles suitable for academic papers
func AcademicDocumentStyles() *DocumentStyles {
	styles := DefaultDocumentStyles()

	// Academic standards
	styles.Normal.Run.FontFamily = "Times New Roman"
	styles.Normal.Run.FontSize = 24           // 12pt
	styles.Normal.Paragraph.LineSpacing = 480 // Double spacing
	styles.Normal.Paragraph.LineSpacingRule = "auto"

	// Conservative heading styles
	styles.Heading1.Run.FontFamily = "Times New Roman"
	styles.Heading1.Run.FontSize = 24 // Same size as body
	styles.Heading1.Run.Color = "000000"
	styles.Heading1.Run.Bold = isTrue
	styles.Heading1.Paragraph.Alignment = "center"

	styles.Heading2.Run.FontFamily = "Times New Roman"
	styles.Heading2.Run.FontSize = 24
	styles.Heading2.Run.Color = "000000"
	styles.Heading2.Run.Bold = isTrue
	styles.Heading2.Paragraph.Alignment = "left"

	return styles
}

// Convenience methods for DocumentStyles
func (ds *DocumentStyles) AddCustomStyle(id string, style *ParagraphStyle) {
	if ds.Custom == nil {
		ds.Custom = make(map[string]*ParagraphStyle)
	}
	ds.Custom[id] = style
}

func (ds *DocumentStyles) GetStyle(id string) *ParagraphStyle {
	switch id {
	case "Normal":
		return ds.Normal
	case "Heading1":
		return ds.Heading1
	case "Heading2":
		return ds.Heading2
	case "Heading3":
		return ds.Heading3
	case "Heading4":
		return ds.Heading4
	case "Heading5":
		return ds.Heading5
	case "Heading6":
		return ds.Heading6
	case "Title":
		return ds.Title
	case "Subtitle":
		return ds.Subtitle
	case "Quote":
		return ds.Quote
	case "Caption":
		return ds.Caption
	case "ListParagraph":
		return ds.ListParagraph
	default:
		if ds.Custom != nil {
			return ds.Custom[id]
		}
		return nil
	}
}

func (ds *DocumentStyles) GetAllStyles() []*ParagraphStyle {
	styles := []*ParagraphStyle{
		ds.Normal,
		ds.Heading1,
		ds.Heading2,
		ds.Heading3,
		ds.Heading4,
		ds.Heading5,
		ds.Heading6,
		ds.Title,
		ds.Subtitle,
		ds.Quote,
		ds.Caption,
		ds.ListParagraph,
	}

	// Add custom styles
	for _, style := range ds.Custom {
		styles = append(styles, style)
	}

	return styles
}

// Helper functions for common measurements

// PointsToTwips converts points to twips (1 point = 20 twips)
func PointsToTwips(points float64) int {
	return int(points * 20)
}

// InchesToTwips converts inches to twips (1 inch = 1440 twips)
func InchesToTwips(inches float64) int {
	return int(inches * 1440)
}

// PointsToHalfPoints converts points to half-points for font sizes
func PointsToHalfPoints(points float64) int {
	return int(points * 2)
}
