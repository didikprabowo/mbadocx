// ============================================
// Clean Document Settings - No JSON Tags
// ============================================

// settings/settings.go
package settings

// DocumentSettings represents comprehensive document settings (grouped approach)
type DocumentSettings struct {
	// Page layout and margins
	Page *PageSettings

	// Font and typography
	Font *FontSettings

	// Language and locale
	Language *LanguageSettings

	// View and display
	View *ViewSettings

	// Compatibility with Word versions
	Compatibility *CompatibilitySettings

	// Document protection and security
	Protection *ProtectionSettings

	// Revision tracking and changes
	Tracking *TrackingSettings

	// Print and output settings
	Print *PrintSettings

	// Drawing and grid settings
	Grid *GridSettings

	// Web and HTML settings
	Web *WebSettings

	// Footnotes and endnotes
	Notes *NotesSettings

	// Line numbering
	LineNumbering *LineNumberingSettings

	// Document borders
	Borders *BorderSettings

	// Mail merge settings
	MailMerge *MailMergeSettings

	// Document variables
	Variables map[string]string

	// Style pane settings
	StylePane *StylePaneSettings

	// Proofing settings
	Proofing *ProofingSettings
}

// ============================================
// PAGE SETTINGS
// ============================================
type PageSettings struct {
	Width         int    // Page width in twips
	Height        int    // Page height in twips
	Orientation   string // "portrait" or "landscape"
	MarginTop     int    // Top margin in twips
	MarginRight   int    // Right margin in twips
	MarginBottom  int    // Bottom margin in twips
	MarginLeft    int    // Left margin in twips
	MarginHeader  int    // Header margin in twips
	MarginFooter  int    // Footer margin in twips
	MarginGutter  int    // Gutter margin in twips
	MirrorMargins bool   // Mirror margins for binding
	GutterAtTop   bool   // Position gutter at top
}

// ============================================
// FONT SETTINGS
// ============================================
type FontSettings struct {
	DefaultFont             string // Default font family
	DefaultSize             int    // Default size in half-points
	DefaultColor            string // Default color (hex)
	EastAsianFont           string // East Asian font
	ComplexScriptFont       string // Complex script font
	DefaultTabStop          int    // Default tab stop in twips
	CharacterSpacingControl string // Character spacing
	EmbedTrueTypeFonts      bool   // Embed TrueType fonts
	SaveSubsetFonts         bool   // Save font subsets
}

// ============================================
// LANGUAGE SETTINGS
// ============================================
type LanguageSettings struct {
	Default                 string // Default language
	EastAsian               string // East Asian language
	BiDirectional           string // Bi-directional language
	StrictFirstAndLastChars bool   // Strict character rules
	NoLineBreaksAfter       string // Characters that can't end lines
	NoLineBreaksBefore      string // Characters that can't start lines
}

// ============================================
// VIEW SETTINGS
// ============================================
type ViewSettings struct {
	ZoomPercent            int    // Zoom percentage
	ZoomType               string // Zoom type
	DefaultView            string // Default view mode
	ShowGridlines          bool   // Show gridlines
	ShowRuler              bool   // Show ruler
	ShowParagraphMarks     bool   // Show paragraph marks
	ShowEnvelope           bool   // Show envelope
	DisplayBackgroundShape bool   // Display background
	SavePreviewPicture     bool   // Save preview picture
	PrintTwoOnOne          bool   // Print two pages on one
}

// ============================================
// COMPATIBILITY SETTINGS
// ============================================
type CompatibilitySettings struct {
	CompatibilityMode          int    // Word compatibility mode
	URI                        string // Compatibility URI
	DoNotPrompt                bool   // Don't prompt for compatibility
	AlignBordersAndEdges       bool   // Align borders and edges
	BordersDoNotSurroundHeader bool   // Borders don't surround header
	BordersDoNotSurroundFooter bool   // Borders don't surround footer
	AdjustLineHeightInTable    bool   // Adjust line height in tables
	DoNotBreakWrappedTables    bool   // Don't break wrapped tables
}

// ============================================
// PROTECTION SETTINGS
// ============================================
type ProtectionSettings struct {
	EnforceProtection bool   // Enable protection
	ProtectionType    string // Protection type
	Password          string // Password hash
	Salt              string // Password salt
	AllowEditingMode  string // Editing restrictions
}

// ============================================
// TRACKING SETTINGS
// ============================================
type TrackingSettings struct {
	TrackRevisions       bool // Track changes
	TrackFormatting      bool // Track formatting changes
	DoNotTrackMoves      bool // Don't track moves
	DoNotTrackFormatting bool // Don't track formatting
	ShowInsertions       bool // Show insertions
	ShowDeletions        bool // Show deletions
	ShowFormatting       bool // Show formatting changes
	ShowComments         bool // Show comments
}

// ============================================
// PRINT SETTINGS
// ============================================
type PrintSettings struct {
	PrintPostScriptOverText bool // Print PostScript over text
	PrintFractionalWidth    bool // Print fractional widths
	PrintFormsData          bool // Print only form data
	SaveFormsData           bool // Save form data
	EmbedSystemFonts        bool // Embed system fonts
}

// ============================================
// GRID SETTINGS
// ============================================
type GridSettings struct {
	DrawingGridHorizontalSpacing int  // Horizontal grid spacing
	DrawingGridVerticalSpacing   int  // Vertical grid spacing
	DisplayHorizontalDrawingGrid int  // Display horizontal grid
	DisplayVerticalDrawingGrid   int  // Display vertical grid
	UseMarginsForDrawingGrid     bool // Use margins for grid
}

// ============================================
// WEB SETTINGS
// ============================================
type WebSettings struct {
	WebLayoutView                    bool // Web layout view
	OptimizeForBrowser               bool // Optimize for browser
	DoNotUseHTMLParagraphAutoSpacing bool // Don't use HTML auto spacing
	DoNotAutofit                     bool // Don't autofit
	DoNotValidateAgainstSchema       bool // Don't validate schema
	SaveInvalidXML                   bool // Save invalid XML
	IgnoreMixedContent               bool // Ignore mixed content
	AlwaysShowPlaceholderText        bool // Always show placeholder
	DoNotDemarcateInvalidXML         bool // Don't demarcate invalid XML
}

// ============================================
// NOTES SETTINGS (Footnotes/Endnotes)
// ============================================
type NotesSettings struct {
	Footnotes *FootnoteSettings
	Endnotes  *EndnoteSettings
}

type FootnoteSettings struct {
	Position     string // Position of footnotes
	NumberFormat string // Number format
	StartNumber  int    // Starting number
	RestartRule  string // Restart numbering rule
}

type EndnoteSettings struct {
	Position     string // Position of endnotes
	NumberFormat string // Number format
	StartNumber  int    // Starting number
	RestartRule  string // Restart numbering rule
}

// ============================================
// LINE NUMBERING SETTINGS
// ============================================
type LineNumberingSettings struct {
	CountBy     int    // Count by increment
	Start       int    // Starting number
	Distance    int    // Distance from text
	RestartRule string // Restart rule
}

// ============================================
// BORDER SETTINGS
// ============================================
type BorderSettings struct {
	PageBorders          *PageBorderSettings
	SurroundHeader       bool // Surround header with borders
	SurroundFooter       bool // Surround footer with borders
	AlignBordersAndEdges bool // Align borders and edges
}

type PageBorderSettings struct {
	Display    string // Display borders
	OffsetFrom string // Offset from page or text
}

// ============================================
// MAIL MERGE SETTINGS
// ============================================
type MailMergeSettings struct {
	MainDocumentType string // Main document type
	DataType         string // Data source type
	ConnectString    string // Connection string
	Query            string // Data query
	DataSource       string // Data source path
}

// ============================================
// STYLE PANE SETTINGS
// ============================================
type StylePaneSettings struct {
	FormatFilter string // Format filter
	SortMethod   string // Sort method
}

// ============================================
// PROOFING SETTINGS
// ============================================
type ProofingSettings struct {
	SpellingState             string // Spelling state
	GrammarState              string // Grammar state
	HideSpellingErrors        bool   // Hide spelling errors
	HideGrammaticalErrors     bool   // Hide grammar errors
	DoNotPromptForConvert     bool   // Don't prompt for convert
	DoNotAutoCompressPictures bool   // Don't auto compress pictures
}

// ============================================
// DEFAULT SETTINGS FACTORY
// ============================================
func DefaultSettings() *DocumentSettings {
	return &DocumentSettings{
		Page: &PageSettings{
			Width:         12240, // 8.5 inches
			Height:        15840, // 11 inches
			Orientation:   "portrait",
			MarginTop:     1440, // 1 inch
			MarginRight:   1440,
			MarginBottom:  1440,
			MarginLeft:    1440,
			MarginHeader:  708, // 0.5 inch
			MarginFooter:  708,
			MarginGutter:  0,
			MirrorMargins: false,
			GutterAtTop:   false,
		},
		Font: &FontSettings{
			DefaultFont:             "Times New Roman",
			DefaultSize:             24, // 12pt
			DefaultColor:            "000000",
			EastAsianFont:           "Times New Roman",
			ComplexScriptFont:       "Times New Roman",
			DefaultTabStop:          708, // 0.5 inch
			CharacterSpacingControl: "doNotCompress",
			EmbedTrueTypeFonts:      false,
			SaveSubsetFonts:         true,
		},
		Language: &LanguageSettings{
			Default:                 "en-US",
			EastAsian:               "en-US",
			BiDirectional:           "en-US",
			StrictFirstAndLastChars: false,
		},
		View: &ViewSettings{
			ZoomPercent:            100,
			ZoomType:               "none",
			DefaultView:            "print",
			ShowGridlines:          false,
			ShowRuler:              false,
			ShowParagraphMarks:     false,
			ShowEnvelope:           false,
			DisplayBackgroundShape: true,
			SavePreviewPicture:     false,
			PrintTwoOnOne:          false,
		},
		Compatibility: &CompatibilitySettings{
			CompatibilityMode:          15, // Word 2013
			URI:                        "http://schemas.microsoft.com/office/word",
			DoNotPrompt:                false,
			AlignBordersAndEdges:       false,
			BordersDoNotSurroundHeader: false,
			BordersDoNotSurroundFooter: false,
			AdjustLineHeightInTable:    false,
			DoNotBreakWrappedTables:    false,
		},
		Protection: &ProtectionSettings{
			EnforceProtection: false,
			ProtectionType:    "readOnly",
			AllowEditingMode:  "unrestricted",
		},
		Tracking: &TrackingSettings{
			TrackRevisions:       false,
			TrackFormatting:      false,
			DoNotTrackMoves:      false,
			DoNotTrackFormatting: false,
			ShowInsertions:       true,
			ShowDeletions:        true,
			ShowFormatting:       true,
			ShowComments:         true,
		},
		Print: &PrintSettings{
			PrintPostScriptOverText: true,
			PrintFractionalWidth:    true,
			PrintFormsData:          false,
			SaveFormsData:           false,
			EmbedSystemFonts:        false,
		},
		Grid: &GridSettings{
			DrawingGridHorizontalSpacing: 180,
			DrawingGridVerticalSpacing:   180,
			DisplayHorizontalDrawingGrid: 0,
			DisplayVerticalDrawingGrid:   0,
			UseMarginsForDrawingGrid:     false,
		},
		Web: &WebSettings{
			WebLayoutView:                    false,
			OptimizeForBrowser:               false,
			DoNotUseHTMLParagraphAutoSpacing: false,
			DoNotAutofit:                     false,
			DoNotValidateAgainstSchema:       false,
			SaveInvalidXML:                   false,
			IgnoreMixedContent:               false,
			AlwaysShowPlaceholderText:        false,
			DoNotDemarcateInvalidXML:         false,
		},
		Notes: &NotesSettings{
			Footnotes: &FootnoteSettings{
				Position:     "pageBottom",
				NumberFormat: "decimal",
				StartNumber:  1,
				RestartRule:  "continuous",
			},
			Endnotes: &EndnoteSettings{
				Position:     "sectEnd",
				NumberFormat: "lowerRoman",
				StartNumber:  1,
				RestartRule:  "eachSect",
			},
		},
		Borders: &BorderSettings{
			SurroundHeader:       false,
			SurroundFooter:       false,
			AlignBordersAndEdges: false,
		},
		StylePane: &StylePaneSettings{
			FormatFilter: "3F01",
			SortMethod:   "name",
		},
		Proofing: &ProofingSettings{
			SpellingState:             "clean",
			GrammarState:              "clean",
			HideSpellingErrors:        false,
			HideGrammaticalErrors:     false,
			DoNotPromptForConvert:     false,
			DoNotAutoCompressPictures: false,
		},
		Variables: make(map[string]string),
	}
}

// ============================================
// PRESET SETTINGS
// ============================================

// A4Settings returns A4 page settings
func A4Settings() *DocumentSettings {
	settings := DefaultSettings()
	settings.Page.Width = 11906  // 210mm
	settings.Page.Height = 16838 // 297mm
	return settings
}

// LandscapeSettings returns landscape settings
func LandscapeSettings() *DocumentSettings {
	settings := DefaultSettings()
	settings.Page.Orientation = "landscape"
	settings.Page.Width, settings.Page.Height = settings.Page.Height, settings.Page.Width
	return settings
}

// WebDocumentSettings returns settings optimized for web
func WebDocumentSettings() *DocumentSettings {
	settings := DefaultSettings()
	settings.Web.WebLayoutView = true
	settings.Web.OptimizeForBrowser = true
	settings.View.DefaultView = "web"
	return settings
}

// SecureDocumentSettings returns settings with protection
func SecureDocumentSettings() *DocumentSettings {
	settings := DefaultSettings()
	settings.Protection.EnforceProtection = true
	settings.Protection.ProtectionType = "readOnly"
	settings.Tracking.TrackRevisions = true
	return settings
}

// ============================================
// CONVENIENCE METHODS
// ============================================

// SetA4Page sets A4 page dimensions
func (ds *DocumentSettings) SetA4Page() {
	if ds.Page == nil {
		ds.Page = &PageSettings{}
	}
	ds.Page.Width = 11906
	ds.Page.Height = 16838
}

// SetLandscape sets landscape orientation
func (ds *DocumentSettings) SetLandscape() {
	if ds.Page == nil {
		ds.Page = DefaultSettings().Page
	}
	ds.Page.Orientation = "landscape"
	ds.Page.Width, ds.Page.Height = ds.Page.Height, ds.Page.Width
}

// SetMargins sets all margins
func (ds *DocumentSettings) SetMargins(top, right, bottom, left int) {
	if ds.Page == nil {
		ds.Page = &PageSettings{}
	}
	ds.Page.MarginTop = top
	ds.Page.MarginRight = right
	ds.Page.MarginBottom = bottom
	ds.Page.MarginLeft = left
}

// EnableTrackChanges enables revision tracking
func (ds *DocumentSettings) EnableTrackChanges() {
	if ds.Tracking == nil {
		ds.Tracking = &TrackingSettings{}
	}
	ds.Tracking.TrackRevisions = true
	ds.Tracking.TrackFormatting = true
}

// SetDefaultFont sets the default font
func (ds *DocumentSettings) SetDefaultFont(fontName string, fontSize int) {
	if ds.Font == nil {
		ds.Font = &FontSettings{}
	}
	ds.Font.DefaultFont = fontName
	ds.Font.DefaultSize = fontSize * 2 // Convert to half-points
}
