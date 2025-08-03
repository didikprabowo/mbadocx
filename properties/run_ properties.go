// File: properties/run_properties.go
package properties

import (
	"fmt"
)

// RunProperties defines text formatting properties
type RunProperties struct {
	// Basic formatting
	Bold         *bool  // Bold text
	Italic       *bool  // Italic text
	Underline    string // Underline style: none, single, double, thick, dotted, dash, dotDash, dotDotDash, wave
	Strike       *bool  // Strikethrough
	DoubleStrike *bool  // Double strikethrough

	// Font properties
	FontSize   float64 // Font size in points
	FontFamily string  // Font family name

	// Colors
	Color     string // Text color in RGB hex format (e.g., "FF0000" for red)
	Highlight string // Highlight color: yellow, green, cyan, magenta, blue, red, darkBlue, darkCyan, darkGreen, darkMagenta, darkRed, darkYellow, darkGray, lightGray, black

	// Text effects
	VerticalAlign string // Vertical alignment: baseline, superscript, subscript
	AllCaps       *bool  // All capitals
	SmallCaps     *bool  // Small capitals
	Outline       *bool  // Outline effect
	Shadow        *bool  // Shadow effect
	Emboss        *bool  // Emboss effect
	Imprint       *bool  // Imprint/engrave effect
	Vanish        *bool  // Hidden/vanish text

	// Spacing and positioning
	Spacing  int     // Character spacing in twips (1/20th of a point)
	Kerning  float64 // Kerning in points (minimum font size for kerning)
	Position int     // Text position (raise/lower) in half-points

	// Style reference
	StyleID string // Character style ID

	// Language and locale
	Language string // Language code (e.g., "en-US")

	// Advanced properties
	NoProof       *bool // Disable spell/grammar checking
	SnapToGrid    *bool // Snap to document grid
	Hidden        *bool // Hidden text (same as Vanish)
	WebHidden     *bool // Hidden in web view
	SpecVanish    *bool // Special vanish
	RightToLeft   *bool // Right-to-left text
	ComplexScript *bool // Complex script formatting

	// Border
	Border *RunBorder // Text border

	// Shading
	Shading *RunShading // Text shading/background

	// Fit text
	FitText *int // Fit text width in twips

	// Animation (legacy)
	Animation string // Text animation effect (legacy Word feature)
}

// RunBorder defines text border properties
type RunBorder struct {
	Type   string // Border type: single, double, triple, etc.
	Width  int    // Border width in eighths of a point
	Color  string // Border color in hex
	Space  int    // Space between border and text in points
	Shadow bool   // Shadow effect
}

// RunShading defines text shading properties
type RunShading struct {
	Fill         string // Fill color in hex
	Color        string // Foreground color in hex
	Pattern      string // Pattern type: clear, solid, horzStripe, vertStripe, etc.
	PatternColor string // Pattern color in hex
}

// NewRunProperties creates new run properties with defaults
func NewRunProperties() *RunProperties {
	return &RunProperties{
		FontSize:      11,        // Default 11pt
		FontFamily:    "Calibri", // Default font
		Underline:     "",        // No underline by default
		Color:         "",        // Default color (black)
		Highlight:     "",        // No highlight
		VerticalAlign: "",        // Baseline by default
		Language:      "en-US",   // Default language
	}
}

// Clone creates a deep copy of RunProperties
func (rp *RunProperties) Clone() *RunProperties {
	if rp == nil {
		return nil
	}

	clone := &RunProperties{
		Underline:     rp.Underline,
		FontSize:      rp.FontSize,
		FontFamily:    rp.FontFamily,
		Color:         rp.Color,
		Highlight:     rp.Highlight,
		VerticalAlign: rp.VerticalAlign,
		Spacing:       rp.Spacing,
		Kerning:       rp.Kerning,
		Position:      rp.Position,
		StyleID:       rp.StyleID,
		Language:      rp.Language,
		Animation:     rp.Animation,
	}

	// Clone pointer fields
	if rp.Bold != nil {
		b := *rp.Bold
		clone.Bold = &b
	}
	if rp.Italic != nil {
		i := *rp.Italic
		clone.Italic = &i
	}
	if rp.Strike != nil {
		s := *rp.Strike
		clone.Strike = &s
	}
	if rp.DoubleStrike != nil {
		ds := *rp.DoubleStrike
		clone.DoubleStrike = &ds
	}
	if rp.AllCaps != nil {
		ac := *rp.AllCaps
		clone.AllCaps = &ac
	}
	if rp.SmallCaps != nil {
		sc := *rp.SmallCaps
		clone.SmallCaps = &sc
	}
	if rp.Outline != nil {
		o := *rp.Outline
		clone.Outline = &o
	}
	if rp.Shadow != nil {
		s := *rp.Shadow
		clone.Shadow = &s
	}
	if rp.Emboss != nil {
		e := *rp.Emboss
		clone.Emboss = &e
	}
	if rp.Imprint != nil {
		i := *rp.Imprint
		clone.Imprint = &i
	}
	if rp.Vanish != nil {
		v := *rp.Vanish
		clone.Vanish = &v
	}
	if rp.NoProof != nil {
		np := *rp.NoProof
		clone.NoProof = &np
	}
	if rp.SnapToGrid != nil {
		stg := *rp.SnapToGrid
		clone.SnapToGrid = &stg
	}
	if rp.Hidden != nil {
		h := *rp.Hidden
		clone.Hidden = &h
	}
	if rp.WebHidden != nil {
		wh := *rp.WebHidden
		clone.WebHidden = &wh
	}
	if rp.SpecVanish != nil {
		sv := *rp.SpecVanish
		clone.SpecVanish = &sv
	}
	if rp.RightToLeft != nil {
		rtl := *rp.RightToLeft
		clone.RightToLeft = &rtl
	}
	if rp.ComplexScript != nil {
		cs := *rp.ComplexScript
		clone.ComplexScript = &cs
	}
	if rp.FitText != nil {
		ft := *rp.FitText
		clone.FitText = &ft
	}

	// Clone complex objects
	if rp.Border != nil {
		clone.Border = &RunBorder{
			Type:   rp.Border.Type,
			Width:  rp.Border.Width,
			Color:  rp.Border.Color,
			Space:  rp.Border.Space,
			Shadow: rp.Border.Shadow,
		}
	}

	if rp.Shading != nil {
		clone.Shading = &RunShading{
			Fill:         rp.Shading.Fill,
			Color:        rp.Shading.Color,
			Pattern:      rp.Shading.Pattern,
			PatternColor: rp.Shading.PatternColor,
		}
	}

	return clone
}

// Merge merges another RunProperties into this one
// Other properties take precedence where they are set
func (rp *RunProperties) Merge(other *RunProperties) {
	if other == nil {
		return
	}

	// Merge basic formatting
	if other.Bold != nil {
		rp.Bold = other.Bold
	}
	if other.Italic != nil {
		rp.Italic = other.Italic
	}
	if other.Underline != "" {
		rp.Underline = other.Underline
	}
	if other.Strike != nil {
		rp.Strike = other.Strike
	}
	if other.DoubleStrike != nil {
		rp.DoubleStrike = other.DoubleStrike
	}

	// Merge font properties
	if other.FontSize > 0 {
		rp.FontSize = other.FontSize
	}
	if other.FontFamily != "" {
		rp.FontFamily = other.FontFamily
	}

	// Merge colors
	if other.Color != "" {
		rp.Color = other.Color
	}
	if other.Highlight != "" {
		rp.Highlight = other.Highlight
	}

	// Merge text effects
	if other.VerticalAlign != "" {
		rp.VerticalAlign = other.VerticalAlign
	}
	if other.AllCaps != nil {
		rp.AllCaps = other.AllCaps
	}
	if other.SmallCaps != nil {
		rp.SmallCaps = other.SmallCaps
	}
	if other.Outline != nil {
		rp.Outline = other.Outline
	}
	if other.Shadow != nil {
		rp.Shadow = other.Shadow
	}
	if other.Emboss != nil {
		rp.Emboss = other.Emboss
	}
	if other.Imprint != nil {
		rp.Imprint = other.Imprint
	}
	if other.Vanish != nil {
		rp.Vanish = other.Vanish
	}

	// Merge spacing
	if other.Spacing != 0 {
		rp.Spacing = other.Spacing
	}
	if other.Kerning > 0 {
		rp.Kerning = other.Kerning
	}
	if other.Position != 0 {
		rp.Position = other.Position
	}

	// Merge other properties
	if other.StyleID != "" {
		rp.StyleID = other.StyleID
	}
	if other.Language != "" {
		rp.Language = other.Language
	}
	if other.Border != nil {
		rp.Border = other.Border
	}
	if other.Shading != nil {
		rp.Shading = other.Shading
	}
}

// Reset clears all formatting
func (rp *RunProperties) Reset() {
	*rp = *NewRunProperties()
}

// Validate validates the run properties
func (rp *RunProperties) Validate() error {
	// Validate underline values
	validUnderlines := map[string]bool{
		"":           true,
		"none":       true,
		"single":     true,
		"double":     true,
		"thick":      true,
		"dotted":     true,
		"dash":       true,
		"dotDash":    true,
		"dotDotDash": true,
		"wave":       true,
		"wavyHeavy":  true,
		"wavyDouble": true,
		"words":      true, // Underline words only
	}
	if !validUnderlines[rp.Underline] {
		return fmt.Errorf("invalid underline value: %s", rp.Underline)
	}

	// Validate highlight colors
	validHighlights := map[string]bool{
		"":            true,
		"none":        true,
		"black":       true,
		"blue":        true,
		"cyan":        true,
		"darkBlue":    true,
		"darkCyan":    true,
		"darkGray":    true,
		"darkGreen":   true,
		"darkMagenta": true,
		"darkRed":     true,
		"darkYellow":  true,
		"green":       true,
		"lightGray":   true,
		"magenta":     true,
		"red":         true,
		"white":       true,
		"yellow":      true,
	}
	if !validHighlights[rp.Highlight] {
		return fmt.Errorf("invalid highlight color: %s", rp.Highlight)
	}

	// Validate vertical alignment
	validVertAlign := map[string]bool{
		"":            true,
		"baseline":    true,
		"superscript": true,
		"subscript":   true,
	}
	if !validVertAlign[rp.VerticalAlign] {
		return fmt.Errorf("invalid vertical alignment: %s", rp.VerticalAlign)
	}

	// Validate font size
	if rp.FontSize < 0 {
		return fmt.Errorf("font size cannot be negative: %f", rp.FontSize)
	}

	// Validate kerning
	if rp.Kerning < 0 {
		return fmt.Errorf("kerning cannot be negative: %f", rp.Kerning)
	}

	// Validate border
	if rp.Border != nil {
		validBorderTypes := map[string]bool{
			"single":     true,
			"double":     true,
			"triple":     true,
			"thick":      true,
			"hairline":   true,
			"dot":        true,
			"dash":       true,
			"dotDash":    true,
			"dashDotDot": true,
			"wave":       true,
			"doubleWave": true,
			"none":       true,
		}
		if !validBorderTypes[rp.Border.Type] {
			return fmt.Errorf("invalid border type: %s", rp.Border.Type)
		}

		if rp.Border.Width < 0 {
			return fmt.Errorf("border width cannot be negative: %d", rp.Border.Width)
		}
	}

	// Validate shading pattern
	if rp.Shading != nil && rp.Shading.Pattern != "" {
		validPatterns := map[string]bool{
			"clear":             true,
			"solid":             true,
			"horzStripe":        true,
			"vertStripe":        true,
			"reverseDiagStripe": true,
			"diagStripe":        true,
			"horzCross":         true,
			"diagCross":         true,
			"thinHorzStripe":    true,
			"thinVertStripe":    true,
			"pct5":              true,
			"pct10":             true,
			"pct20":             true,
			"pct25":             true,
			"pct30":             true,
			"pct40":             true,
			"pct50":             true,
			"pct60":             true,
			"pct70":             true,
			"pct75":             true,
			"pct80":             true,
			"pct90":             true,
		}
		if !validPatterns[rp.Shading.Pattern] {
			return fmt.Errorf("invalid shading pattern: %s", rp.Shading.Pattern)
		}
	}

	return nil
}

// IsEmpty returns true if no formatting is applied
func (rp *RunProperties) IsEmpty() bool {
	if rp == nil {
		return true
	}

	return rp.Bold == nil &&
		rp.Italic == nil &&
		rp.Underline == "" &&
		rp.Strike == nil &&
		rp.DoubleStrike == nil &&
		rp.FontSize == 0 &&
		rp.FontFamily == "" &&
		rp.Color == "" &&
		rp.Highlight == "" &&
		rp.VerticalAlign == "" &&
		rp.AllCaps == nil &&
		rp.SmallCaps == nil &&
		rp.Outline == nil &&
		rp.Shadow == nil &&
		rp.Emboss == nil &&
		rp.Imprint == nil &&
		rp.Vanish == nil &&
		rp.Spacing == 0 &&
		rp.Kerning == 0 &&
		rp.Position == 0 &&
		rp.StyleID == "" &&
		rp.Border == nil &&
		rp.Shading == nil
}

// HasEffect returns true if any text effect is applied
func (rp *RunProperties) HasEffect() bool {
	if rp == nil {
		return false
	}

	return (rp.AllCaps != nil && *rp.AllCaps) ||
		(rp.SmallCaps != nil && *rp.SmallCaps) ||
		(rp.Outline != nil && *rp.Outline) ||
		(rp.Shadow != nil && *rp.Shadow) ||
		(rp.Emboss != nil && *rp.Emboss) ||
		(rp.Imprint != nil && *rp.Imprint) ||
		(rp.Vanish != nil && *rp.Vanish) ||
		rp.Animation != ""
}
